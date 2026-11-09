import re
import json
import time
import base64
import requests
import pandas as pd
import os
from pathlib import Path
from itertools import combinations
from collections import Counter
from io import BytesIO
from PIL import Image

# ================= 1. 配置区 =================
API_KEY = os.environ.get("DASHSCOPE_API_KEY", "")

PROJECT_DIR = Path(__file__).resolve().parent
INPUT_EXCEL = Path(os.environ.get("WORKFLOW_INPUT", PROJECT_DIR / "输入表.xlsx"))
FRAME_BASE_DIR = Path(os.environ.get("FRAME_BASE_DIR", PROJECT_DIR / "frames"))
OUTPUT_EXCEL = Path(os.environ.get("FINAL_OUTPUT", PROJECT_DIR / "最终编码结果.xlsx"))

API_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
TEXT_MODEL = "qwen3.8-flash"
VISION_MODEL = "qwen3.8-flash"

PLATFORM_KEYWORDS = r"(官方企划|平台活动|独家首播|联合出品|自制综艺|B站官方|百大UP主|召集计划|官方账号|官方活动)"

SUCCESS_MODES = {"1", "2", "3"}
DONE_9_STATUS = "done_9"              # 模型明确返回 9，不重跑
FAILED_STATUS = "failed_retryable"    # 网络/API/本地无图等，可重跑
OK_STATUS = "ok"


# ================= 2. 通用工具 =================
def norm_key(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def norm_mode(x):
    if pd.isna(x):
        return "9"
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


# ================= 3. 城市结构变量计算 =================
def calc_city_vars(participant_cities_str):
    if pd.isna(participant_cities_str) or not str(participant_cities_str).strip():
        return "C9", 0.0, 0, 0.0, 0
    cities = [c.strip() for c in str(participant_cities_str).replace("，", ",").split(",") if c.strip()]
    k = len(cities)
    if k == 0:
        return "C9", 0.0, 0, 0.0, 0
    n_city = len(set(cities))
    city_counts = Counter(cities)
    max_city_share = max(city_counts.values()) / k
    dyad_share = 1.0 if k == 1 else sum(1 for a, b in combinations(cities, 2) if a == b) / len(
        list(combinations(cities, 2)))
    if n_city == 1:
        config = "C1"
    elif n_city > 1 and max_city_share >= 0.60:
        config = "C2"
    elif n_city > 1 and max_city_share < 0.60:
        config = "C3"
    else:
        config = "C9"
    return config, round(dyad_share, 4), n_city, round(max_city_share, 4), k


# ================= 4. 安全保存 =================
def safe_save(dataframe, excel_path):
    csv_path = excel_path.with_suffix('.csv')
    try:
        dataframe.to_csv(csv_path, index=False, encoding='utf-8-sig')
    except Exception as e:
        print(f"\n⚠️ CSV 保存失败: {e}")
    try:
        dataframe.to_excel(excel_path, index=False)
    except PermissionError:
        print(f"\n⚠️ Excel 被占用，进度已存入 {csv_path.name}（不影响续跑）。")
    except Exception as e:
        print(f"\n⚠️ Excel 保存失败: {e}（进度已存入 CSV）。")


# ================= 5. 加载历史进度（区分 done_9 / failed_retryable） =================
def load_progress():
    csv_path = OUTPUT_EXCEL.with_suffix('.csv')
    existing = None

    if csv_path.exists():
        try:
            existing = pd.read_csv(csv_path)
            print(f"📂 从 CSV 加载历史：{len(existing)} 条")
        except Exception as e:
            print(f"CSV 读取失败: {e}")
    if existing is None and OUTPUT_EXCEL.exists():
        try:
            existing = pd.read_excel(OUTPUT_EXCEL)
            print(f"📂 从 XLSX 加载历史：{len(existing)} 条")
        except Exception as e:
            print(f"XLSX 读取失败: {e}")

    if existing is None or existing.empty:
        return [], set()

    if "序号(number)" not in existing.columns:
        print("⚠️ 历史文件缺少 '序号(number)' 列，忽略历史。")
        return [], set()

    existing["_num"] = existing["序号(number)"].apply(norm_key)
    existing["_mode"] = existing["空间协作方式(spatial_mode)"].apply(norm_mode)

    # 兼容旧文件：没有 status 列时，按 spatial_mode 推断
    if "status" not in existing.columns:
        # 旧逻辑：1/2/3 = ok；9 一律视为 done_9（保守，不重跑，避免浪费额度）
        existing["status"] = existing["_mode"].apply(
            lambda m: OK_STATUS if m in SUCCESS_MODES else DONE_9_STATUS
        )
        print("ℹ️ 历史文件无 status 列，已按 spatial_mode 推断（9 视为 done_9 不重跑）。")

    # 只有 ok / done_9 视为已处理，不重跑；failed_retryable 会重跑
    done_mask = existing["status"].isin([OK_STATUS, DONE_9_STATUS])
    processed_numbers = set(existing.loc[done_mask, "_num"].tolist())

    # 保留全部历史（含失败），失败条目重跑后覆盖
    results = existing.drop(columns=["_num", "_mode"]).to_dict("records")

    n_ok = (existing["status"] == OK_STATUS).sum()
    n_9 = (existing["status"] == DONE_9_STATUS).sum()
    n_fail = (existing["status"] == FAILED_STATUS).sum()
    print(f"✅ 已处理不重跑：ok={n_ok}, done_9={n_9}；"
          f"待重跑(failed_retryable)={n_fail}。")
    return results, processed_numbers


# ================= 6. 图片压缩与请求辅助 =================
def compress_image_base64(img_path, max_width=512):
    try:
        with Image.open(img_path) as img:
            if img.width > max_width:
                ratio = max_width / img.width
                new_size = (max_width, int(img.height * ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
            buffered = BytesIO()
            img.convert("RGB").save(buffered, format="JPEG", quality=70)
            return base64.b64encode(buffered.getvalue()).decode("utf-8")
    except Exception:
        return None


def send_api_request(data):
    if not API_KEY:
        raise RuntimeError("未设置 DASHSCOPE_API_KEY 环境变量")
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json; charset=utf-8"
    }
    payload = json.dumps(data, ensure_ascii=False).encode('utf-8')
    return requests.post(API_URL, headers=headers, data=payload, timeout=180).json()


# ================= 7. 大模型调用（返回内部 _fail_reason 用于判定 status） =================
def call_llm_text(title, desc, tags):
    prompt = f"""请阅读以下 B 站视频元数据，判断是否包含 B 站官方活动或企划标识。
标题：{title} 简介：{desc} 标签：{tags}
判定标准：platform_marker 取值 1（有明确官方标识）、0（无）、9（资料不足）。
仅返回纯 JSON：{{"platform_marker": 1, "reason": "简述原因"}}"""
    data = {"model": TEXT_MODEL, "messages": [{"role": "user", "content": prompt}], "temperature": 0.1}

    last_err = ""
    for attempt in range(3):
        try:
            resp = send_api_request(data)
            if "choices" in resp:
                res_text = resp["choices"][0]["message"]["content"]
                match = re.search(r'\{[\s\S]*\}', res_text)
                if match:
                    out = json.loads(match.group())
                    out["_fail_reason"] = ""   # 成功解析
                    return out
                last_err = "文本AI无JSON"
            else:
                last_err = f"文本API异常:{resp.get('error', {}).get('message', str(resp))}"
        except Exception as e:
            last_err = f"文本网络异常:{e}"
        time.sleep(2)
    return {"platform_marker": 9, "reason": "文本AI失败", "_fail_reason": last_err or "文本AI失败"}


def call_llm_vision(image_paths, number):
    if not image_paths:
        # 本地无图片：属于可重跑（可能抽帧没做/路径错），标记 failed_retryable
        return {"spatial_mode": "9", "physical_space": "9",
                "evidence": "本地无图片", "_fail_reason": "本地无图片"}

    prompt_text = """请观察这些来自同一视频的截图（可能不足8张），完成以下任务：
1. spatial_mode：综合判断空间协作方式。1=线下共在（同一场景内有多人互动）；2=远程协作（视频连线/分屏/不同背景）；3=混合；9=无法判断。
2. frame_spaces：逐图判断每一张图的场景类型。请用数组按顺序返回每张图的判断结果。判断标准：
   - 1 = 室内（特征：卧室、宿舍、家中客厅、办公室、教室等封闭空间）。
   - 2 = 室外（特征：户外街道、商场、公园、餐厅等开放或半开放公共空间）。
   - 9 = 无法判断。
3. evidence：一句话简述整体判断依据。

严格返回纯 JSON：{"spatial_mode": 1, "frame_spaces": [1, 1, 2, 2, 2, 1, 2, 2], "evidence": "多人同框在室内"}
"""
    content = [{"type": "text", "text": prompt_text}]

    for img_path in image_paths:
        b64 = compress_image_base64(img_path)
        if b64:
            content.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}})

    data = {"model": VISION_MODEL, "messages": [{"role": "user", "content": content}], "temperature": 0.1}

    last_err = ""
    for attempt in range(1, 4):
        try:
            resp = send_api_request(data)

            if "choices" not in resp:
                err_msg = resp.get('error', {}).get('message', str(resp))
                print(f"[API异常:{err_msg}] ", end="")
                last_err = f"视觉API异常:{err_msg}"
                time.sleep(5)
                continue

            res_text = resp["choices"][0]["message"]["content"]
            match = re.search(r'\{[\s\S]*\}', res_text)
            if match:
                res_json = json.loads(match.group())

                frame_spaces = res_json.get("frame_spaces", [])
                outdoor_count = frame_spaces.count(2)

                if outdoor_count >= 3:
                    final_physical_space = 2
                elif len(frame_spaces) > 0:
                    final_physical_space = 1
                else:
                    final_physical_space = 9

                return {
                    "spatial_mode": res_json.get("spatial_mode", "9"),
                    "physical_space": final_physical_space,
                    "evidence": f"{res_json.get('evidence', '')} (逐图判断:{frame_spaces}, 室外图达{outdoor_count}张)",
                    "_fail_reason": ""   # 成功拿到模型结果
                }
            last_err = "视觉AI无JSON"
        except Exception as e:
            print(f"[网络异常,等待重试] ", end="")
            last_err = f"视觉网络异常:{e}"
            time.sleep(5)

    return {"spatial_mode": "9", "physical_space": "9",
            "evidence": "重试耗尽", "_fail_reason": last_err or "重试耗尽"}


# ================= 8. 判定单条 status =================
def decide_status(spatial_mode, vision_fail_reason, text_fail_reason, row_exception=None):
    """
    ok               -> spatial_mode ∈ {1,2,3}
    failed_retryable -> 有网络/API/本地无图/单条异常
    done_9           -> 模型明确返回 9（无失败原因）
    """
    if row_exception:
        return FAILED_STATUS
    if norm_mode(spatial_mode) in SUCCESS_MODES:
        return OK_STATUS
    if vision_fail_reason or text_fail_reason:
        return FAILED_STATUS
    return DONE_9_STATUS


# ================= 9. 主程序 =================
def main():
    df = pd.read_excel(INPUT_EXCEL)

    results, processed_numbers = load_progress()

    # 用 bvid 做辅助键
    processed_bvids = set()
    for r in results:
        if str(r.get("status", "")) in (OK_STATUS, DONE_9_STATUS):
            processed_bvids.add(norm_key(r.get("视频编号(bvid)", "")))

    results_by_key = {}
    for r in results:
        results_by_key[norm_key(r.get("序号(number)", ""))] = r

    total = len(df)
    new_ok = new_9 = new_fail = 0
    print(f"开始处理，共 {total} 条输入。\n")

    try:
        for idx, row in df.iterrows():
            number = norm_key(row.get("number", ""))
            bvid = norm_key(row.get("bvid", ""))

            if number in processed_numbers:
                continue
            if bvid and bvid in processed_bvids:
                continue

            title = str(row.get("title", ""))
            desc = str(row.get("description", ""))
            tags = str(row.get("tags", ""))

            print(f"[{idx + 1}/{total}] {number} ...", end=" ", flush=True)

            try:
                city_config, dyad_share, n_city, max_share, k = calc_city_vars(row.get("participant_cities", ""))

                combined = title + " " + desc + " " + tags
                p_marker, text_reason = 9, ""
                text_fail_reason = ""
                if re.search(PLATFORM_KEYWORDS, combined, re.IGNORECASE):
                    p_marker, text_reason = 1, "命中规则词"
                else:
                    text_res = call_llm_text(title, desc, tags)
                    p_marker = text_res.get("platform_marker", 9)
                    text_reason = text_res.get("reason", "")
                    text_fail_reason = text_res.get("_fail_reason", "")

                video_dir = FRAME_BASE_DIR / number
                image_paths = sorted(video_dir.glob("*.jpg")) if video_dir.exists() else []
                vision_res = call_llm_vision(image_paths, number)
                vision_fail_reason = vision_res.get("_fail_reason", "")

                status = decide_status(
                    vision_res.get("spatial_mode", "9"),
                    vision_fail_reason,
                    text_fail_reason,
                )

                record = {
                    "视频编号(bvid)": bvid, "序号(number)": number, "标题(title)": title,
                    "参与者数量(k)": k, "城市关系构型(city_config)": city_config,
                    "同城二元关系占比(within_city_dyad_share)": dyad_share,
                    "涉及地级市数量(n_city)": n_city, "最大城市参与者占比(max_city_share)": max_share,
                    "空间协作方式(spatial_mode)": vision_res.get("spatial_mode", "9"),
                    "场景物理空间(physical_space)": vision_res.get("physical_space", "9"),
                    "平台官方标识(platform_marker)": p_marker,
                    "编码依据(evidence)": f"文本:{text_reason}; 画面:{vision_res.get('evidence', '')}",
                    "置信度(confidence)": "高" if status == OK_STATUS else ("中" if status == DONE_9_STATUS else "低"),
                    "status": status,
                }

                results_by_key[number] = record

                if status in (OK_STATUS, DONE_9_STATUS):
                    processed_numbers.add(number)
                    if bvid:
                        processed_bvids.add(bvid)
                    if status == OK_STATUS:
                        new_ok += 1
                    else:
                        new_9 += 1
                else:
                    new_fail += 1

                print(f"C={city_config}, 协作={record['空间协作方式(spatial_mode)']}, "
                      f"场景={record['场景物理空间(physical_space)']}, 平台={p_marker}, status={status}")

            except Exception as e:
                print(f"❌ 本条异常: {e}（标记 failed_retryable，下次重跑）")
                new_fail += 1
                results_by_key[number] = {
                    "视频编号(bvid)": bvid, "序号(number)": number, "标题(title)": title,
                    "空间协作方式(spatial_mode)": "9", "场景物理空间(physical_space)": "9",
                    "平台官方标识(platform_marker)": 9,
                    "编码依据(evidence)": f"处理异常: {e}",
                    "置信度(confidence)": "低",
                    "status": FAILED_STATUS,
                }

            results = list(results_by_key.values())
            safe_save(pd.DataFrame(results), OUTPUT_EXCEL)
            time.sleep(2)

    except KeyboardInterrupt:
        print("\n\n⏸ 手动中断，正在保存当前进度...")
    finally:
        results = list(results_by_key.values())
        safe_save(pd.DataFrame(results), OUTPUT_EXCEL)

        in_nums = {norm_key(x) for x in df["number"]}
        out_nums = {norm_key(x) for x in pd.DataFrame(results)["序号(number)"]}
        missing = in_nums - out_nums
        df_out = pd.DataFrame(results)
        n_ok = (df_out["status"] == OK_STATUS).sum() if "status" in df_out.columns else 0
        n_9 = (df_out["status"] == DONE_9_STATUS).sum() if "status" in df_out.columns else 0
        n_fail = (df_out["status"] == FAILED_STATUS).sum() if "status" in df_out.columns else 0

        print(f"\n✅ 本次运行：新增 ok={new_ok}, done_9={new_9}, failed={new_fail}")
        print(f"📊 输出总计：{len(results)} 条 | ok={n_ok}, done_9={n_9}, failed_retryable={n_fail}")
        if missing:
            print(f"⚠️ 仍有 {len(missing)} 条未产出（下次自动重跑）：{sorted(missing)[:30]}")
        else:
            print("✅ 输入所有 number 均已在输出中。")
        if n_fail:
            print(f"🔁 有 {n_fail} 条 failed_retryable，下次运行会自动重跑（不会重跑 ok/done_9）。")
        print(f"结果：{OUTPUT_EXCEL}")


if __name__ == "__main__":
    main()
