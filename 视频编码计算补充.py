import re
import shutil
import pandas as pd
import os
from pathlib import Path
from collections import Counter

# ================= 配置 =================
PROJECT_DIR  = Path(__file__).resolve().parent
INPUT_EXCEL  = Path(os.environ.get("WORKFLOW_INPUT", PROJECT_DIR / "输入表.xlsx"))
FINAL_EXCEL  = Path(os.environ.get("FINAL_OUTPUT", PROJECT_DIR / "最终编码结果.xlsx"))
LOC_SUFFIX   = "_location"

# 最终表里要写入/更新的列名（必须与现有列名完全一致，否则会新增列）
COL_K          = "参与者数量(k)"
COL_CONFIG     = "城市关系构型(city_config)"
COL_DYAD       = "同城二元关系占比(within_city_dyad_share)"
COL_NCITY      = "涉及地级市数量(n_city)"
COL_MAXSHARE   = "最大城市参与者占比(max_city_share)"
JOIN_KEY_FINAL = "序号(number)"   # 最终表里的连接键
JOIN_KEY_SRC   = "number"         # 源表里的连接键

# ---------- 中国大陆地级市名单（293 地级市 + 4 直辖市 + 省直辖县级市） ----------
MAINLAND_CITIES = {
    "北京","上海","天津","重庆",
    "石家庄","唐山","秦皇岛","邯郸","邢台","保定","张家口","承德","沧州","廊坊","衡水",
    "太原","大同","阳泉","长治","晋城","朔州","晋中","运城","忻州","临汾","吕梁",
    "呼和浩特","包头","乌海","赤峰","通辽","鄂尔多斯","呼伦贝尔","巴彦淖尔","乌兰察布",
    "沈阳","大连","鞍山","抚顺","本溪","丹东","锦州","营口","阜新","辽阳","盘锦","铁岭","朝阳","葫芦岛",
    "长春","吉林","四平","辽源","通化","白山","松原","白城",
    "哈尔滨","齐齐哈尔","鸡西","鹤岗","双鸭山","大庆","伊春","佳木斯","七台河","牡丹江","黑河","绥化",
    "南京","无锡","徐州","常州","苏州","南通","连云港","淮安","盐城","扬州","镇江","泰州","宿迁",
    "杭州","宁波","温州","嘉兴","湖州","绍兴","金华","衢州","舟山","台州","丽水",
    "合肥","芜湖","蚌埠","淮南","马鞍山","淮北","铜陵","安庆","黄山","滁州","阜阳","宿州","六安","亳州","池州","宣城",
    "福州","厦门","莆田","三明","泉州","漳州","南平","龙岩","宁德",
    "南昌","景德镇","萍乡","九江","新余","鹰潭","赣州","吉安","宜春","抚州","上饶",
    "济南","青岛","淄博","枣庄","东营","烟台","潍坊","济宁","泰安","威海","日照","临沂","德州","聊城","滨州","菏泽",
    "郑州","开封","洛阳","平顶山","安阳","鹤壁","新乡","焦作","濮阳","许昌","漯河","三门峡","南阳","商丘","信阳","周口","驻马店",
    "武汉","黄石","十堰","宜昌","襄阳","鄂州","荆门","孝感","荆州","黄冈","咸宁","随州",
    "长沙","株洲","湘潭","衡阳","邵阳","岳阳","常德","张家界","益阳","郴州","永州","怀化","娄底",
    "广州","韶关","深圳","珠海","汕头","佛山","江门","湛江","茂名","肇庆","惠州","梅州","汕尾","河源","阳江","清远","东莞","中山","潮州","揭阳","云浮",
    "南宁","柳州","桂林","梧州","北海","防城港","钦州","贵港","玉林","百色","贺州","河池","来宾","崇左",
    "海口","三亚","三沙","儋州",
    "成都","自贡","攀枝花","泸州","德阳","绵阳","广元","遂宁","内江","乐山","南充","眉山","宜宾","广安","达州","雅安","巴中","资阳",
    "贵阳","六盘水","遵义","安顺","毕节","铜仁",
    "昆明","曲靖","玉溪","保山","昭通","丽江","普洱","临沧",
    "拉萨","日喀则","昌都","林芝","山南","那曲",
    "西安","铜川","宝鸡","咸阳","渭南","延安","汉中","榆林","安康","商洛",
    "兰州","嘉峪关","金昌","白银","天水","武威","张掖","平凉","酒泉","庆阳","定西","陇南",
    "西宁","海东",
    "银川","石嘴山","吴忠","固原","中卫",
    "乌鲁木齐","克拉玛依","吐鲁番","哈密",
    "仙桃","潜江","天门","神农架","济源","石河子","阿拉尔","图木舒克","五家渠","北屯","铁门关","双河","可克达拉","昆玉","胡杨河","新星","白杨",
}

OVERSEAS_KEYWORDS = [
    "台湾","香港","澳门","台北","高雄","台中","台南","新北","桃园",
    "美国","韩国","日本","芬兰","法国","西班牙","德国","英国","澳大利亚","加拿大","新加坡",
    "泰国","意大利","荷兰","瑞典","挪威","丹麦","瑞士","新西兰","马来西亚","菲律宾","越南",
    "印度","俄罗斯","巴西","墨西哥","阿根廷","南非","埃及","土耳其","阿联酋","沙特","以色列",
    "波兰","捷克","奥地利","比利时","葡萄牙","希腊","爱尔兰","冰岛","卢森堡","摩纳哥","梵蒂冈",
    "首尔","东京","大阪","京都","北海道","冲绳","巴黎","伦敦","纽约","洛杉矶","旧金山","西雅图",
    "波士顿","芝加哥","华盛顿","柏林","慕尼黑","法兰克福","罗马","米兰","马德里","巴塞罗那",
    "阿姆斯特丹","布鲁塞尔","维也纳","苏黎世","日内瓦","斯德哥尔摩","奥斯陆","哥本哈根",
    "赫尔辛基","雷克雅未克","悉尼","墨尔本","奥克兰","惠灵顿","多伦多","温哥华","蒙特利尔",
    "迪拜","阿布扎比","多哈","利雅得","特拉维夫","伊斯坦布尔","莫斯科","圣彼得堡","新德里",
    "孟买","曼谷","清迈","普吉","吉隆坡","雅加达","马尼拉","河内","胡志明","金边","万象",
    "仰光","加德满都","科伦坡","马尔代夫","开普敦","约翰内斯堡","开罗","内罗毕","拉各斯",
    "圣保罗","里约","布宜诺斯艾利斯","圣地亚哥","利马","波哥大","墨西哥城",
    "弗洛里达","加利福尼亚","德克萨斯","卡尔斯",
]

LOC_SUFFIXES = ["特别行政区","维吾尔自治区","壮族自治区","回族自治区","自治区",
                "自治州","自治县","地区","盟","市","省","县","区"]


# ================= 工具函数 =================
def norm_key(x):
    """把 number / 序号(number) 统一成字符串，便于 join。"""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def clean_loc(x):
    if pd.isna(x):
        return "", ""
    s = str(x).strip()
    s = re.sub(r"[\u200b-\u200f\u202a-\u202e\ufeff]", "", s)
    s = s.strip(" \t\r\n,，;；、")
    if s == "" or s.lower() in ("nan", "none", "null", "0", "０"):
        return "", ""
    s = re.sub(r"(市|省|区|县){2,}$", lambda m: m.group(0)[0], s)
    core = s
    for suf in LOC_SUFFIXES:
        if core.endswith(suf):
            core = core[: -len(suf)]
            break
    return core.strip(), s


def classify_loc(x):
    core, raw = clean_loc(x)
    if core == "":
        return "invalid", ""
    for kw in OVERSEAS_KEYWORDS:
        if kw in raw or kw in core:
            return "overseas", core
    if core in MAINLAND_CITIES:
        return "mainland", core
    return "unknown", core


def calc_city_vars(valid_cities):
    """valid_cities: 仅中国大陆地级市（可重复）。返回 config, dyad, n_city, max_share, k。"""
    k = len(valid_cities)
    if k == 0:
        return "C9", float("nan"), 0, float("nan"), 0
    counts = Counter(valid_cities)
    n_city = len(counts)
    max_city_share = max(counts.values()) / k
    if k <= 1:
        dyad = float("nan")
    else:
        total_pairs = k * (k - 1) // 2
        same_pairs = sum(c * (c - 1) // 2 for c in counts.values())
        dyad = same_pairs / total_pairs
    if n_city == 1:
        config = "C1"
    elif n_city > 1 and max_city_share >= 0.60:
        config = "C2"
    elif n_city > 1 and max_city_share < 0.60:
        config = "C3"
    else:
        config = "C9"
    dyad_out = round(dyad, 4) if dyad == dyad else float("nan")
    return config, dyad_out, n_city, round(max_city_share, 4), k


# ================= 第 1 步：从 test.xlsx 计算城市变量 =================
def compute_city_vars():
    df = pd.read_excel(INPUT_EXCEL, sheet_name=0)
    loc_cols = [c for c in df.columns if str(c).endswith(LOC_SUFFIX)]
    def _up_num(c):
        m = re.search(r"UP(\d+)", str(c))
        return int(m.group(1)) if m else 10**9
    loc_cols.sort(key=_up_num)
    print(f"识别到 {len(loc_cols)} 个地点列：{loc_cols[:5]} ... {loc_cols[-3:]}")

    rows, unknown_counter = [], Counter()
    for _, row in df.iterrows():
        valid, overseas = [], []
        for col in loc_cols:
            cls, core = classify_loc(row.get(col, ""))
            if cls == "mainland":
                valid.append(core)
            elif cls == "overseas":
                overseas.append(core)
            elif cls == "unknown":
                unknown_counter[core] += 1
                overseas.append(core)  # 未识别暂归境外，便于审计

        config, dyad, n_city, max_share, k = calc_city_vars(valid)
        rows.append({
            JOIN_KEY_SRC: norm_key(row.get(JOIN_KEY_SRC, "")),
            "bvid": row.get("bvid", ""),
            "title": row.get("title", ""),
            "valid_cities": ",".join(valid),
            "overseas_locs": ",".join(overseas),
            "k": k,
            "n_city": n_city,
            "max_share": max_share,
            "dyad": dyad,
            "config": config,
        })

    out = pd.DataFrame(rows)
    print("\n=== city_config 分布（源表） ===")
    print(out["config"].value_counts(dropna=False))
    if unknown_counter:
        print("\n⚠️ 未识别地点（已暂归境外，建议人工复核）：")
        for kk, vv in unknown_counter.most_common(30):
            print(f"  {kk}: {vv}")
    return out


# ================= 第 2 步：合并回最终表并原地保存 =================
def merge_into_final(city_df):
    final = pd.read_excel(FINAL_EXCEL)
    print(f"\n最终表原始行数：{len(final)}，列数：{len(final.columns)}")

    if JOIN_KEY_FINAL not in final.columns:
        raise KeyError(f"最终表缺少连接键列：{JOIN_KEY_FINAL}")

    # 统一 key
    final["_key"] = final[JOIN_KEY_FINAL].apply(norm_key)
    city_df = city_df.copy()
    city_df["_key"] = city_df[JOIN_KEY_SRC].apply(norm_key)

    # 若源表 number 有重复，保留最后一条（正常不会重复）
    city_df = city_df.drop_duplicates(subset=["_key"], keep="last")

    # 映射表：key -> 五列
    mapping = city_df.set_index("_key")[[
        "k", "config", "dyad", "n_city", "max_share"
    ]]

    # 按列名写入（若列不存在则新建；存在则按 key 覆盖）
    col_targets = {
        COL_K:        "k",
        COL_CONFIG:   "config",
        COL_DYAD:     "dyad",
        COL_NCITY:    "n_city",
        COL_MAXSHARE: "max_share",
    }
    for target_col, src_col in col_targets.items():
        if target_col not in final.columns:
            final[target_col] = pd.NA
        # 用 map 覆盖：能匹配上的用新值，匹配不上的保持原值（若原本为空则仍为空）
        mapped = final["_key"].map(mapping[src_col])
        final[target_col] = mapped.where(mapped.notna(), final[target_col])

    final = final.drop(columns=["_key"])

    # 统计匹配情况
    matched = final[JOIN_KEY_FINAL].apply(norm_key).isin(set(city_df["_key"])).sum()
    print(f"成功匹配并写入的行数：{matched} / {len(final)}")

    # 备份后原地保存
    backup = FINAL_EXCEL.with_name(FINAL_EXCEL.stem + "_备份" + FINAL_EXCEL.suffix)
    shutil.copyfile(FINAL_EXCEL, backup)
    print(f"已备份原文件 -> {backup}")

    final.to_csv(FINAL_EXCEL.with_suffix(".csv"), index=False, encoding="utf-8-sig")
    try:
        final.to_excel(FINAL_EXCEL, index=False)
        print(f"✅ 已写入：{FINAL_EXCEL}")
    except PermissionError:
        print(f"⚠️ Excel 被占用，已写入 CSV：{FINAL_EXCEL.with_suffix('.csv')}")
        print("   请关闭 Excel 后重跑，或手动把 CSV 另存为 xlsx。")

    return final


# ================= 主流程 =================
def main():
    city_df = compute_city_vars()
    merge_into_final(city_df)


if __name__ == "__main__":
    main()
