"""
根据 Excel 中的 bvid,下载 B 站视频音频并转写为文本(极速断点续跑版)。
"""
from __future__ import annotations

import argparse
import json
import random
import re
import shutil
import sys
import time
from pathlib import Path
from typing import Any
import os
import subprocess

import pandas as pd
import yt_dlp
from faster_whisper import WhisperModel

# Hugging Face 镜像 (保留以防某些底层包强制联网)
os.environ["HF_ENDPOINT"] = "https://hf-mirror.com"

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0",
]
EXCEL_ILLEGAL_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def clean_bvid(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    match = re.search(r"(?:https?://www\.bilibili\.com/video/)?((?:BV|av)[A-Za-z0-9]+)", text, re.I)
    return match.group(1) if match else text


def read_input(path: Path, sheet_name: str | int = 0) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path, sheet_name=sheet_name)
    if "bvid" not in df.columns:
        raise ValueError(f"输入文件没有找到 bvid 列。实际列名为:{list(df.columns)}")
    df = df.copy()
    df["bvid"] = df["bvid"].map(clean_bvid)
    df = df[df["bvid"].ne("")].drop_duplicates("bvid").reset_index(drop=True)
    return df


def load_progress(path: Path) -> dict[str, dict[str, Any]]:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        print(f"警告:无法读取进度文件,将从头处理:{path}", file=sys.stderr)
        return {}


def save_progress(path: Path, progress: dict[str, dict[str, Any]]) -> None:
    path.write_text(json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8")


def clean_excel_value(value: Any) -> Any:
    if isinstance(value, str):
        return EXCEL_ILLEGAL_CHARS.sub("", value)
    return value


def write_results(results: list[dict[str, Any]], csv_path: Path, result_path: Path) -> None:
    clean_df = pd.DataFrame(results).map(clean_excel_value)
    clean_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    try:
        clean_df.to_excel(result_path, index=False)
    except Exception as exc:
        print(f"警告: Excel 导出失败，CSV 和 progress.json 已保存: {type(exc).__name__}: {exc}", file=sys.stderr)


def resolve_ffmpeg_location(value: str | None) -> str | None:
    if value:
        candidate = Path(value).expanduser()
        if candidate.is_file():
            candidate = candidate.parent
        if (candidate / "ffmpeg.exe").exists() and (candidate / "ffprobe.exe").exists():
            return str(candidate)
        raise ValueError(f"--ffmpeg-location 无效，目录内需同时包含 ffmpeg.exe 和 ffprobe.exe: {value}")

    ffmpeg = shutil.which("ffmpeg")
    ffprobe = shutil.which("ffprobe")
    if ffmpeg and ffprobe:
        return str(Path(ffmpeg).parent)
    return None


def build_ydl_options(audio_dir: Path, bvid: str, cookies: str | None, proxy: str | None, ffmpeg_location: str | None) -> dict[str, Any]:
    outtmpl = str(audio_dir / f"{bvid}.%(ext)s")
    headers = {
        "User-Agent": random.choice(USER_AGENTS),
        "Referer": f"https://www.bilibili.com/video/{bvid}/",
        "Origin": "https://www.bilibili.com",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Cache-Control": "max-age=0",
        "Sec-Ch-Ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
        "Sec-Ch-Ua-Mobile": "?0",
        "Sec-Ch-Ua-Platform": '"Windows"',
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
    }
    options: dict[str, Any] = {
        "format": "bestaudio/best",
        "outtmpl": outtmpl,
        "noplaylist": True,
        "quiet": True,
        "no_warnings": True,
        "retries": 1,
        "fragment_retries": 1,
        "continuedl": True,
        "overwrites": False,
        "socket_timeout": 30,
        "http_headers": headers,
        "sleep_interval_requests": random.uniform(1.0, 3.0),
    }
    if ffmpeg_location:
        options["ffmpeg_location"] = ffmpeg_location
        options["postprocessors"] = [{"key": "FFmpegExtractAudio", "preferredcodec": "mp3", "preferredquality": "128"}]
    if cookies:
        options["cookiefile"] = cookies
    if proxy:
        options["proxy"] = proxy
    return options


def download_audio(bvid: str, audio_dir: Path, cookies: str | None = None, sleep_seconds: float = 6.0, proxy: str | None = None, retries: int = 1, ffmpeg_location: str | None = None) -> tuple[Path | None, dict[str, Any], str | None, bool]:
    audio_dir.mkdir(parents=True, exist_ok=True)
    url = f"https://www.bilibili.com/video/{bvid}/"

    last_error = ""
    info: dict[str, Any] = {}
    for attempt in range(1, retries + 1):
        options = build_ydl_options(audio_dir, bvid, cookies, proxy, ffmpeg_location)
        try:
            with yt_dlp.YoutubeDL({"quiet": True, "no_warnings": True}) as ydl:
                info = ydl.extract_info(url, download=False)

            duration = info.get("duration", 0)
            if duration and duration > 15 * 60:
                return None, info, "", True

            with yt_dlp.YoutubeDL(options) as ydl:
                info = ydl.extract_info(url, download=True) or {}

            audio_path = audio_dir / f"{bvid}.mp3"
            if not audio_path.exists():
                candidates = sorted(p for p in audio_dir.glob(f"{bvid}.*") if p.suffix not in {".part", ".ytdl"} and p.is_file())
                audio_path = candidates[0] if candidates else audio_path

            if not audio_path.exists():
                return None, info, "下载完成但找不到音频文件。", False

            return audio_path, info, None, False

        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            lower = last_error.lower()
            hint = ""
            if "412" in lower or "403" in lower or "forbidden" in lower:
                hint = "(疑似风控/登录态失效,请更新 cookies.txt 或稍后再试)"
            elif "unable to extract initial state" in lower:
                hint = "(页面解析失败；确认正在使用 .venv312 的新版 yt-dlp)"
            elif "http error 567" in lower:
                hint = "(B站拒绝请求；停止自动重试，稍后再续跑并确认 cookies 有效)"

            print(f"  下载失败 第{attempt}/{retries}次:{last_error} {hint}", file=sys.stderr)

            if any(marker in lower for marker in ("http error 567", "unable to extract initial state", "412", "403", "forbidden")):
                break
            if attempt < retries:
                backoff = (2 ** (attempt - 1)) * random.uniform(2.0, 4.0)
                time.sleep(backoff)

    return None, info, f"下载失败(已重试{retries}次):{last_error}", False


def transcribe(audio_path: Path, model: WhisperModel, language: str | None) -> tuple[str, float]:
    try:
        segments, info = model.transcribe(str(audio_path), language=language, beam_size=5, vad_filter=True, condition_on_previous_text=True)
        parts = [segment.text.strip() for segment in segments if segment.text.strip()]
        return "\n".join(parts), float(getattr(info, "duration", 0.0) or 0.0)
    except TypeError as e:
        if "metadata_errors" in str(e):
            print("  [转写纠错] 检测到音频格式底层解析报错，尝试转为 wav 后重试...")
            ffmpeg_cmd = "ffmpeg"
            local_ffmpeg = Path(__file__).parent / "ffmpeg.exe"
            if local_ffmpeg.exists():
                ffmpeg_cmd = str(local_ffmpeg)

            wav_path = audio_path.with_suffix('.wav')
            try:
                subprocess.run([ffmpeg_cmd, "-y", "-i", str(audio_path), str(wav_path)], capture_output=True, check=True)
                segments, info = model.transcribe(str(wav_path), language=language, beam_size=5, vad_filter=True, condition_on_previous_text=True)
                parts = [segment.text.strip() for segment in segments if segment.text.strip()]
                return "\n".join(parts), float(getattr(info, "duration", 0.0) or 0.0)
            finally:
                if wav_path.exists():
                    wav_path.unlink()
        else:
            raise e


def main() -> None:
    parser = argparse.ArgumentParser(description="按 bvid 下载 B 站音频并转写为文本(极速断点续跑)")
    project_dir = Path(__file__).resolve().parent
    parser.add_argument("--input", default=os.environ.get("WORKFLOW_INPUT", str(project_dir / "输入表.xlsx")))
    parser.add_argument("--output-dir", default=os.environ.get("TRANSCRIPT_DIR", str(project_dir / "transcripts")))
    parser.add_argument("--audio-dir", default=os.environ.get("AUDIO_DIR", str(project_dir / "audio")))
    parser.add_argument("--cookies", default=os.environ.get("BILIBILI_COOKIES_FILE", ""))
    parser.add_argument("--sheet", default="0")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--model", default="base")
    parser.add_argument("--device", default="auto", choices=["auto", "cpu", "cuda"])
    parser.add_argument("--compute-type", default="auto")
    parser.add_argument("--language", default="zh")
    parser.add_argument("--sleep", type=float, default=6.0)
    parser.add_argument("--retries", type=int, default=1)
    parser.add_argument("--proxy", default=None)
    parser.add_argument("--ffmpeg-location", default=None)
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    audio_dir = Path(args.audio_dir).expanduser().resolve()
    audio_dir.mkdir(parents=True, exist_ok=True)

    progress_path = output_dir / "progress.json"
    result_path = output_dir / "transcript_results.xlsx"
    csv_path = output_dir / "transcript_results.csv"

    cookies_path = Path(args.cookies).expanduser().resolve() if args.cookies else None
    if cookies_path and not cookies_path.exists():
        print(f"警告:cookies 文件不存在:{cookies_path}", file=sys.stderr)
        cookies_path = None

    ffmpeg_location = resolve_ffmpeg_location(args.ffmpeg_location)
    if ffmpeg_location:
        print(f"FFmpeg: {ffmpeg_location}")
    else:
        print("未检测到 FFmpeg 环境变量：将保留原始音频直接转写。")

    sheet: str | int = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
    df = read_input(input_path, sheet)
    if args.limit > 0:
        df = df.head(args.limit)

    device = "cpu" if args.device == "auto" or args.device == "cpu" else "cuda"
    compute_type = "int8" if args.compute_type == "auto" and device == "cpu" else ("float16" if args.compute_type == "auto" else args.compute_type)
    language = None if str(args.language).lower() in {"none", "auto", ""} else args.language

    print(f"读取 {len(df)} 个 bvid;模型={args.model},设备={device},compute_type={compute_type}")

    model_path = args.model
    local_model_paths = {
        "small": os.environ.get("WHISPER_SMALL_PATH", ""),
        "base": os.environ.get("WHISPER_BASE_PATH", ""),
    }
    if args.model in local_model_paths:
        local_path = local_model_paths[args.model]
        if local_path and Path(local_path).exists():
            model_path = local_path
            print(f"✔️ 使用本地 {args.model} 模型：{local_path}")
        else:
            print(f"⚠️ 警告：本地 {args.model} 模型目录不存在，将尝试联网下载。", file=sys.stderr)

    model = WhisperModel(model_path, device=device, compute_type=compute_type)

    # 加载进度
    progress = load_progress(progress_path)
    results: list[dict[str, Any]] = []

    try:
        for i, row in df.iterrows():
            bvid = row["bvid"]
            base = {"bvid": bvid, "url": f"https://www.bilibili.com/video/{bvid}/"}

            # ================= 最强断点跳过逻辑 =================
            # 1. 如果在 progress.json 里有记录（无论成功、失败还是跳过），直接跳过
            if bvid in progress:
                results.append({**row.to_dict(), **progress[bvid]})
                print(f"[{i + 1}/{len(df)}] 跳过已处理记录:{bvid}")
                continue

            # 2. 如果本地 E 盘已经有这个 mp3/m4a 等文件了，也直接跳过下载，进入转写（如果你要补转以前的音频）
            existing_audio = list(audio_dir.glob(f"{bvid}.*"))
            existing_audio = [p for p in existing_audio if p.suffix not in {".part", ".ytdl"}]
            if existing_audio:
                print(f"[{i + 1}/{len(df)}] 发现本地已有音频，直接转写:{bvid}")
                audio_path = existing_audio[0]
                info = {}
                error = None
                is_skipped = False
            else:
                print(f"[{i + 1}/{len(df)}] 处理:{bvid}")
                audio_path, info, error, is_skipped = download_audio(
                    bvid, audio_dir,
                    cookies=str(cookies_path) if cookies_path else None,
                    sleep_seconds=args.sleep,
                    proxy=args.proxy,
                    retries=args.retries,
                    ffmpeg_location=ffmpeg_location,
                )

            record = {
                **row.to_dict(),
                **base,
                "title_fetched": info.get("title", ""),
                "uploader_fetched": info.get("uploader", ""),
                "duration_seconds": info.get("duration", ""),
                "audio_path": str(audio_path) if audio_path else "",
                "transcript": "",
                "status": "skipped" if is_skipped else "failed",
                "error": "视频超过15分钟，已跳过" if is_skipped else (error or ""),
            }

            if is_skipped:
                print(f"  [跳过] 视频时长为 {info.get('duration', 0) / 60:.1f} 分钟，超过限制。")

            elif audio_path and not error:
                try:
                    text, duration = transcribe(audio_path, model, language)
                    record["transcript"] = text
                    record["duration_seconds"] = duration or record["duration_seconds"]
                    record["status"] = "success"
                    record["error"] = ""
                except Exception as exc:
                    record["error"] = f"转写失败:{type(exc).__name__}: {exc}"

                # 下载并转写成功后，防封等待
                if not error:
                    sleep_time = args.sleep + random.uniform(1.0, 3.0)
                    print(f"  [防封等待] 休眠 {sleep_time:.1f} 秒...")
                    time.sleep(sleep_time)

            # 无论结果如何，全都写入 progress.json（这就是为什么以后绝不回头的关键）
            progress[bvid] = {k: v for k, v in record.items() if not isinstance(v, (pd.Timestamp,))}
            save_progress(progress_path, progress)

            results.append(record)
            write_results(results, csv_path, result_path)

            error_lower = str(record.get("error", "")).lower()
            if any(marker in error_lower for marker in ("http error 567", "unable to extract initial state", "412", "403", "forbidden")):
                print("\n[⚠️ 风控拦截] 检测到 B 站限制。已保存进度并暂停，等一段时间后重新启动即可续跑。", file=sys.stderr)
                break

    except KeyboardInterrupt:
        save_progress(progress_path, progress)
        print("\n用户中断。已保存进度；重新运行将跳过已成功条目。", file=sys.stderr)
        return

    print(f"\n✅ 完成。结果:{result_path}")

if __name__ == "__main__":
    main()
