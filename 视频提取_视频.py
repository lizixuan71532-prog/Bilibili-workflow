"""
第一阶段：批量下载 B 站视频，按指定比例截取 8 张关键帧，完成后删除视频。
"""

import random
import time
import subprocess
import pandas as pd
import os
import yt_dlp
from pathlib import Path

# ================= 路径配置 =================
# 你的视频源数据表
PROJECT_DIR = Path(__file__).resolve().parent
INPUT_EXCEL = Path(os.environ.get("WORKFLOW_INPUT", PROJECT_DIR / "输入表.xlsx"))
# 视频临时缓存目录（抽完帧后立刻删除）
TEMP_VIDEO_DIR = Path(os.environ.get("TEMP_VIDEO_DIR", PROJECT_DIR / "tmp_video"))
# 最终提取出来的图片保存的根目录
FRAME_OUTPUT_BASE = Path(os.environ.get("FRAME_BASE_DIR", PROJECT_DIR / "frames"))

TEMP_VIDEO_DIR.mkdir(parents=True, exist_ok=True)
FRAME_OUTPUT_BASE.mkdir(parents=True, exist_ok=True)

# 你指定的 8 个时间百分比
TIME_PERCENTAGES = [0.10, 0.20, 0.35, 0.50, 0.65, 0.80, 0.90, 0.95]


def download_and_extract_frames(bvid: str, number: str, video_dir: Path, frame_out_dir: Path):
    """下载视频，用 ffprobe 读取真实时长，按比例截取 8 张图，完成后删除视频。"""
    url = f"https://www.bilibili.com/video/{bvid}/"
    outtmpl = str(video_dir / f"{bvid}.%(ext)s")

    # yt-dlp 配置：只下最烂画质，且只下单个视频
    ydl_opts = {
        "format": "worstvideo*+worstaudio*/worst",
        "outtmpl": outtmpl,
        "quiet": True,
        "no_warnings": True,
        "retries": 1,
        "noplaylist": True,
    }

    try:
        # 1. 下载视频
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.extract_info(url, download=True)

        # 2. 找到下载的文件（格式可能是 mp4, flv, webm）
        candidates = [p for p in video_dir.glob(f"{bvid}.*") if p.suffix not in {".part", ".ytdl"}]
        if not candidates:
            return None, "找不到下载的视频文件"
        video_path = candidates[0]

        # 3. 定位本地的 ffmpeg 和 ffprobe
        ffmpeg_cmd = "ffmpeg"
        ffprobe_cmd = "ffprobe"
        local_ffmpeg = Path(__file__).parent / "ffmpeg.exe"
        local_ffprobe = Path(__file__).parent / "ffprobe.exe"
        if local_ffmpeg.exists():
            ffmpeg_cmd = str(local_ffmpeg)
        if local_ffprobe.exists():
            ffprobe_cmd = str(local_ffprobe)

        # 4. 用 ffprobe 直接读取本地视频的真实时长
        cmd_probe = [
            ffprobe_cmd, "-v", "error", "-show_entries",
            "format=duration", "-of",
            "default=noprint_wrappers=1:nokey=1", str(video_path)
        ]
        result = subprocess.run(cmd_probe, capture_output=True, text=True)
        try:
            duration = float(result.stdout.strip())
        except:
            duration = 0

        if not duration or duration < 1:
            video_path.unlink()
            return None, f"本地视频读取时长异常({duration})，跳过"

        # 5. 创建该视频专属的图片文件夹（以 number 命名）
        frame_out_dir.mkdir(parents=True, exist_ok=True)

        frame_paths = []
        last_timestamp = -1

        # 6. 按比例截取 8 张图
        for i, pct in enumerate(TIME_PERCENTAGES, start=1):
            timestamp = int(duration * pct)

            # 防止短视频截到同一秒
            if timestamp <= last_timestamp:
                timestamp = last_timestamp + 1
            if timestamp >= duration:
                timestamp = max(0, int(duration) - 1)
            last_timestamp = timestamp

            frame_path = frame_out_dir / f"{i}.jpg"

            cmd = [
                ffmpeg_cmd, "-y", "-ss", str(timestamp),
                "-i", str(video_path), "-vframes", "1",
                "-q:v", "2", str(frame_path)
            ]
            subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

            if frame_path.exists():
                frame_paths.append(frame_path)

        # 7. 抽完帧，删除庞大的视频文件
        video_path.unlink()
        return frame_paths, "成功"

    except Exception as e:
        # 清理可能残留的残缺文件
        for p in video_dir.glob(f"{bvid}.*"):
            if p.suffix not in {".part", ".ytdl"}:
                try: p.unlink()
                except: pass
        return None, f"报错: {e}"


def main():
    df = pd.read_excel(INPUT_EXCEL)

    if "number" not in df.columns:
        print("❌ 错误：Excel 文件中缺少 'number' 列！请检查表头。")
        return
    if "bvid" not in df.columns:
        print("❌ 错误：Excel 文件中缺少 'bvid' 列！")
        return

    total = len(df)
    print(f"总共需要处理 {total} 个视频。图片将保存在: {FRAME_OUTPUT_BASE}")

    for idx, row in df.iterrows():
        bvid = str(row["bvid"]).strip()

        # 强制把 number 转为纯数字字符串，防止 1.0 导致文件夹匹配失败
        raw_number = row.get("number", "")
        try:
            number = str(int(float(raw_number))) if pd.notna(raw_number) else ""
        except:
            number = str(raw_number).strip()

        # 该视频的专属文件夹
        video_frame_dir = FRAME_OUTPUT_BASE / number

        # 断点续传检查：如果第 8 张图已经存在，说明这个 number 已经抽完了
        check_frame = video_frame_dir / "8.jpg"
        if check_frame.exists():
            print(f"[{idx + 1}/{total}] 跳过已完成: {number} (bvid: {bvid})")
            continue

        print(f"[{idx + 1}/{total}] 正在处理: {number} (bvid: {bvid}) ...", end=" ", flush=True)

        frames, msg = download_and_extract_frames(bvid, number, TEMP_VIDEO_DIR, video_frame_dir)

        if frames:
            print(f"成功提取 {len(frames)} 帧。")
        else:
            print(f"跳过/失败 ({msg})。")

        # 防封等待，随机波动
        sleep_time = random.uniform(3.0, 7.0)
        time.sleep(sleep_time)

    print("\n所有视频抽帧处理完毕！")


if __name__ == "__main__":
    main()
