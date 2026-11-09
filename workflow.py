"""按四个步骤运行 B 站视频文本、画面、编码与计算流程。"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent


def run_stage(label: str, script: str, extra_args: list[str] | None = None) -> None:
    print(f"\n=== {label} ===", flush=True)
    subprocess.run(
        [sys.executable, str(PROJECT_DIR / script), *(extra_args or [])],
        cwd=PROJECT_DIR,
        check=True,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="提取 B 站视频文本和画面，完成编码与城市计算")
    parser.add_argument("--input", default=os.environ.get("WORKFLOW_INPUT", str(PROJECT_DIR / "输入表.xlsx")))
    parser.add_argument("--skip-transcript", action="store_true")
    parser.add_argument("--skip-frames", action="store_true")
    parser.add_argument("--skip-encoding", action="store_true")
    parser.add_argument("--skip-city-vars", action="store_true")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.is_file():
        raise FileNotFoundError(f"找不到输入表：{input_path}")

    os.environ["WORKFLOW_INPUT"] = str(input_path)
    os.environ.setdefault("WORKFLOW_DIR", str(PROJECT_DIR))
    os.environ.setdefault("INPUT_FILE_NAME", input_path.name)
    os.environ.setdefault("FINAL_OUTPUT", str(PROJECT_DIR / "最终编码结果.xlsx"))
    os.environ.setdefault("FRAME_BASE_DIR", str(PROJECT_DIR / "frames"))
    os.environ.setdefault("TRANSCRIPT_DIR", str(PROJECT_DIR / "transcripts"))
    os.environ.setdefault("AUDIO_DIR", str(PROJECT_DIR / "audio"))
    os.environ.setdefault("TEMP_VIDEO_DIR", str(PROJECT_DIR / "tmp_video"))

    stages = [
        ("提取视频文本信息（音频转写）", "视频提取文本_音频.py", ["--input", str(input_path)], args.skip_transcript),
        ("提取视频画面信息（关键帧）", "视频提取_视频.py", [], args.skip_frames),
        ("多模态视频编码", "视频编码.py", [], args.skip_encoding),
        ("计算城市变量并合并最终表", "视频编码计算补充.py", [], args.skip_city_vars),
    ]
    for index, (label, script, extra_args, skipped) in enumerate(stages, start=1):
        if not skipped:
            run_stage(f"{index}/4 {label}", script, extra_args)

    print(f"\n工作流完成。最终结果：{os.environ['FINAL_OUTPUT']}")


if __name__ == "__main__":
    main()
