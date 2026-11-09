# B站视频文本—画面—编码工作流

此副本只包含四步工作流，不包含“视频简介，标签”环节：

1. 从视频音频提取文本（Whisper 转写）
2. 从视频抽取关键帧作为画面信息
3. 对视频内容进行多模态编码
4. 计算城市变量并合并最终表

## 输入和结果

把 `.xlsx` 输入表放在项目目录，默认命名 `输入表.xlsx`，或在运行时通过 `--input` 指定。表格需含 `bvid`、`number` 及编码/城市计算需要的字段。最终表默认为 `最终编码结果.xlsx`，使用输入表 `number` 与结果表 `序号(number)` 连接。

城市变量仅统计中国大陆地级市；境外及港澳台不计入。输入表、结果表、音频、关键帧和进度文件均为本地数据，不应上传。

## 安装

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

还需安装 FFmpeg，并确保 `ffmpeg` 和 `ffprobe` 可从 PATH 调用。`faster-whisper` 可使用 CPU；如需 GPU，请按本机 CUDA 环境配置。

## 配置密钥和 Cookie

多模态编码需要 DashScope API Key。在 PowerShell 当前会话设置：

```powershell
$env:DASHSCOPE_API_KEY = "你的 DashScope API Key"
```

如下载需要 B 站登录态，把 Netscape 格式 Cookie 文件保存在项目目录以外，再设置：

```powershell
$env:BILIBILI_COOKIES_FILE = "D:\private\cookies.txt"
```

不要把 Key/Cookie 写入源码或上传。API Key 和 Cookie 曾经出现在原始源码中，建议先轮换/重新生成后再使用。

## 运行

```powershell
python workflow.py --input .\输入表.xlsx
```

可用 `--skip-transcript`、`--skip-frames`、`--skip-encoding`、`--skip-city-vars` 跳过已完成阶段。输出目录可通过 `FINAL_OUTPUT`、`FRAME_BASE_DIR`、`TRANSCRIPT_DIR`、`AUDIO_DIR`、`TEMP_VIDEO_DIR` 环境变量调整。

## 上传前检查

```powershell
git status --short
```

确保没有输入/结果表、CSV、Cookie、`.env` 或媒体文件被暂存。`.gitignore` 已排除这些数据文件。
