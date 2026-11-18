# 工具与故障处理

需要 DSH shell/终端、文件及图像查看能力，Python 3.10+，FFmpeg 与 ffprobe，lxml、Pillow、yt-dlp；生成默认模板时还需 openpyxl。ASR 可选 faster-whisper，逐字台词必须有真实音频识别。

命令（以 Windows PowerShell 为例，先将 $skill 设为技能绝对目录）：

    python "$skill/scripts/build_template.py"
    python "$skill/scripts/workflow.py" prepare "https://example.com/video.mp4" "$work" --interval 1
    python "$skill/scripts/workflow.py" inspect "$skill/assets/default-template.xlsx" --output "$work/template.json"
    python "$skill/scripts/workflow.py" frame "$video" 4.5 "$work/shot-02.jpg"
    python "$skill/scripts/workflow.py" fill "$skill/assets/default-template.xlsx" "$work/plan.json" "$work/output.xlsx"

视频实际路径以 manifest.json 的 source 为准。prepare 最多采样 600 帧，发现截断要补帧。链接下载受站点权限、地区及平台更新影响，失败时请用户提供可合法访问的文件，不绕过访问控制。

可选本地 ASR：经用户授权安装 faster-whisper，首次模型下载需联网。

    from faster_whisper import WhisperModel
    model = WhisperModel('small', device='cpu', compute_type='int8')
    segments, info = model.transcribe('audio.wav', language='zh', vad_filter=True)
    for s in segments:
        print(f'{s.start:.2f} - {s.end:.2f}: {s.text}')

JSON plan：sheets 的键是真实表名；allowed_cells 是可编辑格，protected_cells 是保护格；cells 的值为字符串、有限数值或 null；allowed_images/images 使用 inspect 返回的图位 id。图片路径使用绝对路径。仅支持存在的普通单元格及内嵌 JPEG/PNG，不支持所有复杂模板。原 Excel 内可能保留旧媒体或隐藏数据，向外分发前需额外隐私检查。脚本验证不等于真实 Excel 渲染。
