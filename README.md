# dsh-mp4-to-execl

适用于 DSH 的信息流广告视频拆解 Skill：输入一条视频链接或本地 MP4，由 AI 分析画面、字幕及原声，生成可编辑的 Excel 分镜拍摄脚本。支持自定义模板，遵循“标题、表头和格式不动，只替换内容”的原则。

> 这是一套 Agent Skills 工作流，不是脱离 AI 视觉模型即可自动理解视频的独立软件。项目名称沿用 dsh-mp4-to-execl。

## 功能

- 视频 URL、平台分享文本、本地 MP4；多个视频分别输出。
- 视频下载、FFprobe 元数据、SHA-256 去重、按时间抽帧及音轨提取。
- AI 镜头分析、景别、拍摄方法、道具、场景、台词及画面字幕。
- 将脚本、时长及真实截图填入 XLSX；保留原模板已有布局、样式和图片锚点。
- 默认生成通用的 15 镜位 Excel 模板；用户提供自定义模板时优先使用自定义模板。
- 校验原模板受保护单元格、公式、合并、图片锚点及未修改部件。

## 工作流

视频链接 / 本地 MP4 → 获取素材 → 提取帧和音频 → AI 分析及核实 → 生成 JSON 填充计划 → OOXML 原位内容替换 → Excel 校验及输出。

## 环境要求

| 工具 | 作用 |
|---|---|
| DSH 及可用的终端、文件读取工具 | Skill 执行 |
| AI 视觉模型或可用图像识别工具 | 实际查看画面与字幕 |
| Python 3.10+ | 工作流脚本 |
| FFmpeg / ffprobe | 抽帧、音轨及视频信息 |
| yt-dlp | 可访问的视频链接下载 |
| lxml / Pillow | OOXML 定点修改及图片处理 |
| openpyxl | 只用于生成默认模板 |
| faster-whisper 或其他 ASR | 可选；逐字台词需真实音频识别 |

## 安装（Windows PowerShell）

先安装 Git、Python 3.10+ 和 FFmpeg，将可执行文件加入 PATH。FFmpeg 安装说明：https://ffmpeg.org/download.html 。

先确认 DSH 当前启用的 Skills 目录。以下以常见的用户目录为例；若本机配置不同，替换为真实目录。

```powershell
$skills = Join-Path $env:USERPROFILE '.dsh\skills'
if ($env:DSH_HOME) { $skills = Join-Path $env:DSH_HOME 'skills' }
New-Item -ItemType Directory -Force $skills | Out-Null
Set-Location $skills
git clone https://github.com/akiraseclab/dsh-mp4-to-execl.git video-script-excel
Set-Location video-script-excel
.\setup.ps1
```

如果 PowerShell 执行策略不允许执行 setup.ps1，按下面命令手工安装即可，无需更改安全策略：

```powershell
python -m pip install -r requirements.txt
python scripts/build_template.py
ffmpeg -version
ffprobe -version
```

安装成功后确认目录中有 SKILL.md，assets/default-template.xlsx 和 assets/blank.jpg。刷新 DSH Skills 列表或新建会话。不同 DSH 版本的技能发现路径、工具权限及模型能力可能不同，请以客户端实际配置为准。

不使用 Git 的用户也可以从仓库的 Code → Download ZIP 下载，解压后按上述路径安装。

## 使用示例

直接对 DSH 发送：

> 使用 video-script-excel，将 https://example.com/video.mp4 拆解为信息流脚本，按照默认模板输出 Excel。

也可以提供本地文件：

> 将 D:\\videos\\ad.mp4 拆解，使用我上传的 XLSX 模板，只替换视频内容和截图，不改标题、格式和镜号。

批量任务：

> 拆解这三条视频链接，每条视频输出一份 Excel。如果视频内容完全相同，按 SHA-256 去重并复用分析结果。

## Excel 输出

默认表头：镜号、镜头类型、画面、示意图、景别、拍摄方法、文本、秒数、拍摄道具、道具图片、拍摄场景。

自定义模板始终按实际表头和原图片位置匹配。禁止覆盖原模板、修改标题、公式或未经确认的单元格。复杂模板（加密、受保护、特殊绘图等）可能不受支持；遇到无法无损处理的情况应停止并说明。

## 手动调用辅助工具

```powershell
python scripts/workflow.py prepare "https://example.com/video.mp4" work --interval 1
python scripts/workflow.py inspect assets/default-template.xlsx --output work/template.json
python scripts/workflow.py frame "D:\\videos\\ad.mp4" 4.5 work/shot-02.jpg
python scripts/workflow.py fill assets/default-template.xlsx work/plan.json work/output.xlsx
```

其中 plan.json 由具备视觉能力的 AI 根据实际视频与 inspect 结果编制。examples/fill-plan.json 是结构示例，不是可直接运行的完整视频分析。

## 文件结构

```text
dsh-mp4-to-execl/
├── SKILL.md
├── README.md
├── setup.ps1
├── requirements.txt
├── scripts/
│   ├── workflow.py
│   └── build_template.py
├── assets/                         # 安装后自动生成
│   ├── default-template.xlsx
│   └── blank.jpg
├── references/
│   ├── tools.md
│   └── validation.md
└── examples/
    └── fill-plan.json
```

## 注意事项

- 仅分析用户有权访问和使用的视频，不绕过登录、付费、验证码或 DRM。
- 链接下载支持取决于平台和实际访问权限。下载失败时请提供可访问文件，不生成虚假分析。
- 未取得音轨转录时必须注明“仅据画面字幕整理”；无法辨认处标注，不编造。
- 原始安装包已完成本地 MP4 抽帧、音轨提取及 Excel 定点写入测试。GitHub 版本增加了模板生成与安装脚本，未在真实 Windows DSH 客户端、各平台链接、ASR 或 Excel GUI 中进行完整端到端测试。
- 默认模板为全新通用模板，公开仓库不包含原始商业模板、原视频或用户私有数据。
- 原 OOXML 定点写入可能保留自定义模板内部旧媒体、共享字符串等数据；对外发布前需检查并清理私有内容。
