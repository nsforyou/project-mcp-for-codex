# Project MCP for codex

作者：JhonPrice、ChatGPT-6.1 Sol

[English](README.en.md) · 中文

[非商用许可](LICENSE) · [提交问题或建议](https://github.com/nsforyou/project-mcp-for-codex/issues)

让 Codex 读取和修改 Windows 上的 Microsoft Project 计划。你可以保持 Project 打开，与 Codex 讨论修改方案，确认后让改动显示在当前窗口，再决定保存或撤销。

当前版本 **0.4.0**，提供 **30 个 MCP 工具**，也可直接通过 CLI 使用。仓库仅包含任务 **A/B**、资源组 **AA/BB** 的合成测试计划。

## 目录

- [可以做什么](#可以做什么)
- [环境要求](#环境要求)
- [安装与检查](#安装与检查)
- [接入 Codex](#接入-codex)
- [第一次使用：用 A/B 测试计划体验](#第一次使用用-ab-测试计划体验)
- [在线操作的完整 CLI 示例](#在线操作的完整-cli-示例)
- [备注中的工程量和状态字段](#备注中的工程量和状态字段)
- [关闭 Project 后进行离线修改](#关闭-project-后进行离线修改)
- [从 XML 新建 MPP](#从-xml-新建-mpp)
- [甘特图排版与 PDF 交付](#甘特图排版与-pdf-交付)
- [30 个工具](#30-个工具)
- [结果状态、备份与恢复](#结果状态备份与恢复)
- [常见问题](#常见问题)
- [测试与更新](#测试与更新)
- [项目结构与许可](#项目结构与许可)

## 可以做什么

| 需求 | 使用方式 | Project 是否需要关闭 |
|---|---|---|
| 开着计划讨论、读取未保存的当前内容 | 在线读取 | 否 |
| 确认后修改当前任务，在窗口里查看 | 在线预览 → 在线执行 | 否；先保存并结束单元格编辑 |
| 分析已保存的任务、依赖、资源、日历 | 普通读取，多数使用 MPXJ | 否；只能看到磁盘版本 |
| 读取视图、完整备注，预览或执行离线修改 | 离线 COM 通道 | 是 |
| 备份后修改并保存原 MPP | 离线变更集 | 是 |
| 从基础 Project XML 生成一个新 MPP | 新建工具 | 是 |

支持的常见工作包括：调整名称、起止时间、工期、完成率和前置关系，修改资源名称，维护备注中的工程量和当前状态，导出 Excel，以及检查依赖、关键任务和基线信息。

当前不提供自动施工方案决策或完整的资源平衡算法。已有计划如果缺少依赖，或任务使用手动排程，不能据此假定“修改日期后，所有下游任务都会自动正确顺延”。

## 环境要求

| 组件 | 要求 |
|---|---|
| 系统 | Windows；在运行 Project 的同一用户桌面会话中执行 |
| Microsoft Project | 本机桌面版，提供 COM 接口；已验证 Professional 2019 |
| Python | 推荐 64 位 Python 3.13，已在该版本验证 |
| Java | Java 21，用于 MPXJ 读取和独立核验 |
| Git | 克隆和更新仓库 |
| Codex | 使用 MCP 时需要；只使用本项目 CLI 时不需要 |

Project 网页版、Planner，以及单独的 macOS/Linux/WSL 环境不能直接代替 Windows Project COM。请使用 **Windows Python**，并先手动打开一次 Project，完成首次启动、激活或提示处理。

本项目服务在本机读写文件，不直接调用 OpenAI API，也不要求配置 OpenAI API Key。Codex 自身的登录与使用条件另行适用。文件在本机并不代表模型离线运行：工具返回的计划内容可能由宿主提供给所选模型。

## 安装与检查

以下示例统一使用：

- 仓库：C:\ProjectMCP\project-mcp-for-codex
- MPXJ 依赖：C:\ProjectMCP\deps
- 演示计划：C:\ProjectMCP\plans\demo.mpp

可换成自己的路径。MPXJ 原生依赖目录请使用纯 ASCII 路径，避免中文路径加载问题。

### 1. 下载并安装 Python 依赖

在 Windows PowerShell 中执行：

~~~powershell
git clone https://github.com/nsforyou/project-mcp-for-codex.git C:\ProjectMCP\project-mcp-for-codex
Set-Location C:\ProjectMCP\project-mcp-for-codex

python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

$env:MPXJ_LIB = 'C:\ProjectMCP\deps'
.\.venv\Scripts\python.exe -m pip install --target $env:MPXJ_LIB -r requirements-mpxj.txt
~~~

无需激活虚拟环境；后续直接调用虚拟环境中的 python.exe，避免解释器或执行策略不一致。

### 2. 检查 Java 和工具清单

先安装 Java 21，设置 JAVA_HOME 或将 Java 加入 PATH，然后检查：

~~~powershell
java -version
.\.venv\Scripts\python.exe -c "import struct; print('Python bits:', struct.calcsize('P') * 8)"
.\.venv\Scripts\python.exe -c "import win32com.client, psutil, olefile, openpyxl; print('Python dependencies OK')"

$tools = .\.venv\Scripts\python.exe mcp\project_cli.py tools | ConvertFrom-Json
$tools.Count
~~~

正常情况下工具数量为 **30**。列出工具成功只说明服务代码和基础依赖可加载，不代表 Java、MPXJ 或 Project COM 都已验证。

PowerShell 中的 $env:MPXJ_LIB 只影响当前会话及其子进程。新终端需要重新设置；Codex 启动服务时使用下面配置中的环境变量。未设置时，桥接器默认查找 %LOCALAPPDATA%\ProjectMCP\deps。

## 接入 Codex

选择以下一种方式即可，避免重复注册同一个服务。Codex 的 MCP 命令和配置格式可参考 [OpenAI 官方 MCP 文档](https://developers.openai.com/codex/mcp/)。

### 方式一：修改配置文件

编辑 Windows 默认用户配置文件 %USERPROFILE%\.codex\config.toml；如果自定义了 CODEX_HOME，以实际位置为准。在现有配置中添加或更新以下部分：

~~~toml
[mcp_servers.microsoft-project]
command = 'C:\ProjectMCP\project-mcp-for-codex\.venv\Scripts\python.exe'
args = ['C:\ProjectMCP\project-mcp-for-codex\mcp\project_mcp_server.py']
startup_timeout_sec = 20
tool_timeout_sec = 180

[mcp_servers.microsoft-project.env]
MPXJ_LIB = 'C:\ProjectMCP\deps'
PROJECT_COM_TIMEOUT = '120'
~~~

这里使用 TOML 单引号字符串，Windows 路径的反斜杠不需要重复转义。请使用真实的绝对路径；虚拟环境的 Python 路径尤其重要。

microsoft-project 是本例的配置键，服务报告的标识为 project-mcp-for-codex。PROJECT_COM_TIMEOUT 控制离线桥接超时；在线通道自身的工作进程超时固定为 20 秒。tool_timeout_sec 是 Codex 等待工具结果的时间，两者不是同一个设置。

### 方式二：通过 Codex CLI 注册

在已有 Codex CLI 的 Windows PowerShell 中执行：

~~~powershell
codex mcp add microsoft-project --env MPXJ_LIB=C:\ProjectMCP\deps --env PROJECT_COM_TIMEOUT=120 -- C:\ProjectMCP\project-mcp-for-codex\.venv\Scripts\python.exe C:\ProjectMCP\project-mcp-for-codex\mcp\project_mcp_server.py
codex mcp list
codex mcp get microsoft-project
~~~

如果要调整客户端工具超时，再编辑配置文件中的对应项。注册后重启或重新连接 Codex 的 MCP 服务；如果当前会话仍显示旧工具清单，重新开启会话。

这是本地 stdio 服务，客户端负责启动它；不需要另外开启 HTTP 端口。手动运行 project_mcp_server.py 后等待输入是正常现象，它不会显示图形界面。普通 ChatGPT 网页不会因你修改本机 Codex 配置而自动连接此服务。

## 第一次使用：用 A/B 测试计划体验

### 1. 复制测试文件

首次演示先创建副本，避免直接修改仓库里的标准样本：

~~~powershell
New-Item -ItemType Directory -Path C:\ProjectMCP\plans -Force | Out-Null
if (Test-Path -LiteralPath C:\ProjectMCP\plans\demo.mpp) {
    throw 'demo.mpp 已存在，请换一个演示文件名，避免覆盖。'
}
Copy-Item -LiteralPath .\mcp\tests\fixtures\A_B_AA_BB.mpp -Destination C:\ProjectMCP\plans\demo.mpp
~~~

标准样本包含两个任务 A、B，前置关系为 A → B（FS）。资源组 AA、BB 的容量分别为 6、4；A 使用 AA 两人，B 使用 BB 四人。任务工期均为两个标准工作日。

### 2. 打开 Project，再与 Codex 对话

手动用 Project 打开 demo.mpp，使其成为当前活动项目。一次只保留一个 Project 进程。

可以直接向 Codex 说：

> 使用 microsoft-project 的 project_live_read 读取 C:\ProjectMCP\plans\demo.mpp。先列出任务的 UniqueID、名称、起止日期、资源和备注，只读取，不修改。

然后继续讨论：

> 把任务 B 的完成日期调整到 2026-11-06 17:00，备注中记录当前状态为“施工中”。先用在线通道生成修改预览，等我确认后再执行。

请根据实际返回的任务信息确认方案，不要仅凭任务名称或屏幕行号定位。

### 3. 确认并在线修改

1. 在 Project 中 Ctrl+S 保存当前工作，按 Enter 结束单元格编辑，关闭模态对话框。
2. 让 Codex 调用 project_live_preview，查看修改前后的内容。
3. 确认方案后明确说“按这个预览执行”。
4. Codex 使用该 LIVE 变更 ID 调用 project_live_apply。
5. 执行的几秒内停止手动编辑。
6. 查看同一窗口里的结果，满意后 Ctrl+S 保存；不满意可从撤销列表中撤销整批“MCP 在线调整”。

在线执行默认只改内存，**不会自动保存**。如果你在预览后又手动修改、保存计划，或重启了 Project，需要重新预览，不能复用旧 ID。

## 在线操作的完整 CLI 示例

不使用 Codex 也可以调用相同处理器。以下命令在仓库根目录执行；将 JSON 代码块分别保存为 UTF-8 文件，文件名与命令一致。不要把 JSON 当成 PowerShell 命令运行。

每次 CLI 调用都是独立进程，因此读取和预览时都要明确传入 file。在线工具会读取 Project 当前内存；普通读取工具分析的是已保存文件。

### 1. 读取

将以下内容保存为 read.json：

~~~json
{
  "file": "C:\\ProjectMCP\\plans\\demo.mpp"
}
~~~

~~~powershell
.\.venv\Scripts\python.exe mcp\project_cli.py call project_live_read --args-file read.json
~~~

返回内容包括 file、pid、saved、tasks、resources 和 fingerprint。在线 tasks 中的 UniqueID 是稳定任务标识，ID 是当前行号。两者不能混用；插入或移动任务后，行号可能改变。

### 2. 生成预览

先在 Project 中保存。将以下内容保存为 preview.json：

~~~json
{
  "file": "C:\\ProjectMCP\\plans\\demo.mpp",
  "title": "调整任务 B 并更新备注",
  "ops": [
    {
      "unique_id": 2,
      "fields": {
        "Finish": "2026-11-06T17:00"
      },
      "note_fields": {
        "工程量": 120,
        "单位": "米",
        "当前状态": "施工中",
        "已完成量": 35
      }
    }
  ]
}
~~~

标准样本 B 的 UniqueID 为 2；自己的计划必须以实际读取值为准。

~~~powershell
.\.venv\Scripts\python.exe mcp\project_cli.py call project_live_preview --args-file preview.json
~~~

此时不修改任务。记录返回的 change_id，例如 LIVE- 后跟一段实际标识，并核对 preview 中的 before 和 fields。

### 3. 确认执行

将以下内容保存为 apply.json，并把示例占位内容替换为实际 ID：

~~~json
{
  "change_id": "LIVE-替换为预览返回的真实ID",
  "approved_by": "用户已确认本次调整"
}
~~~

~~~powershell
.\.venv\Scripts\python.exe mcp\project_cli.py call project_live_apply --args-file apply.json
~~~

检查 status、checks 和 other_changes；后者列出 Project 计算导致的其他字段变化。状态 memory_verified 表示所请求字段在当前内存中核验一致，仍需用户查看并保存，不代表已落盘。

### 在线字段和单位

| 字段 | 示例 | 说明 |
|---|---|---|
| Name | "A" | 任务名称 |
| Notes | "现场说明" | 直接写 Notes 会替换完整备注；维护字段时优先用 note_fields |
| Start / Finish | "2026-11-04T08:00" | 推荐明确到分钟，按本地时间理解 |
| Duration | 960 | 工作分钟数；不要把它当成自然天数 |
| PercentComplete | 25 | 完成率，范围 0～100 |
| Predecessors | "1FS" | Project 前置关系文本，引用当前行号，不是 UniqueID |
| ResourceNames | "AA" | 修改资源名称分配，不自动新增资源、设置人数或完成资源平衡 |
| note_fields | {"当前状态":"施工中"} | 逐项合并备注内字段 |

Duration=960 在每天 8 小时的标准工作日历下对应 2 个工作日；自定义日历下不要直接据此换算。在线摘要任务只允许修改 Name 和 Notes。

## 备注中的工程量和状态字段

Project 的 Notes 本质上仍是一个备注字段。本项目把每一行“字段名：值”解释为结构化数据，例如：

~~~text
现场说明文字，保留在字段区外。

【MCP任务字段:v1】
工程量：120
单位：米
当前状态：施工中
已完成量：35
责任班组：AA
【/MCP任务字段】
~~~

字段名可自定义；数量和单位分开写。工程量、已完成量要求非负数字。值为单行文字、数字或布尔值；读取时通常返回文字。当前状态没有强制枚举。

- 在线修改：在 project_live_preview 的每项操作中使用 note_fields。
- 在线读取：project_live_read 返回完整 Notes，可据此解析和讨论。
- 离线读取/更新：关闭 Project 后使用 project_get_task_note_fields、project_preview_task_note_fields。
- 不带标记的旧备注也能按冒号识别；首次更新会集中字段行，保留其余说明。
- 重复字段、破损标记或字段区内的非字段说明会拒绝更新，避免误删内容。
- 备注字段不会自动改变 Project 的完成率、工期或资源分配，也不会自动成为可供 Project 筛选的自定义列。

离线按状态批量查询的参数示例：

~~~json
{
  "file": "C:\\ProjectMCP\\plans\\demo.mpp",
  "field_name": "当前状态",
  "equals": "施工中"
}
~~~

将它作为参数文件传给 project_get_task_note_fields；不传 unique_id 时批量读取。

## 关闭 Project 后进行离线修改

需要直接保存到 MPP，或使用原有变更集工具时，先保存并关闭所有 Project 进程。不要用 project_open 代替在线连接：它只建立服务的只读文件上下文，不打开图形窗口。

预览参数示例，可保存为 offline-preview.json：

~~~json
{
  "file": "C:\\ProjectMCP\\plans\\demo.mpp",
  "title": "调整 B 的名称",
  "ops": [
    {
      "op": "set_task_fields",
      "unique_id": 2,
      "fields": {
        "Name": "B - 已复核"
      }
    }
  ]
}
~~~

~~~powershell
.\.venv\Scripts\python.exe mcp\project_cli.py call project_preview_changes --args-file offline-preview.json
~~~

确认 errors 为空，再将实际 change_set_id 写入 offline-apply.json：

~~~json
{
  "change_set_id": "CS-替换为预览返回的真实ID",
  "approved_by": "用户已确认本次调整",
  "allow_in_place": true
}
~~~

~~~powershell
.\.venv\Scripts\python.exe mcp\project_cli.py call project_apply_changes --args-file offline-apply.json
~~~

离线流程会备份、写入、保存并重读核验。执行后再打开 Project 查看。

请明确使用 allow_in_place=true。当前变更集默认采用备份后就地保存的模式；这个参数表示允许覆盖源文件。传 false 不会自动切换为另存模式，而会拒绝执行该就地变更集。

离线操作示例还包括：

~~~json
{"op":"append_task_note","unique_id":2,"note":"现场复核记录"}
~~~

~~~json
{"op":"set_predecessors","row_id":2,"predecessors":"1FS"}
~~~

前置关系操作使用 row_id；其他常见任务字段使用 unique_id。不要在插入任务后继续沿用过时的行号。

## 从 XML 新建 MPP

这是基础创建器，不是任意 MPP/XML 的完整无损转换器。Project 应处于关闭状态，目标文件必须不存在。

保存 create-preview.json，xml 请替换成仓库中的实际绝对路径：

~~~json
{
  "xml": "C:\\ProjectMCP\\project-mcp-for-codex\\mcp\\tests\\basic_plan.xml",
  "output": "C:\\ProjectMCP\\plans\\new-demo.mpp",
  "preview": true
}
~~~

~~~powershell
.\.venv\Scripts\python.exe mcp\project_cli.py call project_create_from_xml --args-file create-preview.json
~~~

确认预览后，在参数中设置 preview=false，并补上返回的 source_sha256，再调用同一工具。源 XML 变化后必须重新预览。

支持基础层级、明细时间、PT 格式工期、资源 MaxUnits、分配 Units、四类依赖及工作分钟延迟、里程碑。新建计划使用手动排程；摘要日期由子任务汇总。资源容量缺失时采用 1，不从名称猜测人数。XML 的资源 Units/MaxUnits 中 1 代表一个完整资源单位，6 代表六个资源单位，在 Project 中通常显示为 600%。

自定义日历、实际日期、基线、已识别的硬约束和截止日期会被拒绝。成本、费率、分段工作量、原 MPP 视图等不在导入范围。不要用它重建正式 MPP 来代替原生修改。

创建结果旁的 .creation.json 保存核验信息。生成文件但核验失败时，不会报告为成功，也不会自动覆盖或删除该文件。

## 甘特图排版与 PDF 交付

0.4.0 新增三个共用 MCP/CLI 的工具：`project_preview_print_layout`、`project_export_pdf`、`project_check_pdf_layout`。

支持 A3/A4、横向/纵向、显示列与列宽、长名称换行、9–16 磅字号、日/周刻度、日期窗口、纯文本页眉页脚，以及甘特条内部显示“完成百分比”。百分比绑定 Project 字段，不改变实际进度。

~~~powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-print.txt
~~~

例如告诉 Codex：“把计划导出为 A3 横向 PDF，任务名称换行，甘特条内部显示完成百分比；先给排版预览，再导出到新目录并检查 PDF。”

预览不连接 Project、不写文件。执行需匹配的 `preview_sha256`，只在源 MPP 的副本上排版和保存，拒绝覆盖输出目录，并核验排版前后和重新打开后的任务、资源及分配一致。源文件保持原样。导出使用共享 COM 锁，要求 Project 已关闭。

输出包括 `layout.mpp`、`native.pdf`、补充页码的 `print.pdf` 和 `verification.json`。成功状态为 `exported_needs_visual_review`，基础检查不等于逐页版面验收；此版本不向实体打印机提交作业。完整参数、CLI示例、已知限制及并行实例的锁配置见 [排版使用指南](docs/PRINTING.md)。

## 30 个工具

准确参数以实际服务的工具清单为准：

~~~powershell
.\.venv\Scripts\python.exe mcp\project_cli.py tools
~~~

| 分类 | 工具 |
|---|---|
| 在线协作（3） | project_live_read、project_live_preview、project_live_apply |
| 文件上下文（2） | project_open、project_close |
| 项目与计划（3） | project_get_project_info、project_get_schedule_summary、project_get_critical_path |
| 任务与依赖（4） | project_list_tasks、project_find_tasks、project_get_task、project_get_dependencies |
| 资源与视图（4） | project_list_resources、project_list_calendars、project_get_baselines、project_get_view_layer |
| 离线变更集（5） | project_preview_changes、project_apply_changes、project_verify_changes、project_rollback_changes、project_list_change_sets |
| 备注与进度（4） | project_append_task_note、project_set_actual_progress、project_get_task_note_fields、project_preview_task_note_fields |
| 新建与导出（2） | project_create_from_xml、project_export_excel |
| 排版与 PDF（3） | project_preview_print_layout、project_export_pdf、project_check_pdf_layout |

project_close 只清理 MCP 的文件上下文，不会关闭用户的 Project 窗口。离线完整备注和视图读取走 COM，仍要求关闭 Project。

## 结果状态、备份与恢复

| 状态 | 含义与下一步 |
|---|---|
| preview | 已生成预览，没有执行 |
| memory_verified | 在线修改已在内存核验；用户查看后保存或撤销 |
| verified | 离线修改已落盘，且该次核验通过 |
| applied_with_warnings | 已写入但业务字段或视图检查存在警告，需查看 verification |
| needs_review | 在线结果不确定或核验不一致，检查当前 Project，勿自动重试 |
| failed | 执行失败；可能发生部分写入，应检查文件和日志 |
| rolled_back | 离线文件已按备份恢复，并记录回滚状态 |

运行资料默认位于仓库目录，Git 已忽略它们：

| 位置 | 内容 |
|---|---|
| backups/ | 修改前的 MPP 备份 |
| mcp/changesets/CS-*.json | 离线预览、批准、输出与验证记录 |
| mcp/changesets/live/LIVE-*.json | 在线预览、前后内存状态和执行结果 |
| logs/ | 离线操作与验证日志 |
| mcp/work/ | 临时读取文件、锁和测试副本 |

在线恢复：优先使用当前 Project 的原生撤销记录，确认选中了本次 MCP 撤销组；不要在继续手动编辑后盲目撤销。需要文件恢复时，先保存另一个副本保护后续工作，关闭 Project，再人工核对并恢复备份。

离线回滚：将 {"change_set_id":"实际CS-ID"} 保存为参数文件，调用 project_rollback_changes。只有当前输出仍匹配写入后的哈希，且备份也匹配写入前哈希时才自动恢复；检测到后续修改或旧记录缺少必要哈希时会拒绝。

视图内部流变化会使离线结果降为警告。新建文件初始化也可能改变视图流，因此“字节变化”不等于已经证实图形丢失；应结合字段核验和 Project 内目视检查。不要只凭文件大小或工具退出码判断工程计划正确。

## 常见问题

| 现象 | 检查与处理 |
|---|---|
| No module named win32com / psutil | 使用安装依赖时同一个 Windows Python，优先用 .venv 的绝对路径 |
| JPype、JAR 或 MPXJ 无法加载 | 检查 MPXJ_LIB，确认依赖已安装到该目录且路径为纯 ASCII；仅设置变量不会自动安装依赖 |
| 找不到 jvm.dll | 检查 Java 21、JAVA_HOME、PATH，以及 Python/JVM 的位数是否兼容 |
| 没有发现 Project COM / GetActiveObject 失败 | 安装桌面版 Project；在线模式先手动打开计划；程序与 CLI 使用同一用户会话和相同权限级别 |
| 提示“当前活动项目不一致” | 切换到指定 MPP；核对传入路径，不要只核对窗口标题 |
| 提示“未保存” | Ctrl+S 后重新生成在线预览，无需关闭窗口 |
| 提示正在编辑或有对话框 | Enter 完成编辑、关闭菜单/对话框后重试 |
| 离线工具提示 Project 正在运行 | 关闭 Project；如果希望保留窗口，改用 project_live_* 在线工具 |
| 旧预览被拒绝 | 计划或进程发生变化，重新读取、重新预览并确认 |
| 在线调用超时 | 停止重复提交，检查 Project 与 LIVE 日志；已提交的 COM 调用可能已产生部分修改 |
| 字段返回值不同于请求 | 查看 checks 和 other_changes，Project 可能做了排程或格式调整；不要直接当作全部通过 |
| MCP 工具没有出现或仍是旧版本 | 核对配置中的绝对路径与解释器，检查 codex mcp list，重新连接服务或开启新会话 |
| 手动启动服务后没有界面 | stdio 服务等待 MCP 请求是正常现象；直接使用 CLI 子命令可得到输出 |

## 测试与更新

单元测试：

~~~powershell
.\.venv\Scripts\python.exe -m unittest discover -s mcp\tests -v
~~~

真实集成测试需要 Project、Java、MPXJ。先保存并关闭其他 Project 窗口：

~~~powershell
$env:MPXJ_LIB = 'C:\ProjectMCP\deps'
.\.venv\Scripts\python.exe mcp\tests\integration_smoke.py
.\.venv\Scripts\python.exe mcp\tests\live_integration.py
~~~

它们创建自己的 A/B、AA/BB 测试副本，验证写入、长备注字段、冲突拒绝、一次撤销及哈希回滚。在线集成测试会显示测试自有窗口，结束后关闭它；这不代表在线库会关闭用户窗口。不要将测试文件名中的时间戳当成正式计划版本。

标准样本为 mcp/tests/basic_plan.xml 和 mcp/tests/fixtures/A_B_AA_BB.mpp。需要再生成一个中性作者元数据的样本时，可运行：

~~~powershell
.\.venv\Scripts\python.exe mcp\tests\build_fixture.py --output C:\ProjectMCP\plans\new-fixture.mpp
~~~

目标必须不存在。若它已存在，请换文件名；不要删除自己的工程文件来通过检查。

更新仓库前先处理自己的本地改动，然后：

~~~powershell
git pull --ff-only
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
~~~

requirements-mpxj.txt 有变化时再更新 MPXJ 依赖。更新后重新连接 MCP；旧进程不会自动加载新代码。

## 项目结构与许可

| 文件 | 用途 |
|---|---|
| mcp/project_mcp_server.py | stdio MCP 服务与工具定义 |
| mcp/project_cli.py | CLI，共用 MCP 处理器 |
| mcp/live_project.py | 在线读取、预览和执行 |
| mcp/change_set.py | 离线变更集、验证与回滚 |
| mcp/note_fields.py | 备注字段解析和合并 |
| mcp/plan_builder.py | 基础 MPP 创建与独立核验 |
| scripts/com_modify.ps1、com_snapshot.ps1 | 原生 COM 写入及快照 |
| scripts/mpp_probe.py | MPXJ 读取与导出 |
| mcp/tests/ | 合成样本和测试 |

本项目按 **PolyForm Noncommercial License 1.0.0** 提供非商用许可，允许用途及非商业组织范围以 [LICENSE](LICENSE) 原文为准。超出许可允许用途的商业使用须取得维护者另行书面授权。第三方库和 Microsoft 产品适用各自许可。

本许可含非商用限制，请勿把本项目描述为允许自由商用的开源软件。分发时保留 LICENSE 及 [NOTICE](NOTICE) 的必要声明。

欢迎在 [GitHub Issues](https://github.com/nsforyou/project-mcp-for-codex/issues) 提交问题和建议。建议说明 Windows、Project、Python、Java 版本、使用的是在线还是离线模式、操作步骤及脱敏错误信息。请勿上传真实工程计划、账号凭据或未脱敏日志。

## 支持方式

[Buy Me a Coffee — 支持 OpenAI / 订阅 Codex](https://chatgpt.com/codex/pricing/) · [OpenAI 官网](https://openai.com/)

如果希望支持本项目背后的工具，可以前往官方页面选择适合自己的订阅方案。这里指向 OpenAI 官方订阅信息，不是个人收款账户；订阅不是本项目额外设置的强制条件，实际权益及使用条件以官方页面为准。

## Engineering Philosophy | 工程理念

**中文**

我们相信，人工智能不应该取代工程师的判断，而应该扩展工程师的能力。

工程师负责定义目标、建立约束、识别风险并验证结果；AI 负责处理繁琐、重复且容易出错的工作。

本项目尝试将自然语言、AI Agent 与 Microsoft Project 结合，让工程计划从编制、调整、校验到排版交付，逐步形成可追溯、可验证、可自动化的工作流程。

**AI 不替工程师承担责任，但可以帮助工程师完成更多工作。**

**English**

We believe artificial intelligence should not replace engineering judgment, but extend engineering capabilities.

Engineers define objectives, establish constraints, identify risks, and validate results. AI handles repetitive, time-consuming, and error-prone tasks.

This project explores the integration of natural language, AI agents, and Microsoft Project to build a traceable, verifiable, and increasingly automated workflow for project planning, scheduling, validation, formatting, and delivery.

**AI does not replace engineering accountability. It empowers engineers to accomplish more.**

---

*Developed through human–AI collaboration.*

*Author: JhonPrice*

> Built by an engineer, developed with AI, verified in the real world.
>
> 由工程师提出需求，与人工智能协作开发，并在真实工程实践中验证。
