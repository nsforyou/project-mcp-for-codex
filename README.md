# Project MCP for codex

作者：ChatGPT-6.1 Sol

许可：[PolyForm Noncommercial 1.0.0](LICENSE) · [提交 Issue](https://github.com/nsforyou/project-mcp-for-codex/issues)

Windows 本地 Microsoft Project MCP 服务与 CLI。支持读取计划、预览修改、备份后原生写入、结构化任务备注，以及保持 Project 打开的在线协作。

当前版本 **0.3.0**，**27 个 MCP 工具**。验证环境：Microsoft Project Professional 2019、Python 3.13、Java 21。执行 COM 操作需要本机安装 Microsoft Project。本仓库只含任务 A/B、资源组 AA/BB 的合成数据。

## 安装

```powershell
python -m pip install -r requirements.txt
# MPXJ 原生依赖须安装到纯 ASCII 路径，例如：
$env:MPXJ_LIB = 'C:\ProjectMCP\deps'
python -m pip install --target $env:MPXJ_LIB -r requirements-mpxj.txt
```

MPXJ 还需要 Java 21。设置 JAVA_HOME 或将 Java 加入 PATH。MPXJ_LIB 缺省为 `%LOCALAPPDATA%\ProjectMCP\deps`。本仓库不打包 Office 或第三方库。

## Codex 配置

将实际 Python 和仓库路径填入 MCP 配置：

```toml
[mcp_servers.microsoft-project]
command = "python"
args = ["C:\\your-path\\project-mcp-for-codex\\mcp\\project_mcp_server.py"]

[mcp_servers.microsoft-project.env]
MPXJ_LIB = "C:\\ProjectMCP\\deps"
PROJECT_COM_TIMEOUT = "120"
```

配置键 microsoft-project 保持兼容；服务标识为 project-mcp-for-codex。更新后刷新 MCP 连接。

## 开着 Project 讨论，确认后修改

1. 打开本地 MPP，使它成为活动项目。
2. `project_live_read(file)` 读取内存，供讨论，不改动计划。
3. 确定方案后保存当前工作、结束单元格编辑，保持窗口打开。
4. `project_live_preview(file, ops)` 返回 Before/Proposed 与 LIVE 变更 ID。
5. 用户确认后 `project_live_apply(change_id, approved_by)` 执行修改。
6. 用户在窗口内查看，Ctrl+S 保存，或通过撤销列表撤销整批“MCP 在线调整”。

示例 ops：

```json
[{"unique_id":1,"fields":{"Name":"A"},"note_fields":{"工程量":120,"单位":"米","当前状态":"施工中"}}]
```

在线支持 Name、Notes、Start、Finish、Duration（工作分钟）、PercentComplete、Predecessors、ResourceNames。摘要仅允许名称与备注；最多 500 任务、每批 20 项。资源名称修改不等于新增资源或调整人数容量。执行时请勿同时手动编辑。

只借用已有实例，不打开/关闭文档、不退出 Project、不发送键盘输入、不修改界面控制开关。执行前核对内存指纹和磁盘哈希，备份已保存文件，文件切换或内容变化时拒绝写入。默认不自动保存。20 秒超时只终止工作进程；已提交的 COM 调用不保证撤回，部分修改须人工核对，不能盲目重试。

## 离线修改、备注与新建

既有 MPP：`project_preview_changes` → 用户确认 → `project_apply_changes` → `project_verify_changes`。离线 COM 入口要求 Project 已关闭，并使用跨进程串行锁。默认备份后 FileSave 就地保存，FileSaveAs 可能损失已有 MPP 的视图对象。回滚检查当前文件与备份哈希，拒绝覆盖后续修改。

`project_create_from_xml` 默认预览，执行需匹配 source_sha256，目标必须不存在。基础创建器使用标准日历、手动排程，支持层级、时间、工期、容量、分配、依赖和里程碑。它不是原 MPP 的完整无损导入器；自定义日历、基线、实际进度、硬约束、成本等有明确限制。

备注按“字段名：值”解析。`project_get_task_note_fields` 支持单任务、批量读取及字段筛选；`project_preview_task_note_fields` 预览更新，保留其他说明。备注字段不会自动修改工期、完成率或资源分配。

## CLI

CLI 与 MCP 使用同一处理器，参数通过 UTF-8 JSON 文件传递：

```powershell
python mcp/project_cli.py tools
python mcp/project_cli.py call project_live_read --args-file read.json
```

read.json 示例：`{"file":"C:\\plans\\test.mpp"}`。

## 测试

唯一测试模型：任务 **A、B**，A → B 为 FS；资源组 **AA、BB**，容量分别 6、4；A 分配 AA 两人，B 分配 BB 四人。

- XML：`mcp/tests/basic_plan.xml`
- MPP：`mcp/tests/fixtures/A_B_AA_BB.mpp`

```powershell
python -m unittest discover -s mcp/tests -v
# 集成测试需要 Project、Java、MPXJ；先关闭其他 Project 窗口：
python mcp/tests/integration_smoke.py
python mcp/tests/live_integration.py
```

集成测试只创建自有副本，验证写入、长备注、冲突拒绝、单次撤销和回滚。运行时计划、变更集、日志及备份被 Git 忽略。

## 入口

- mcp/project_mcp_server.py：stdio MCP 服务。
- mcp/project_cli.py：CLI。
- mcp/live_project.py：在线协作。
- mcp/change_set.py：离线变更集与回滚。
- mcp/plan_builder.py：基础新建计划。
- scripts/com_modify.ps1、com_snapshot.ps1：COM 写入与快照。
- scripts/mpp_probe.py：MPXJ 独立读取与校验。

备份、日志默认保存在仓库本地 backups/、logs/、mcp/changesets/，不应提交到 Git。

## 许可与反馈

本项目公开源码，按 **PolyForm Noncommercial License 1.0.0** 提供非商用许可；允许的用途和非商业组织范围以 [LICENSE](LICENSE) 原文为准。超出该许可允许用途的商业使用，须事先取得维护者另行书面授权。第三方库和 Microsoft 产品适用各自许可，不因本项目公开而获得额外授权。

本许可包含非商用限制，不能将本项目描述为允许自由商用的开源软件。请在分发时保留 LICENSE 及 [NOTICE](NOTICE) 的必要声明。

欢迎通过 [GitHub Issues](https://github.com/nsforyou/project-mcp-for-codex/issues) 报告问题、提出功能建议或联系商业授权。请勿上传真实工程计划、账号凭据或未脱敏日志。
