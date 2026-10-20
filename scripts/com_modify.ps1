<#
.SYNOPSIS
    通过 Microsoft Project COM 原生修改 .mpp（Phase 1：验证视图层是否保留）。

.DESCRIPTION
    只做少量字段修改，绝不重建文件：
      复制源文件为工作副本 -> 打开副本 -> 按 UniqueID 定位任务 -> 修改字段 / 追加备注 ->
      另存为新 .mpp -> 关闭。全程记录 Before/After 快照与操作日志。

    安全约定（实测得出）：
      1) 绝不以读写方式打开用户的原文件——Project 即便不保存，也会回写约 76 字节
         文档元数据；因此所有写入都发生在临时工作副本上。
      2) 关闭一律使用 pjDoNotSave(0)。

    支持的操作：
      set_fields    设置任务字段（字段名用 Project 对象模型原名）
      append_note   以 [YYYY-MM-DD][来源][记录人] 格式追加备注（不覆盖原文）

.EXAMPLE
    powershell -File scripts/com_modify.ps1 -OpsJson data/com_ops_phase1.json
#>
param(
    [Parameter(Mandatory = $true)][string]$OpsJson,
    [switch]$KeepVisible,
    [switch]$NoWorkingCopy
)

$ErrorActionPreference = 'Stop'
$OpsJson = (Resolve-Path -LiteralPath $OpsJson).Path
$spec = Get-Content -LiteralPath $OpsJson -Raw -Encoding UTF8 | ConvertFrom-Json
$source = $spec.source
$saveAs = $spec.save_as
$logPath = $spec.log

function Get-Prop($obj, $name) {
    if ($null -eq $obj) { return $null }
    try { return $obj.$name } catch { return $null }
}
function Get-DateStr($value) {
    if ($null -eq $value) { return $null }
    if ($value -is [datetime]) { return $value.ToString('yyyy-MM-dd HH:mm') }
    $s = [string]$value
    if ($s -match '^\s*$') { return $null }
    return $s
}
function Find-TaskByUniqueId($proj, $uid) {
    foreach ($t in $proj.Tasks) {
        if ([int](Get-Prop $t 'UniqueID') -eq [int]$uid) { return $t }
    }
    return $null
}
function Find-TaskByRowId($proj, $rowId) {
    foreach ($t in $proj.Tasks) {
        if ([int](Get-Prop $t 'ID') -eq [int]$rowId) { return $t }
    }
    return $null
}
function Find-OpTask($proj, $op) {
    if ($null -ne $op.unique_id) { return Find-TaskByUniqueId $proj ([int]$op.unique_id) }
    if ($null -ne $op.row_id) { return Find-TaskByRowId $proj ([int]$op.row_id) }
    return $null
}
function Snapshot-Task($t) {
    return [ordered]@{
        unique_id        = [int](Get-Prop $t 'UniqueID')
        id               = [int](Get-Prop $t 'ID')
        wbs              = [string](Get-Prop $t 'WBS')
        name             = [string](Get-Prop $t 'Name')
        outline_level    = [int](Get-Prop $t 'OutlineLevel')
        start            = Get-DateStr (Get-Prop $t 'Start')
        finish           = Get-DateStr (Get-Prop $t 'Finish')
        duration         = [string](Get-Prop $t 'Duration')
        actual_start     = Get-DateStr (Get-Prop $t 'ActualStart')
        actual_finish    = Get-DateStr (Get-Prop $t 'ActualFinish')
        remaining        = [string](Get-Prop $t 'RemainingDuration')
        percent_complete = [double](Get-Prop $t 'PercentComplete')
        constraint_type  = [string](Get-Prop $t 'ConstraintType')
        constraint_date  = Get-DateStr (Get-Prop $t 'ConstraintDate')
        predecessors     = [string](Get-Prop $t 'Predecessors')
        successors       = [string](Get-Prop $t 'Successors')
        resource_names   = [string](Get-Prop $t 'ResourceNames')
        notes            = [string](Get-Prop $t 'Notes')
    }
}
function Set-TaskField($task, $field, $value, $defaultStart, $defaultFinish) {
    # 日期类字段必须带上项目默认的开始/完成时分，否则会把 08:00 / 17:00 写成 00:00
    $startFields = @('Start', 'ActualStart', 'ConstraintDate', 'Deadline', 'BaselineStart')
    $finishFields = @('Finish', 'ActualFinish', 'BaselineFinish')
    if ($startFields -contains $field -or $finishFields -contains $field) {
        $d = [datetime]::Parse([string]$value, [System.Globalization.CultureInfo]::InvariantCulture)
        if ([string]$value -match '[T ]\d{2}:\d{2}') {
            $task.$field = $d
            return
        }
        if ($finishFields -contains $field) {
            $task.$field = $d.Date.AddHours($defaultFinish.Hour).AddMinutes($defaultFinish.Minute)
        } else {
            $task.$field = $d.Date.AddHours($defaultStart.Hour).AddMinutes($defaultStart.Minute)
        }
        return
    }
    $task.$field = $value
}

$log = [ordered]@{
    timestamp   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    channel     = 'Microsoft Project COM (native .mpp modify)'
    source      = $source
    save_as     = $saveAs
    source_hash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    ops         = @()
    errors      = @()
}

if ($spec.expected_source_sha256 -and $log.source_hash -ne $spec.expected_source_sha256) {
    throw '源文件在排队等待期间发生变化，已拒绝打开写入，请重新预览。'
}

# 写入一律在工作副本上进行，保证源文件字节不变
$openPath = $source
$workingCopy = $null
if (-not $NoWorkingCopy) {
    $workingCopy = Join-Path $env:TEMP ("mcp_work_{0}.mpp" -f ([guid]::NewGuid().ToString('N').Substring(0, 8)))
    Copy-Item -LiteralPath $source -Destination $workingCopy -Force
    $openPath = $workingCopy
    $log.working_copy = $workingCopy
    Write-Host "[com] 已创建工作副本: $workingCopy"
}

Write-Host "[com] 启动 Microsoft Project ..."
$app = New-Object -ComObject MSProject.Application
try {
    if (-not $KeepVisible) { $app.Visible = $false }
    $app.DisplayAlerts = $false
    Write-Host "[com] 打开: $openPath"
    $app.FileOpenEx($openPath, $false) | Out-Null
    # Project COM 在 FileOpenEx 返回后可能仍在后台装载任务集合；若立刻读取会得到
    # ActiveProject.Tasks.Count=0。最多等待 30 秒，避免把“尚未装载”误判为无任务计划。
    $proj = $null
    $taskCount = 0
    for ($loadAttempt = 1; $loadAttempt -le 30; $loadAttempt++) {
        $proj = $app.ActiveProject
        $taskCollection = Get-Prop $proj 'Tasks'
        $taskCount = if ($null -eq $taskCollection) { 0 } else { [int]$taskCollection.Count }
        if ($taskCount -gt 0) { break }
        Start-Sleep -Seconds 1
    }
    if ($taskCount -le 0) { throw 'Project 已打开文件但任务集合在 30 秒内未完成装载。' }
    Write-Host "[com] 任务数: $taskCount"
    # 某些 Project 2019 COM 会话对 Project.DefaultStartTime/DefaultFinishTime 返回 null，
    # 即使计划已成功打开。任务日期无须改写时也不应因此阻断整个变更集；采用 Project
    # 标准日历的 08:00 / 17:00 作为安全回退，仅供日期字段自动补时分使用。
    $defaultStartValue = Get-Prop $proj 'DefaultStartTime'
    $defaultFinishValue = Get-Prop $proj 'DefaultFinishTime'
    $defaultStart = if ($null -eq $defaultStartValue) { (Get-Date).Date.AddHours(8) } else { [datetime]$defaultStartValue }
    $defaultFinish = if ($null -eq $defaultFinishValue) { (Get-Date).Date.AddHours(17) } else { [datetime]$defaultFinishValue }
    Write-Host ("[com] 默认工作时间: {0:HH:mm} ~ {1:HH:mm}" -f $defaultStart, $defaultFinish)

    foreach ($op in $spec.ops) {
        $applied = [ordered]@{}

        # 新建类操作没有目标 UniqueID，必须在定位任务之前处理
        if ($op.op -eq 'create_milestone') {
            try {
                $when = [datetime]::ParseExact([string]$op.date, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                $when = $when.Date.AddHours($defaultStart.Hour).AddMinutes($defaultStart.Minute)
                # Tasks.Add(Name, Before)：Before 指定插入到哪一行之前（比剪切/粘贴可靠）
                if ($null -ne $op.insert_before_row) {
                    $new = $proj.Tasks.Add([string]$op.name, [int]$op.insert_before_row)
                    $applied['InsertedBeforeRow'] = [int]$op.insert_before_row
                } else {
                    $new = $proj.Tasks.Add([string]$op.name)
                }
                $new.Milestone = $true
                $new.Start = $when
                $new.Finish = $when
                $new.Duration = 0
                $new.Estimated = $false
                $new.ConstraintType = [Microsoft.Office.Interop.MSProject.PjConstraint]::pjSNET
                $new.ConstraintDate = $when
                if ($null -ne $op.outline_level) {
                    # 缩进到父任务之下：位置 + 层级共同决定父级
                    $new.OutlineLevel = [int]$op.outline_level
                    $applied['OutlineLevel'] = [int]$op.outline_level
                }
                $applied['CreatedMilestone'] = [string]$op.name
                $applied['MilestoneDate'] = $when.ToString('yyyy-MM-dd')
                $applied['NewRowId'] = [int](Get-Prop $new 'ID')
                $applied['NewWbs'] = [string](Get-Prop $new 'WBS')
            } catch {
                $msg = "新建里程碑失败: $($_.Exception.Message)"
                Write-Host "[com][错误] $msg"
                $log.errors += $msg
            }
            $log.ops += [ordered]@{ op = $op.op; unique_id = $null; applied = $applied }
            Write-Host ("[com] 新建里程碑 {0} @ {1} 行号 {2} WBS {3}" -f $op.name, $applied['MilestoneDate'], $applied['NewRowId'], $applied['NewWbs'])
            continue
        }

        $uid = if ($null -ne $op.unique_id) { [int]$op.unique_id } else { [int]$op.row_id }
        $task = Find-OpTask $proj $op
        if ($null -eq $task) {
            $msg = "找不到任务（unique_id=$($op.unique_id) / row_id=$($op.row_id)）"
            Write-Host "[com][错误] $msg"
            $log.errors += $msg
            continue
        }
        $before = Snapshot-Task $task

        if ($op.op -eq 'set_fields') {
            foreach ($p in $op.fields.PSObject.Properties) {
                try {
                    Set-TaskField $task $p.Name $p.Value $defaultStart $defaultFinish
                    $applied[$p.Name] = $p.Value
                } catch {
                    $msg = "设置 $($p.Name) 失败: $($_.Exception.Message)"
                    Write-Host "[com][错误] $msg"
                    $log.errors += $msg
                }
            }
        }

        if ($op.op -eq 'shift_dates' -or $op.op -eq 'shift_task_dates') {
            # 整体平移并同步硬约束日期（否则 Project 打开后会按旧约束回拉）
            $days = [int]$op.days
            $start = [datetime](Get-Prop $task 'Start')
            $finish = [datetime](Get-Prop $task 'Finish')
            $newStart = $start.AddDays($days)
            $newFinish = $finish.AddDays($days)
            $task.Start = $newStart
            $task.Finish = $newFinish
            $ctype = [string](Get-Prop $task 'ConstraintType')
            if ($ctype -eq 'Finish No Earlier Than') { $task.ConstraintDate = $newFinish }
            elseif ($ctype -eq 'Start No Earlier Than') { $task.ConstraintDate = $newStart }
            $applied['ShiftDays'] = $days
            $applied['NewStart'] = $newStart.ToString('yyyy-MM-dd')
            $applied['NewFinish'] = $newFinish.ToString('yyyy-MM-dd')
            $applied['ConstraintRefreshed'] = $ctype
        }

        if ($op.op -eq 'set_predecessors') {
            # COM 的 Predecessors 是整串文本（"142FS-1d,148FS+0d"），必须一次性写入
            $task.Predecessors = [string]$op.predecessors
            $applied['Predecessors'] = [string]$op.predecessors
        }

        if ($op.append_note) {
            $old = [string](Get-Prop $task 'Notes')
            $line = [string]$op.append_note
            if ($old -and $old.Trim().Length -gt 0) {
                # 追加而不是覆盖
                $task.Notes = ($old.TrimEnd() + "`r`n" + $line)
            } else {
                $task.Notes = $line
            }
            $applied['Notes'] = $line
        }

        $after = Snapshot-Task $task
        $log.ops += [ordered]@{
            op        = $op.op
            unique_id = $uid
            name      = $after.name
            applied   = $applied
            before    = $before
            after     = $after
        }
        Write-Host ("[com] 已修改 UID={0} {1} | {2}" -f $uid, $after.name, ($applied.Keys -join ','))
    }

    Write-Host "[com] 重新计算 ..."
    try { $app.CalculateProject() | Out-Null } catch { Write-Host "[com][警告] CalculateProject 失败: $($_.Exception.Message)" }

    $saveAs = if ($spec.save_in_place) { $openPath } else { $saveAs }
    $modeText = if ($spec.save_in_place) { '就地保存(FileSave)' } else { '另存为(FileSaveAs)' }
    Write-Host "[com] $modeText : $saveAs"
    if (-not $spec.save_in_place -and (Test-Path -LiteralPath $saveAs)) { Remove-Item -LiteralPath $saveAs -Force }
    $format = 0   # pjMPP
    # 杀软实时扫描可能短暂占用刚写出的文件，保存与哈希读取都加重试
    $saved = $false
    for ($attempt = 1; $attempt -le 4 -and -not $saved; $attempt++) {
        try {
            if ($spec.save_in_place) {
                $app.FileSave() | Out-Null
            } else {
                $app.FileSaveAs($saveAs, $format) | Out-Null
            }
            $saved = $true
        } catch {
            Write-Host "[com][重试 $attempt] 保存失败: $($_.Exception.Message)"
            Start-Sleep -Seconds 2
        }
    }
    if (-not $saved) { throw "FileSaveAs 连续 4 次失败" }

    for ($attempt = 1; $attempt -le 6; $attempt++) {
        try {
            $log.saved_hash = (Get-FileHash -LiteralPath $saveAs -Algorithm SHA256).Hash
            $log.saved_size_kb = [math]::Round((Get-Item -LiteralPath $saveAs).Length / 1KB, 1)
            break
        } catch {
            Start-Sleep -Seconds 2
        }
    }
    if ($log.saved_hash) {
        Write-Host ("[com] 已保存 {0} ({1} KB) sha256={2}" -f $saveAs, $log.saved_size_kb, $log.saved_hash.Substring(0, 16))
    } else {
        Write-Host "[com][警告] 文件已保存但哈希读取失败"
    }
}
catch {
    $log.errors += $_.Exception.Message
    Write-Host "[com][异常] $($_.Exception.Message)"
}
finally {
    try { $app.FileCloseEx(0) } catch {}   # 0 = pjDoNotSave
    try { $app.Quit(0) } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
    [System.GC]::Collect()
    if ($workingCopy -and (Test-Path -LiteralPath $workingCopy)) {
        Remove-Item -LiteralPath $workingCopy -Force -ErrorAction SilentlyContinue
        Write-Host "[com] 工作副本已清理"
    }
}

$outDir = Split-Path -Parent $logPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
[System.IO.File]::WriteAllText($logPath, ($log | ConvertTo-Json -Depth 8), (New-Object System.Text.UTF8Encoding($false)))
Write-Host "[com] 操作日志: $logPath"
if ($log.errors.Count -gt 0) { exit 1 }
exit 0
