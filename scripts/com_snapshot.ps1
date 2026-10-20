<#
.SYNOPSIS
    采集 .mpp 的 COM 指纹（任务数据 + 文件级视图层集合），输出 UTF-8 JSON。

.DESCRIPTION
    以只读方式打开 Microsoft Project 文件，采集：
      1) 项目信息与当前视图/表格/筛选/分组
      2) 文件级视图层集合：Views / ViewsSingle / ViewsCombination / TaskTables /
         ResourceTables / TaskFilters / ResourceFilters / TaskGroups2 / Reports，
         以及各 *-List 名称表（含内置项）
      3) 日历、资源
      4) 每个任务的完整字段（含自定义字段 Text/Number/Date/Flag）
    用于 COM 原生修改前后的逐项对比。

.EXAMPLE
    powershell -File scripts/com_snapshot.ps1 -MppPath "a.mpp" -Out "data/a.json"
#>
param(
    [Parameter(Mandatory = $true)][string]$MppPath,
    [Parameter(Mandatory = $true)][string]$Out,
    [switch]$KeepVisible
)

$ErrorActionPreference = 'Stop'
$MppPath = (Resolve-Path -LiteralPath $MppPath).Path
$outDir = Split-Path -Parent $Out
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$Out = [System.IO.Path]::GetFullPath($Out)

# --- 容错取值 ---------------------------------------------------------------
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
function Get-Num($value) {
    if ($null -eq $value) { return $null }
    try { return [double]$value } catch { return $null }
}
function Get-Bool($value) {
    if ($null -eq $value) { return $null }
    try { return [bool]$value } catch { return $null }
}
function Get-Str($obj, $name) {
    $v = Get-Prop $obj $name
    if ($null -eq $v) { return $null }
    return [string]$v
}
function Get-CollectionNames($collection) {
    $result = @()
    if ($null -eq $collection) { return $result }
    $count = 0
    try { $count = [int]$collection.Count } catch { return $result }
    for ($i = 1; $i -le $count; $i++) {
        $name = '<读取失败>'
        try { $name = [string]$collection.Item($i).Name } catch { }
        $result += [pscustomobject]@{ index = $i; name = $name }
    }
    return $result
}
function Get-StringList($list) {
    $result = @()
    if ($null -eq $list) { return $result }
    foreach ($v in $list) { $result += [string]$v }
    return $result
}

Write-Host "[com] 启动 Microsoft Project ..."
$app = New-Object -ComObject MSProject.Application
try {
    if (-not $KeepVisible) { $app.Visible = $false }
    $app.DisplayAlerts = $false
    Write-Host "[com] 打开(只读): $MppPath"
    $app.FileOpenEx($MppPath, $true) | Out-Null
    $proj = $app.ActiveProject
    $cal0 = Get-Prop $proj 'Calendar'

    Write-Host "[com] 采集项目信息与视图层集合 ..."
    $projectInfo = [ordered]@{
        name             = Get-Str $proj 'Name'
        title            = Get-Str $proj 'Title'
        start            = Get-DateStr (Get-Prop $proj 'ProjectStart')
        finish           = Get-DateStr (Get-Prop $proj 'ProjectFinish')
        current_view     = Get-Str $proj 'CurrentView'
        current_table    = Get-Str $proj 'CurrentTable'
        current_filter   = Get-Str $proj 'CurrentFilter'
        current_group    = Get-Str $proj 'CurrentGroup'
        task_count       = [int](Get-Prop $proj 'Tasks').Count
        resource_count   = [int](Get-Prop $proj 'Resources').Count
        calendar_count   = [int](Get-Prop $proj 'BaseCalendars').Count
        default_calendar = Get-Str $cal0 'Name'
    }
    $viewLayer = [ordered]@{
        views             = Get-CollectionNames (Get-Prop $proj 'Views')
        views_single      = Get-CollectionNames (Get-Prop $proj 'ViewsSingle')
        views_combination = Get-CollectionNames (Get-Prop $proj 'ViewsCombination')
        task_tables       = Get-CollectionNames (Get-Prop $proj 'TaskTables')
        resource_tables   = Get-CollectionNames (Get-Prop $proj 'ResourceTables')
        task_filters      = Get-CollectionNames (Get-Prop $proj 'TaskFilters')
        resource_filters  = Get-CollectionNames (Get-Prop $proj 'ResourceFilters')
        task_groups       = Get-CollectionNames (Get-Prop $proj 'TaskGroups2')
        reports           = Get-CollectionNames (Get-Prop $proj 'Reports')
        view_list         = Get-StringList (Get-Prop $proj 'ViewList')
        task_table_list   = Get-StringList (Get-Prop $proj 'TaskTableList')
        task_filter_list  = Get-StringList (Get-Prop $proj 'TaskFilterList')
        task_group_list   = Get-StringList (Get-Prop $proj 'TaskGroupList')
        report_list       = Get-StringList (Get-Prop $proj 'ReportList')
    }

    Write-Host "[com] 采集日历 ..."
    $calendars = @()
    foreach ($cal in (Get-Prop $proj 'BaseCalendars')) {
        $calendars += [ordered]@{
            name = Get-Str $cal 'Name'
            guid = Get-Str $cal 'Guid'
        }
    }

    Write-Host "[com] 采集资源 ..."
    $resources = @()
    foreach ($res in (Get-Prop $proj 'Resources')) {
        $resources += [ordered]@{
            unique_id = Get-Num (Get-Prop $res 'UniqueID')
            id        = Get-Num (Get-Prop $res 'ID')
            name      = Get-Str $res 'Name'
            type      = Get-Str $res 'Type'
            max_units = Get-Num (Get-Prop $res 'MaxUnits')
        }
    }

    $taskCollection = Get-Prop $proj 'Tasks'
    $taskCount = [int]$taskCollection.Count
    Write-Host "[com] 采集任务字段（$taskCount 条）..."
    $tasks = @()
    $i = 0
    foreach ($t in $taskCollection) {
        $i++
        if ($i % 50 -eq 0) { Write-Host "      ... $i / $taskCount" }
        $tasks += [ordered]@{
            unique_id          = Get-Num (Get-Prop $t 'UniqueID')
            id                 = Get-Num (Get-Prop $t 'ID')
            wbs                = Get-Str $t 'WBS'
            outline_number     = Get-Str $t 'OutlineNumber'
            outline_level      = Get-Num (Get-Prop $t 'OutlineLevel')
            name               = Get-Str $t 'Name'
            summary            = Get-Bool (Get-Prop $t 'Summary')
            milestone          = Get-Bool (Get-Prop $t 'Milestone')
            manual             = Get-Bool (Get-Prop $t 'Manual')
            estimated          = Get-Bool (Get-Prop $t 'Estimated')
            start              = Get-DateStr (Get-Prop $t 'Start')
            finish             = Get-DateStr (Get-Prop $t 'Finish')
            duration           = Get-Str $t 'Duration'
            actual_start       = Get-DateStr (Get-Prop $t 'ActualStart')
            actual_finish      = Get-DateStr (Get-Prop $t 'ActualFinish')
            actual_duration    = Get-Str $t 'ActualDuration'
            remaining_duration = Get-Str $t 'RemainingDuration'
            percent_complete   = Get-Num (Get-Prop $t 'PercentComplete')
            physical_percent   = Get-Num (Get-Prop $t 'PhysicalPercentComplete')
            constraint_type    = Get-Str $t 'ConstraintType'
            constraint_date    = Get-DateStr (Get-Prop $t 'ConstraintDate')
            deadline           = Get-DateStr (Get-Prop $t 'Deadline')
            critical           = Get-Bool (Get-Prop $t 'Critical')
            total_slack        = Get-Str $t 'TotalSlack'
            free_slack         = Get-Str $t 'FreeSlack'
            predecessors       = Get-Str $t 'Predecessors'
            successors         = Get-Str $t 'Successors'
            resource_names     = Get-Str $t 'ResourceNames'
            calendar           = Get-Str $t 'Calendar'
            baseline_start     = Get-DateStr (Get-Prop $t 'BaselineStart')
            baseline_finish    = Get-DateStr (Get-Prop $t 'BaselineFinish')
            baseline_duration  = Get-Str $t 'BaselineDuration'
            baseline_cost      = Get-Num (Get-Prop $t 'BaselineCost')
            notes              = Get-Str $t 'Notes'
            custom             = [ordered]@{
                text1   = Get-Str $t 'Text1'
                text2   = Get-Str $t 'Text2'
                text3   = Get-Str $t 'Text3'
                text5   = Get-Str $t 'Text5'
                text10  = Get-Str $t 'Text10'
                number1 = Get-Num (Get-Prop $t 'Number1')
                number2 = Get-Num (Get-Prop $t 'Number2')
                date1   = Get-DateStr (Get-Prop $t 'Date1')
                flag1   = Get-Bool (Get-Prop $t 'Flag1')
            }
        }
    }

    $snapshot = [ordered]@{
        captured_at  = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        source_file  = $MppPath
        file_size_kb = [math]::Round((Get-Item -LiteralPath $MppPath).Length / 1KB, 1)
        sha256       = (Get-FileHash -LiteralPath $MppPath -Algorithm SHA256).Hash
        project      = $projectInfo
        view_layer   = $viewLayer
        calendars    = $calendars
        resources    = $resources
        tasks        = $tasks
    }

    $json = $snapshot | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($Out, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "[com] 指纹已写入: $Out"
    Write-Host "[com] 任务 $($tasks.Count) / 资源 $($resources.Count) / 日历 $($calendars.Count)"
    Write-Host ("[com] 视图 {0} / 单视图 {1} / 组合视图 {2} / 任务表 {3} / 任务筛选 {4} / 任务分组 {5} / 报表 {6}" -f `
        $viewLayer.views.Count, $viewLayer.views_single.Count, $viewLayer.views_combination.Count, `
        $viewLayer.task_tables.Count, $viewLayer.task_filters.Count, $viewLayer.task_groups.Count, $viewLayer.reports.Count)
}
finally {
    try { $app.FileCloseEx(0) } catch {}
    try { $app.Quit(0) } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch {}
    [System.GC]::Collect()
}
exit 0
