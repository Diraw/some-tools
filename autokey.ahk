; 设置脚本运行速度和精度
SetBatchLines -1
SetMouseDelay -1
SetKeyDelay -1
CoordMode, Pixel, Screen ; 像素坐标相对于屏幕
CoordMode, Mouse, Screen ; 鼠标坐标相对于屏幕

; 全局变量，用于存储找到的图像位置
AddButtonX =
AddButtonY =
XAxisX =
XAxisY =

; 操作基础延迟 (所有自动化操作都将使用这个延迟)
; 注意：这个值将根据GUI中的“超快、快、中”选择或自定义速度选择动态设置
CurrentSleepBase := 0 ; 初始化为0，将在StartAutomation中设置

; 目标窗口标题
WinTitle = Data Extraction

; 图像文件路径 (假设在同一目录下)
AddButtonImage = add_button.png
XAxisImage = x_axis.png

; 添加一个全局变量来标记是否应该退出自动化
global ShouldExitAutomation := false
; 新增：自动化开始时间戳
global AutomationStartTime := 0

; 新增：日志文件路径
global LogFilePath := ""
; 新增：控制是否启用调试模式（true为启用，false为禁用）
global EnableDebugMode := false ; 默认禁用调试模式，因为日志会持续写入文件

; 新增：控制是否启用自定义速度
global EnableCustomSpeed := false ; 默认禁用自定义速度
global CustomSpeedValue := 0 ; 默认自定义速度值

; --- 主 GUI ---
Gui, Add, Text, , 起始值:
Gui, Add, Edit, vStartValue gValidateInput,
Gui, Add, Text, , 结束值:
Gui, Add, Edit, vEndValue gValidateInput,
Gui, Add, Text, , 步长:
Gui, Add, Edit, vStep gValidateInput,

; 速度选择部分 (现在只显示预设速度下拉列表，自定义速度移到配置页面)
Gui, Add, Text, , 自动化速度:
Gui, Add, DropdownList, vSelectedSpeed Choose2, 超快(10ms)|快(50ms)|中(100ms) ; Choose2 默认选择第二个选项 ("快(50ms)")

Gui, Add, Button, xm gFindImages, 查找图像
Gui, Add, Button, xm gStartAutomation, 开始自动化

; **移除 OutputText 编辑框**
; Gui, Add, Edit, vOutputText ReadOnly W400 H150, ; 用于显示输出的文本框

Gui, Add, Text, xm cRed, 提示：按 Esc 键可随时退出自动化 ; 使用 xm 左对齐
Gui, Add, Button, xm gShowConfig, 其他 ; 新增按钮，用于打开配置页面
Gui, Show, , 自动化数据输入

; 设置Esc键为退出自动化的热键
Hotkey, Esc, ExitAutomation, On

; --- 配置 GUI ---
Gui, 2:New, +Owner1 ; 创建第二个GUI实例，并设置为主GUI的子窗口
Gui, 2:Add, Checkbox, vEnableCustomSpeedChecked gToggleCustomSpeed, 自定义速度
Gui, 2:Add, Text, x+10 yp, 数值(ms):
Gui, 2:Add, Edit, vCustomSpeedValueEdit w80 gValidateInput Disabled, ; 默认禁用
Gui, Add, Text, xm cRed, 提示：自定义速度会覆盖外层选择 ; 使用 xm 左对齐
Gui, 2:Add, Checkbox, xm vEnableDebugModeChecked, 调试模式 ; 使用 xm 左对齐
Gui, 2:Add, Button, xm gShowUpdates, 版本历史 ; 使用 xm 左对齐
Gui, 2:Show, Hide, 其他 ; 默认隐藏

; 在脚本启动时，根据全局变量初始化配置GUI的控件状态
GuiControl, 2:, % (EnableCustomSpeed ? "Check" : "Uncheck"), EnableCustomSpeedChecked
GuiControl, 2:, % (EnableCustomSpeed ? "Enable" : "Disable"), CustomSpeedValueEdit
GuiControl, 2:Text, CustomSpeedValueEdit, %CustomSpeedValue% ; 使用 Text 子命令设置文本
GuiControl, 2:, % (EnableDebugMode ? "Check" : "Uncheck"), EnableDebugModeChecked

; 验证输入是否为数字的函数
ValidateInput:
    Gui, Submit, NoHide ; 主GUI的输入验证
    Gui, 2:Submit, NoHide ; 配置GUI的输入验证

    ; 辅助函数：检查是否为有效的数字（包括浮点数）
    IsNumeric(value) {
        if (value = "")
            return true
        ; 允许负数，允许小数点
        return RegExMatch(value, "^-?\d*\.?\d*$")
    }

    ; 检查起始值
    If (!IsNumeric(StartValue))
    {
        MsgBox, 48, 输入错误, 起始值必须是数字！
        GuiControl, , StartValue,
        Return
    }
    ; 检查结束值
    If (!IsNumeric(EndValue))
    {
        MsgBox, 48, 输入错误, 结束值必须是数字！
        GuiControl, , EndValue,
        Return
    }
    ; 检查步长
    If (!IsNumeric(Step))
    {
        MsgBox, 48, 输入错误, 步长必须是数字！
        GuiControl, , Step,
        Return
    }

    ; 检查自定义速度 (从配置GUI中获取值)
    If (EnableCustomSpeedChecked && !IsNumeric(CustomSpeedValueEdit))
    {
        MsgBox, 48, 输入错误, 自定义速度必须是数字！
        GuiControl, 2:Text, CustomSpeedValueEdit, ; 修正为使用 Text 子命令清空内容
        Return
    }
Return

; 切换自定义速度输入框和下拉列表状态的函数 (现在作用于配置GUI)
ToggleCustomSpeed:
    Gui, 2:Submit, NoHide ; 获取 EnableCustomSpeedChecked 的当前状态
    If (EnableCustomSpeedChecked)
    {
        GuiControl, 2:Enable, CustomSpeedValueEdit ; 启用自定义速度输入框
        GuiControl, 2:Focus, CustomSpeedValueEdit ; 将焦点设置到自定义速度输入框
    }
    Else
    {
        GuiControl, 2:Disable, CustomSpeedValueEdit ; 禁用自定义速度输入框
    }
Return

; 显示配置GUI的函数
ShowConfig:
    ; 在显示配置GUI之前，确保其控件状态与全局变量同步
    GuiControl, 2:, % (EnableCustomSpeed ? "Check" : "Uncheck"), EnableCustomSpeedChecked
    GuiControl, 2:, % (EnableCustomSpeed ? "Enable" : "Disable"), CustomSpeedValueEdit
    GuiControl, 2:Text, CustomSpeedValueEdit, %CustomSpeedValue% ; 使用 Text 子命令设置文本
    GuiControl, 2:, % (EnableDebugMode ? "Check" : "Uncheck"), EnableDebugModeChecked

    Gui, 2:Show, , 其他
Return

; 查找图像的函数
FindImages:
    ; 先获取配置GUI的值，确保 EnableDebugMode 是最新的
    Gui, 2:Submit, NoHide ; 获取配置GUI中的输入值 (包括调试模式复选框)
    global EnableDebugMode := EnableDebugModeChecked ; 更新全局变量

    ; 如果调试模式启用，则初始化日志文件
    If (EnableDebugMode && LogFilePath = "") ; 只有在启用且尚未初始化时才初始化
        InitLogFile()

    Output("--- 开始查找图像 ---") ; 添加一个分隔符，方便区分不同次操作的输出

    ; 查找绿色加号按钮
    Output("正在查找图像: " . AddButtonImage)
    ImageSearch, FoundAddButtonX, FoundAddButtonY, 0, 0, A_ScreenWidth, A_ScreenHeight, %AddButtonImage%
    If ErrorLevel = 0
    {
        AddButtonX := FoundAddButtonX + 10 ; 调整到图像中心附近 (估算值，可能需要微调)
        AddButtonY := FoundAddButtonY + 10 ; 调整到图像中心附近 (估算值，可能需要微调)
        Output("找到图像: " . AddButtonImage . "，位置: " . FoundAddButtonX . "," . FoundAddButtonY)
        Output("图像中心估算坐标: " . AddButtonX . "," . AddButtonY)
    }
    Else
    {
        AddButtonX =
        AddButtonY =
        Output("未找到图像: " . AddButtonImage)
    }

    ; 查找x_axis位置
    Output("正在查找图像: " . XAxisImage)
    ImageSearch, FoundXAxisX, FoundXAxisY, 0, 0, A_ScreenWidth, A_ScreenHeight, %XAxisImage%
    If ErrorLevel = 0
    {
        XAxisX := FoundXAxisX + 10 ; 调整到图像中心附近 (估算值，可能需要微调)
        XAxisY := FoundXAxisY + 10 ; 调整到图像中心附近 (估算值，可能需要微调)
        Output("找到图像: " . XAxisImage . "，位置: " . FoundXAxisX . "," . FoundXAxisY)
        Output("图像中心估算坐标: " . XAxisX . "," . XAxisY)
    }
    Else
    {
        XAxisX =
        XAxisY =
        Output("未找到图像: " . XAxisImage)
    }

    If (AddButtonX <> "" && XAxisX <> "")
    {
        Output("已找到所有必要的图像。")
        MsgBox, 64, 查找完成, 已成功找到所有必要的图像。
    }
    Else
    {
        Output("未找到所有所需的图像。")
        MsgBox, 48, 查找未完成, 未找到所有所需的图像。请确保目标窗口可见且图像文件正确。
    }
    Output("--- 查找图像结束 ---")

    GuiControl, Enable, StartAutomation ; 启用开始按钮
    GuiControl, Enable, FindImages ; 启用查找按钮
Return

; 开始自动化的函数
StartAutomation:
    Gui, Submit, NoHide ; 获取主GUI中的输入值
    ; 先获取配置GUI的值，确保 EnableDebugMode 和自定义速度是最新的
    Gui, 2:Submit, NoHide ; 获取配置GUI中的输入值 (包括调试模式复选框和自定义速度)
    global EnableDebugMode := EnableDebugModeChecked ; 更新全局变量
    global EnableCustomSpeed := EnableCustomSpeedChecked
    global CustomSpeedValue := CustomSpeedValueEdit

    global CurrentSleepBase ; 声明 CurrentSleepBase 为全局变量，以便在 Output 中访问
    global PostAutomationDelay ; 声明 PostAutomationDelay 为全局变量

    ; 如果调试模式启用，则初始化日志文件
    If (EnableDebugMode && LogFilePath = "") ; 只有在启用且尚未初始化时才初始化
        InitLogFile()

    Output("GUI提交后，SelectedSpeed的值是: " . SelectedSpeed)
    Output("全局 EnableCustomSpeed 的值是: " . EnableCustomSpeed)
    Output("全局 CustomSpeedValue 的值是: " . CustomSpeedValue)


    ; 判断是否使用自定义速度
    If (EnableCustomSpeed)
    {
        ; 验证自定义速度值是否有效 (在 SaveConfigAndClose 中已验证，这里再检查一次以防万一)
        If (!IsNumeric(CustomSpeedValue) || CustomSpeedValue < 0)
        {
            MsgBox, 48, 输入错误, 自定义速度必须是非负数字！请在“其他”中设置。
            Return
        }
        CurrentSleepBase := CustomSpeedValue
        Output("使用自定义速度: " . CustomSpeedValue . "ms")
    }
    Else
    {
        ; 根据下拉列表选择的速度设置自动化操作的基础延迟
        If (SelectedSpeed = "超快(10ms)")
        {
            CurrentSleepBase := 10   ; 超快模式下，操作延迟 10ms
        }
        Else If (SelectedSpeed = "快(50ms)")
        {
            CurrentSleepBase := 50  ; 快模式下，操作延迟 50ms
        }
        Else If (SelectedSpeed = "中(100ms)")
        {
            CurrentSleepBase := 100  ; 中模式下，操作延迟 100ms
        }
        Else ; 默认值，以防万一（例如，如果下拉列表为空或未选择）
        {
            CurrentSleepBase := 50  ; 默认快模式操作延迟
        }
        Output("使用预设速度: " . SelectedSpeed . " (" . CurrentSleepBase . "ms)")
    }

    ; 将自动化结束后的额外延迟设置为操作基础延迟的两倍
    PostAutomationDelay := CurrentSleepBase * 2

    Output("自动化操作基础延迟: " . CurrentSleepBase . "ms") ; 记录固定的操作延迟
    Output("自动化结束后额外延迟: " . PostAutomationDelay . "ms") ; 记录自动化结束后的延迟

    ; 验证输入
    ; 步长必须大于0
    If (Step <= 0)
    {
        MsgBox, 48, 输入错误, 步长必须大于 0！
        GuiControl, , Step,
        Return
    }

    ; 结束值必须大于起始值
    If (EndValue <= StartValue)
    {
        MsgBox, 48, 输入错误, 结束值必须大于起始值！
        GuiControl, , EndValue,
        Return
    }

    ; 检查图像是否已找到
    If (AddButtonX = "" || XAxisX = "")
    {
        MsgBox, 48, 错误, 未找到所有所需的图像，请先点击 '查找图像'。
        Return
    }

    GuiControl, Disable, StartAutomation ; 禁用开始按钮
    GuiControl, Disable, FindImages ; 禁用查找按钮
    GuiControl, Disable, SelectedSpeed ; 禁用速度下拉列表
    GuiControl, Disable, ShowConfig ; 禁用配置按钮

    ; 重置退出标志
    ShouldExitAutomation := false
    ; 记录自动化开始时间
    AutomationStartTime := A_TickCount

    Output("--- 开始自动化任务 ---")

    ; 查找并激活窗口
    IfWinExist, %WinTitle%
    {
        WinActivate ; 激活窗口
        Sleep, %CurrentSleepBase% * 2 ; 激活窗口可能需要更长的延迟 (这里仍然是 CurrentSleepBase 的两倍)
        Output("窗口 '" . WinTitle . "' 已激活")
    }
    Else
    {
        Output("错误: 未找到窗口 '" . WinTitle . "'")
        MsgBox, 48, 错误, 未找到窗口 '" . WinTitle . "'
        GuiControl, Enable, StartAutomation ; 启用开始按钮
        GuiControl, Enable, FindImages ; 启用查找按钮
        GuiControl, Enable, SelectedSpeed ; 启用速度下拉列表
        GuiControl, Enable, ShowConfig ; 启用配置按钮
        Return
    }

    Output("已找到所有必要的图像，继续执行自动化...")

    ; 计算输入框位置
    InputBoxX := AddButtonX
    InputBoxY := XAxisY

    ; 生成数字序列并输入
    CurrentValue := StartValue
    While (CurrentValue <= EndValue && !ShouldExitAutomation)
    {
        Output("正在输入数字: " . CurrentValue)
        ; 点击输入框
        Click, %InputBoxX%, %InputBoxY%
        Sleep, %CurrentSleepBase% ; 使用动态设置的基础延迟

        ; 检查是否应该退出
        if (ShouldExitAutomation)
        {
            Output("用户已退出自动化")
            break
        }

        ; 全选输入框内容
        Send, ^a ; Ctrl+A 全选
        Sleep, %CurrentSleepBase% ; 使用动态设置的基础延迟

        ; 检查是否应该退出
        if (ShouldExitAutomation)
        {
            Output("用户已退出自动化")
            break
        }

        ; 输入数字
        Send, %CurrentValue%
        Sleep, %CurrentSleepBase% ; 使用动态设置的基础延迟

        ; 检查是否应该退出
        if (ShouldExitAutomation)
        {
            Output("用户已退出自动化")
            break
        }

        ; 点击绿色加号按钮
        Click, %AddButtonX%, %AddButtonY%
        Sleep, %CurrentSleepBase% ; 使用动态设置的基础延迟

        ; 检查是否应该退出
        if (ShouldExitAutomation)
        {
            Output("用户已退出自动化")
            break
        }

        ; 增加步长
        CurrentValue += Step
    }

    ; 计算自动化用时
    AutomationEndTime := A_TickCount
    ElapsedTimeMs := AutomationEndTime - AutomationStartTime
    ElapsedTimeSec := ElapsedTimeMs / 1000
    ElapsedTimeMin := Floor(ElapsedTimeSec / 60)
    ElapsedTimeSecRemainder := Mod(ElapsedTimeSec, 60)

    ; 将 Round() 的结果存储到临时变量
    RoundedSeconds := Round(ElapsedTimeSecRemainder, 2)

    If (!ShouldExitAutomation)
    {
        Output("数字输入完成！")
        Output("自动化总用时: " . ElapsedTimeMin . " 分 " . RoundedSeconds . " 秒")
        MsgBox, 64, 完成, 数字输入完成！`n自动化总用时: %ElapsedTimeMin% 分 %RoundedSeconds% 秒
    }
    Else
    {
        Output("自动化已由用户终止")
        Output("自动化已运行: " . ElapsedTimeMin . " 分 " . RoundedSeconds . " 秒 (终止前)")
        MsgBox, 64, 终止, 自动化已由用户终止`n已运行: %ElapsedTimeMin% 分 %RoundedSeconds% 秒
    }

    ; 应用自动化结束后的额外延迟
    Output("执行自动化结束后的额外延迟: " . PostAutomationDelay . "ms")
    Sleep, %PostAutomationDelay%

    Output("--- 自动化任务结束 ---")

    ; 重新启用GUI控件
    GuiControl, Enable, StartAutomation ; 启用开始按钮
    GuiControl, Enable, FindImages ; 启用查找按钮
    GuiControl, Enable, SelectedSpeed ; 启用速度下拉列表
    GuiControl, Enable, ShowConfig ; 启用配置按钮
Return

; 退出自动化的函数
ExitAutomation:
    ShouldExitAutomation := true
    Output("正在停止自动化...")
Return

; 保存配置并关闭配置GUI的函数
SaveConfigAndClose()
{
    Gui, 2:Submit, NoHide ; 获取配置GUI中的所有值

    ; 更新全局变量
    global EnableCustomSpeed := EnableCustomSpeedChecked
    global CustomSpeedValue := CustomSpeedValueEdit
    global EnableDebugMode := EnableDebugModeChecked

    ; 验证自定义速度值
    If (EnableCustomSpeed && (!IsNumeric(CustomSpeedValue) || CustomSpeedValue < 0))
    {
        MsgBox, 48, 输入错误, 自定义速度必须是非负数字！
        GuiControl, 2:Text, CustomSpeedValueEdit, ; 修正为使用 Text 子命令清空内容
        Return
    }

    Output("配置已保存:")
    Output("  启用自定义速度: " . (EnableCustomSpeed ? "是" : "否"))
    Output("  自定义速度值: " . CustomSpeedValue . "ms")
    Output("  启用调试日志: " . (EnableDebugMode ? "是" : "否"))

    Gui, 2:Hide ; 隐藏配置GUI
Return
}

; 初始化日志文件的函数
InitLogFile()
{
    global LogFilePath
    ; 确保 logs 文件夹存在
    LogFolder := A_ScriptDir . "\logs"
    If (!InStr(FileExist(LogFolder), "D")) ; 如果 logs 文件夹不存在
    {
        FileCreateDir, %LogFolder%
        ; 这里不应该调用 Output，因为它依赖于 LogFilePath 已经设置
        ; Output("已创建日志文件夹: " . LogFolder)
    }

    FormatTime, CurrentDateTime,, yyyyMMdd_HHmmss ; 获取当前日期时间，格式为 年月日_时分秒
    LogFilePath := LogFolder . "\automation_log_" . CurrentDateTime . ".txt"
    FileAppend, ; 创建或清空文件
    (
    自动化日志 - %CurrentDateTime%
    ---------------------------------
    ), %LogFilePath%
    ; 这里可以调用 Output，因为 LogFilePath 已经设置
    ; 注意：如果 InitLogFile 是在 Output 函数内部被调用的，
    ; 并且 Output 函数本身又需要 LogFilePath，可能会导致无限循环或错误。
    ; 确保 InitLogFile 在首次调用 Output 之前被调用，或者 Output 在 LogFilePath 为空时只尝试一次 InitLogFile。
    ; 当前的逻辑是安全的，因为 Output 在 LogFilePath 为空时会尝试 InitLogFile，
    ; 并且 InitLogFile 本身不会调用 Output。
}

; 将文本添加到输出文本框的函数 (只写入文件，不再更新 GUI)
Output(text)
{
    global LogFilePath, EnableDebugMode

    If (!EnableDebugMode) ; 如果调试模式未启用，则不写入日志文件
        Return

    ; 确保 LogFilePath 已经初始化，否则尝试初始化
    If (LogFilePath = "")
    {
        InitLogFile()
        ; 如果 InitLogFile 仍然无法设置 LogFilePath (例如权限问题)，则退出
        If (LogFilePath = "")
            Return
    }

    FormatTime, CurrentTime,, HH:mm:ss ; 获取当前时间，格式为 小时:分钟:秒
    FileAppend, [%CurrentTime%] %text%`n, %LogFilePath%
}

; 显示更新内容的函数
ShowUpdates:
    ; 定义更新内容
    UpdatesText =
    (
    版本历史
    v1.1 - 2025年6月21日
        - 支持通过Esc键随时终止自动化
        - 修复在输入0或0.几时报错的bug
        - 调试信息输出到'logs'子文件夹
        - 增加自动化速度的选择和自定义速度

    v1.0 - 2025年5月5日
        - 初始版本发布
        - 支持指定起始值、结束值和步长进行自动化数据输入
        - 自动查找并激活目标窗口
    )

    ; 创建新的GUI窗口来显示更新内容
    Gui, 3:New, +Owner2 ; 使用第三个GUI实例，并设置为配置GUI的子窗口
    Gui, 3:Add, Edit, ReadOnly W450 H250 VUpdatesOutput, %UpdatesText%
    Gui, 3:Add, Link, , <a href="https://github.com/Diraw/some-tools">源代码</a> ; 可点击的链接
    Gui, 3:Show, , 版本
Return

; 关闭更新内容GUI的函数
3GuiClose:
    Gui, 3:Destroy ; 销毁第三个GUI实例
Return

; 关闭配置GUI的函数
2GuiClose:
    ; 当用户直接关闭配置窗口时，也执行保存操作
    SaveConfigAndClose()
Return

; 当主GUI关闭时退出脚本
GuiClose:
    ExitApp