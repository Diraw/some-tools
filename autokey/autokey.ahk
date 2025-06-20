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

; 目标窗口标题
WinTitle = Data Extraction

; 图像文件路径 (假设在同一目录下)
AddButtonImage = add_button.png
XAxisImage = x_axis.png

; 添加一个全局变量来标记是否应该退出自动化
global ShouldExitAutomation := false

; GUI 变量
Gui, Add, Text, , 起始值:
Gui, Add, Edit, vStartValue gValidateInput,
Gui, Add, Text, , 结束值:
Gui, Add, Edit, vEndValue gValidateInput,
Gui, Add, Text, , 步长:
Gui, Add, Edit, vStep gValidateInput,
Gui, Add, Button, gFindImages, 查找图像
Gui, Add, Button, gStartAutomation, 开始自动化
Gui, Add, Edit, vOutputText ReadOnly W400 H150, ; 用于显示输出的文本框
Gui, Add, Text, cRed, 提示：按 Esc 键可随时退出自动化
Gui, Add, Button, gShowUpdates, 更新内容 ; 更新内容按钮
Gui, Show, , 自动化数据输入

; 设置Esc键为退出自动化的热键
Hotkey, Esc, ExitAutomation, On

; 验证输入是否为数字的函数
ValidateInput:
    Gui, Submit, NoHide

    ; 辅助函数：检查是否为有效的数字（包括浮点数）
    ; 这个函数会更宽松，允许空字符串，并且对浮点数有更好的支持
    IsNumeric(value) {
        ; 如果为空字符串，认为是有效的（允许用户清空输入）
        if (value = "")
            return true
        ; 尝试将值转换为数字。如果转换失败，说明不是数字。
        ; A_IsNumber 变量在表达式中会返回 1 (true) 或 0 (false)
        ; 如果值包含非数字字符（除了一个小数点），则不是数字
        ; RegExMatch 检查是否为整数或浮点数（允许负号和小数点）
        ; ^-?\d*\.?\d*$ 匹配：
        ; ^      - 字符串开始
        ; -?     - 可选的负号
        ; \d*    - 零个或多个数字
        ; \.?    - 可选的小数点
        ; \d*    - 零个或多个数字
        ; $      - 字符串结束
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
Return

; 查找图像的函数
FindImages:
    ; 清空输出文本框
    GuiControl, , OutputText,
    GuiControl, Disable, StartAutomation ; 禁用开始按钮
    GuiControl, Disable, FindImages ; 禁用查找按钮

    Output("正在查找图像...")

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

    GuiControl, Enable, StartAutomation ; 启用开始按钮
    GuiControl, Enable, FindImages ; 启用查找按钮
Return

; 开始自动化的函数
StartAutomation:
    Gui, Submit, NoHide ; 获取GUI中的输入值

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

    ; 清空输出文本框
    GuiControl, , OutputText,
    GuiControl, Disable, StartAutomation ; 禁用开始按钮
    GuiControl, Disable, FindImages ; 禁用查找按钮

    ; 重置退出标志
    ShouldExitAutomation := false

    Output("开始自动化任务...")
    ; Output("按 Esc 键可随时退出自动化") ; 这行可以保留在输出中，也可以只在GUI中显示

    ; 查找并激活窗口
    IfWinExist, %WinTitle%
    {
        WinActivate ; 激活窗口
        Sleep, 1000 ; 等待窗口激活
        Output("窗口 '" . WinTitle . "' 已激活")
    }
    Else
    {
        Output("错误: 未找到窗口 '" . WinTitle . "'")
        MsgBox, 48, 错误, 未找到窗口 '" . WinTitle . "'
        GuiControl, Enable, StartAutomation ; 启用开始按钮
        GuiControl, Enable, FindImages ; 启用查找按钮
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

        ; 检查是否应该退出
        if (ShouldExitAutomation)
        {
            Output("用户已退出自动化")
            break
        }

        ; 全选输入框内容
        Send, ^a ; Ctrl+A 全选

        ; 检查是否应该退出
        if (ShouldExitAutomation)
        {
            Output("用户已退出自动化")
            break
        }

        ; 输入数字
        Send, %CurrentValue%

        ; 检查是否应该退出
        if (ShouldExitAutomation)
        {
            Output("用户已退出自动化")
            break
        }

        ; 点击绿色加号按钮
        Click, %AddButtonX%, %AddButtonY%

        ; 检查是否应该退出
        if (ShouldExitAutomation)
        {
            Output("用户已退出自动化")
            break
        }

        ; 增加步长
        CurrentValue += Step
    }

    if (!ShouldExitAutomation)
    {
        Output("数字输入完成！")
        MsgBox, 64, 完成, 数字输入完成！
    }
    else
    {
        Output("自动化已由用户终止")
        MsgBox, 64, 终止, 自动化已由用户终止
    }

    GuiControl, Enable, StartAutomation ; 启用开始按钮
    GuiControl, Enable, FindImages ; 启用查找按钮
Return

; 退出自动化的函数
ExitAutomation:
    ShouldExitAutomation := true
    Output("正在停止自动化...")
Return

; 将文本添加到输出文本框的函数
Output(text)
{
    GuiControl, , OutputText, % OutputText . text . "`n"
    GuiControl, +VScroll, OutputText ; 确保滚动条在底部
}

; 检查变量是否为数字的函数
IsNumber(var)
{
    Return var is number
}

; 显示更新内容的函数
ShowUpdates:
    ; 定义更新内容
    UpdatesText =
    (
    更新日志
    v1.1 - 2025年6月20日
        - 支持通过Esc键随时终止自动化
        - 修复在输入0或0.几时报错的bug

    v1.0 - 2025年5月5日
        - 初始版本发布
        - 支持指定起始值、结束值和步长进行自动化数据输入
        - 自动查找并激活目标窗口
    )

    ; 创建新的GUI窗口来显示更新内容
    Gui, 2:New ; 使用第二个GUI实例
    Gui, 2:Add, Edit, ReadOnly W450 H250 VUpdatesOutput, %UpdatesText%
    Gui, 2:Add, Link, , <a href="https://github.com/Diraw/some-tools">源代码</a> ; 可点击的链接
    Gui, 2:Show, , 更新内容
Return

; 关闭更新内容GUI的函数
2GuiClose:
    Gui, 2:Destroy ; 销毁第二个GUI实例
Return


; 当主GUI关闭时退出脚本
GuiClose:
    ExitApp