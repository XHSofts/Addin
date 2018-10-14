Module ModMain
    '===编译信息==='
    Public Const FinalBuildDate As String = "20181014"
    Public Const FinalBuildVersion As Integer = 26
    Public Const TodayBuildVersion As Integer = 3
    '===变量定义==='
    Public CurrErrCode As Integer

    '===常量定义==='
    Public Const TitleColor As ConsoleColor = ConsoleColor.Gray
    Public Const InfoColor As ConsoleColor = ConsoleColor.White
    Public Const InfoColor1 As ConsoleColor = ConsoleColor.Green
    Public Const InfoColor2 As ConsoleColor = ConsoleColor.Yellow
    Public Const InfoColor3 As ConsoleColor = ConsoleColor.Magenta
    Public Const TableColor As ConsoleColor = ConsoleColor.Cyan
    Public Const AlertColor As ConsoleColor = ConsoleColor.Red

    Function Main() As Integer

        CurrErrCode = 0

        '===使 Console 窗口能正确显示程序内容===
        If Console.WindowHeight < 30 Then
            Console.WindowHeight = 30
        End If
        If Console.WindowWidth < 125 Then
            Console.WindowWidth = 125
        End If
        If Console.BufferHeight < 9001 Then
            Console.BufferHeight = 9001
        End If
        If Console.BufferWidth < 125 Then
            Console.BufferWidth = 125
        End If

        '===若无参数，则显示自带帮助===
        If Command() = "" Or Command() = "/?" Then

            Output("=========================Addin 命令行程序=========================", 0, TitleColor)
            Output("                          Code By XiaoHe           ", 0, InfoColor)
            Output("                             帮助页面              ", 0, InfoColor1)
            Output("                    Version: " + My.Application.Info.Version.Major.ToString + "." + My.Application.Info.Version.Minor.ToString + "." + My.Application.Info.Version.Build.ToString + " Build:" + FinalBuildDate + "-" + TodayBuildVersion.ToString, 0, InfoColor2)
            Output("=========================Addin 命令行程序=========================", 0, TitleColor)
            Output("", 0, TitleColor)
            Output("本程序将提供更加完备的命令行支持。", 0, InfoColor)
            Output("", 0, TitleColor)
            Output("/? :显示本帮助。", 1, InfoColor1)
            Output("/help HelpContent :显示指定模块的详细帮助。", 1, InfoColor1)
            Output("", 0, InfoColor)
            Output("比如:", 1, InfoColor1)
            Output("/help About : 显示本程序关于(必看)。", 2, InfoColor2)
            Output("/help Math : 显示数学模块的详细帮助。", 2, InfoColor2)
            Output("/help File : 显示文件模块的详细帮助。", 2, InfoColor2)
            Output("/help Web : 显示网络模块的详细帮助。", 2, InfoColor2)
            Output("/help System : 显示系统模块的详细帮助。", 2, InfoColor2)
            Output("/help Graph : 显示图像模块的详细帮助。", 2, InfoColor2)
            Output("/help GUI : 显示图形界面模块的详细帮助。", 2, InfoColor2)
            Output("其中/help可简写为/h。", 1, InfoColor2)
            Output("", 0, InfoColor2)
            Output("/qerrcode ErrCode :根据输入的本程序提供的返回码查询对应的错误解释。", 1, InfoColor1)
            Return 0
            Exit Function
        End If


        Dim Args0 As String = My.Application.CommandLineArgs(0).ToLower.Replace("-", "/")   '===自动替换'-'为'/'===
        '===如果参数要求显示帮助，就直接显示帮助===
        If Args0 = "/help" Or Args0 = "/h" Then

            Try
                ShowHelp(My.Application.CommandLineArgs(1))
                Return CurrErrCode
            Catch ex As Exception   '===如果未输入help的参数===
                Select Case ex.GetType.ToString
                    Case "System.ArgumentOutOfRangeException"
                        Output("错误:未输入要查询帮助的模块名称！详细模块名称列表请使用 /? 参数获取。", 0， AlertColor)
                        Return 1
                    Case Else   '===如果是未知错误===
                        Output("发生错误", 0， AlertColor)
                        Output("错误类型:" + ex.GetType.ToString(), 1， InfoColor)
                        Output("错误描述:" + ex.Message, 1， InfoColor)
                        Output("其它信息:" + ex.Data.ToString, 1， InfoColor)
                        Output("最后处理函数:" + ex.TargetSite.ToString(), 1， InfoColor)
                        Output("请将本段文字复制，并反馈给作者。相关联系方式在/help About 中已写明，谢谢合作！", 0， InfoColor1)
                        Return 99
                End Select
            End Try
            Exit Function

        End If
        '===开始抓取可能的错误===
        Try

            '===判断参数===
            '===TODO:补充完整命令列表===
            Select Case Args0
                Case "/qerrcode"
                    Console.WriteLine(Query_Errcode(My.Application.CommandLineArgs(1)))
                Case "/math"
                    MathBas(My.Application.CommandLineArgs())
                Case Else
                    Output("错误:未找到命令！", 0， AlertColor)
                    CurrErrCode = 3
            End Select
            '===返回码===
            Return CurrErrCode

            '===如果出现错误===
        Catch ex As Exception

            '====分析错误类型===
            Select Case ex.GetType.ToString
                Case "System.ArgumentOutOfRangeException"
                    Output("错误:未输入相应命令所需的必要参数！", 0， AlertColor)
                    Return 1
                Case "System.OverflowException"
                    Output("错误:输入的值太大或太小，或无法计算。", 0， AlertColor)
                    Return 2
                Case "System.InvalidCastException"
                    Output("错误:输入的值不是数字或格式错误。", 0， AlertColor)
                    Return 2
                Case Else   '===如果是未知错误===
                    Output("发生错误", 0， AlertColor)
                    Output("错误类型:" + ex.GetType.ToString(), 1， InfoColor)
                    Output("错误描述:" + ex.Message, 1， InfoColor)
                    Output("其它信息:" + ex.Data.ToString, 1， InfoColor)
                    Output("最后处理函数:" + ex.TargetSite.ToString(), 1， InfoColor)
                    Output("请将本段文字复制，并反馈给作者。相关联系方式在/help About 中已写明，谢谢合作！", 0， InfoColor1)
                    Return 99
            End Select


        End Try
    End Function
    ''' <summary>  
    ''' 获取指定返回码对应的解释，并作为String返回
    ''' </summary> 
    ''' <param name="ErrCodeNum">指定的返回码</param>   
    ''' <returns>指定返回码对应的解释</returns>
    ''' <version> 
    ''' 1.0
    ''' </version> 
    Private Function Query_Errcode(ErrCodeNum As Integer) As String
        Select Case ErrCodeNum
            Case 0
                Return "错误解释:正常返回值，代表程序运行正常。"
                CurrErrCode = 0
            Case 1
                Return "错误解释:命令缺少必要参数或参数过多，无法执行。"
                CurrErrCode = 0
            Case 2
                Return "错误解释:命令参数格式、类型或范围错误，无法执行。"
                CurrErrCode = 0
            Case 3
                Return "错误解释:未找到对应命令，无法执行。"
                CurrErrCode = 0
            Case 99
                Return "错误解释:程序遇到未知错误，要根据程序输出的信息进行处理。"
                CurrErrCode = 0
            Case Else
                Return "错误:请确认输入的返回码是否正确！"
                CurrErrCode = 2
        End Select
    End Function
    ''' <summary>  
    ''' 获取指定模块对应的帮助，并输出到标准输出(Cout)上
    ''' </summary> 
    ''' <param name="HelpContent">需要提供帮助的模块</param>   
    ''' <version> 
    ''' 1.0
    ''' </version> 
    Private Sub ShowHelp(HelpContent As String)

        Dim HelpPart As String
        HelpPart = HelpContent.ToUpper
        Select Case HelpPart
            Case "ABOUT"
                Output("本程序于 2016 年 7 月 22 日最初由小何突发奇想编写，并初步完善。", 0， InfoColor)
                Output("", 0， InfoColor)
                Output("本程序旨在利用 VB.Net 的特性，完善 Console 命令行的功能，", 1， InfoColor1)
                Output("提供原先完全无法想像的功能。", 1， InfoColor1)
                Output("Under GNU GPLv3", 1， InfoColor2)
                Output("", 0， InfoColor)
                Output("本程序的几点注意:", 0， InfoColor)
                Output("1.所有命令及开关均不用注意大小写，会自动转换。", 1， InfoColor1)
                Output("2.本程序部分开关中的""/""可替换为""-""。(某些开关为避免与负数搞混而不会自动替换)", 1， InfoColor1)
                Output("3.部分开关都有位置限制，详情请参见各模块帮助。", 1， InfoColor1)
                Output("4.本程序为了兼容低版本系统，故使用.Net Framework 4.0。如果你能看到这条消息，起码你已经装上了.Net framework 4.0。", 1， InfoColor1)
                Output("5.你可以随意传播本程序，但请不要反编译，或用各种工具修改作者，谢谢合作。", 1， InfoColor1)
                Output("要说的就这么多，如果有任何 BUG 欢迎联系 @xh321 (Github)；xiaohe321@outlook.com (邮箱)；919897176 (QQ)，并注明提交 BUG，不要骚扰哦。", 1， InfoColor1)
                Output("Github 地址：https://github.com/XHSofts/Addin/", 1， InfoColor2)
                CurrErrCode = 0
            Case "MATH"
                Output("Addin 命令行程序下属 Math 模块:", 0， TitleColor)
                Output("", 0， InfoColor)
                Output("命令语法:", 0， InfoColor)
                Output("", 0， InfoColor)
                Output("普通函数：  Addin /math [/rnd|/sqr] [/int] [数字|下限 上限 [种子]]", 0， InfoColor)
                Output("", 0， InfoColor)
                Output("三角函数：  Addin /math [/sin|/cos|/tan|...] [/deg] 角度", 0， InfoColor)
                Output("", 0， InfoColor)
                Output("对数函数：  Addin /math [/lg|/ln|/log] [数字|底数 真数]", 0， InfoColor)
                Output("", 0， InfoColor)
                Output("注1:你可以使用 /deg 开关来指定使用角度模式(默认是弧度模式)， /int 开关来得到整数结果。", 1， InfoColor1)
                Output("附加注意: /deg 用于三角函数/双曲三角函数时是指你输入的是弧度还是角度；而用于反三角函数/反双曲三角函数时是指输出的是弧度还是角度。", 2， InfoColor2)
                Output("", 0， InfoColor)
                Output("注2:取随机数时，如启用 /int 开关取整，默认是没有意义的(随机数默认取0~1之间，取整后永远为0)，所以如果启用 /int 开关，随机数将自动乘100再进行取整。", 1， InfoColor1)
                Output("", 0， InfoColor)
                Output("注3:开关之间顺序可以颠倒，但是开关与参数之间的先后顺序不可颠倒（否则会产生不可预料的错误)，所有开关必须放在参数之前。", 1， InfoColor1)
                Output("", 0， InfoColor)
                Output("注4:使用时请注意定义域，相关定义域如下:", 1， InfoColor1)
                Output("Sqr                                   [0,+∞)", 1， TableColor)
                Output("Sin                                         R", 1， TableColor)
                Output("Cos                                         R", 1， TableColor)
                Output("Tan                   {θ|θ≠kπ+π/2，k∈Z}", 1， TableColor)
                Output("Cot                        {θ|θ≠kπ，k∈Z}", 1， TableColor)
                Output("Sec                   {θ|θ≠kπ+π/2，k∈Z}", 1， TableColor)
                Output("Csc                        {θ|θ≠kπ，k∈Z}", 1， TableColor)
                Output("HSin                                        R", 1， TableColor)
                Output("HCos                                        R", 1， TableColor)
                Output("HTan                                        R", 1， TableColor)
                Output("HCot                                 {x|x≠0}", 1， TableColor)
                Output("HSec                                        R", 1， TableColor)
                Output("HCsc                                 {x|x≠0}", 1， TableColor)
                Output("ArcSin                                [-1，1]", 1， TableColor)
                Output("ArcCos                                [-1，1]", 1， TableColor)
                Output("ArcTan                                      R", 1， TableColor)
                Output("ArcCot                                      R", 1， TableColor)
                Output("ArcSec                   （-∞，-1]U[1，+∞）", 1， TableColor)
                Output("ArcCsc                   （-∞，-1]U[1，+∞）", 1， TableColor)
                Output("HArcSin                                     R", 1， TableColor)
                Output("HArcCos                             [1，+∞）", 1， TableColor)
                Output("HArcTan                             （-1，1）", 1， TableColor)
                Output("HArcCot                （-∞，-1）U（1，+∞）", 1， TableColor)
                Output("HArcSec                               （0，1]", 1， TableColor)
                Output("HArcCsc                              {x|x≠0}", 1， TableColor)
                Output("Log类                              （0，+∞）", 1， TableColor)

                Output("", 0， InfoColor)
                Output("详细命令如下:", 0， InfoColor)
                Output("/rnd [/int] [下限 上限] [种子] :随机数", 1， InfoColor1)
                Output("解释:", 2， InfoColor2)
                Output("一个参数：参数作为种子；", 3， InfoColor3)
                Output("两个参数：参数作为下限和上限；", 3， InfoColor3)
                Output("三个参数：参数作为下限，上限和种子；", 3， InfoColor3)
                Output("/sqr [/int] 数字 : 平方根", 1， InfoColor1)
                Output("/sin [/deg] 角度 : 三角函数正弦", 1， InfoColor1)
                Output("/cos [/deg] 角度 : 三角函数余弦", 1， InfoColor1)
                Output("/tan [/deg] 角度 : 三角函数正切", 1， InfoColor1)
                Output("/cot [/deg] 角度 : 三角函数余切", 1， InfoColor1)
                Output("/sec [/deg] 角度 : 三角函数正割", 1， InfoColor1)
                Output("/csc [/deg] 角度 : 三角函数余割", 1， InfoColor1)
                Output("/hsin [/deg] 角度 : 三角函数双曲正弦", 1， InfoColor1)
                Output("/hcos [/deg] 角度 : 三角函数双曲余弦", 1， InfoColor1)
                Output("/htan [/deg] 角度 : 三角函数双曲正切", 1， InfoColor1)
                Output("/hcot [/deg] 角度 : 三角函数双曲余切", 1， InfoColor1)
                Output("/hsec [/deg] 角度 : 三角函数双曲正割", 1， InfoColor1)
                Output("/hcsc [/deg] 角度 : 三角函数双曲余割", 1， InfoColor1)
                Output("/arcsin [/deg] 角度 : 三角函数反正弦", 1， InfoColor1)
                Output("/arccos [/deg] 角度 : 三角函数反余弦", 1， InfoColor1)
                Output("/arctan [/deg] 角度 : 三角函数反正切", 1， InfoColor1)
                Output("/arccot [/deg] 角度 : 三角函数反余切", 1， InfoColor1)
                Output("/arcsec [/deg] 角度 : 三角函数反正割", 1， InfoColor1)
                Output("/arccsc [/deg] 角度 : 三角函数反余割", 1， InfoColor1)
                Output("/harcsin [/deg] 角度 : 三角函数反双曲正弦", 1， InfoColor1)
                Output("/harccos [/deg] 角度 : 三角函数反双曲余弦", 1， InfoColor1)
                Output("/harctan [/deg] 角度 : 三角函数反双曲正切", 1， InfoColor1)
                Output("/harccot [/deg] 角度 : 三角函数反双曲余切", 1， InfoColor1)
                Output("/harcsec [/deg] 角度 : 三角函数反双曲正割", 1， InfoColor1)
                Output("/harccsc [/deg] 角度 : 三角函数反双曲余割", 1， InfoColor1)
                Output("/lg  数字 : 以10为底的对数", 1， InfoColor1)
                Output("/ln  数字 : 以e为底的对数", 1， InfoColor1)
                Output("/log 底数 真数 : 对数", 1， InfoColor1)
                Output("", 0， InfoColor)
                Output("例子:", 0， InfoColor)
                Output("Addin /Math /Sin /Deg 30        返回结果:0.5", 0， InfoColor)
                CurrErrCode = 0
            Case Else
                Output("错误:请确认输入的模块名称是否有误！详细模块名称列表请使用 /? 参数获取。", 0， AlertColor)
                CurrErrCode = 2
        End Select
    End Sub
    ''' <summary>  
    ''' 输出指定内容到标准输出(Cout)上
    ''' </summary> 
    ''' <param name="Data">需要输出的数据</param>   
    ''' <param name="Indenting">前面需补齐的空格数</param>   
    ''' <param name="Color">输出的字体颜色</param>   
    ''' <version> 
    ''' 1.0
    ''' </version> 
    Public Sub Output(ByVal Data As String, ByVal Indenting As Integer, ByVal Color As System.ConsoleColor)
        Data = New String(" ", Indenting * 4) & Data

        Console.ForegroundColor = Color
        Console.WriteLine(Data)
        Console.ForegroundColor = ConsoleColor.White

        ' Dim Writer As New IO.StreamWriter(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\Wolfram Alpha wrapper log.log", True)
        ' Writer.WriteLine(Data)
        ' Writer.Close()
        ' Writer.Dispose()

    End Sub
End Module
