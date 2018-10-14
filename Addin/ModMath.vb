Imports System.Math
Module ModMath
    Public Structure RndArgs            '===参数列表信息
        Public isInt As Boolean         '===是否包含取整开关
        Public ArgsCount As Integer     '===开关个数
    End Structure

    ''' <summary>  
    ''' 数学基础模块函数
    ''' </summary> 
    ''' <param name="CLA">输入的命令行参数</param>   
    ''' <version> 
    ''' 1.0
    ''' </version> 
    Public Sub MathBas(CLA As ObjectModel.ReadOnlyCollection(Of String))
        '===命令行标识===
        '0为模块命令
        '1为模块命令子命令
        '2为模块命令参数1
        '3为模块命令参数2
        '...
        'BASE-COUNT:1
        '        0     1      2     3      4
        'Addin /Math /CLAS1 /CLAS2 CLA(3) CLA(4)...

        Dim CLAS1 As String = CLA(1).ToLower.Replace("-", "/") '===处理开关
        Dim CLAS2 As String
        If CLA.Count > 2 Then
            CLAS2 = CLA(2).ToLower '.Replace("-", "/")
        End If
        Select Case CLAS1   '===第一个参数：模块命令
            Case "/rnd"     '===随机数函数
                '(CLA(3) - CLA(2) + 1) * Rnd(CLA(4)) + CLA(2).ToString

                Dim CurrArgs As RndArgs = ParseRNDCLA(CLA) '===获取输入的参数
                Dim ArgStartPos As Integer = 0
                ArgStartPos = 2 + CurrArgs.ArgsCount

                If ArgStartPos = CLA.Count Then         '===无参数，即直接取随机数
                    Randomize()
                    Console.WriteLine(IIf(CurrArgs.isInt, Round(Rnd() * 100), Rnd()).ToString) '===如果包含取整开关，则结果乘100再取整，下同
                ElseIf ArgStartPos = CLA.Count - 1 Then '===一个参数，即包含种子的取随机数
                    Randomize()
                    Console.WriteLine(IIf(CurrArgs.isInt, Round(Rnd(CLA(ArgStartPos)) * 100), Rnd(CLA(ArgStartPos))).ToString)
                ElseIf ArgStartPos = CLA.Count - 2 Then '===两个参数，即有上下限的取随机数
                    If CLA(CLA.Count - 1) <> "" And CLA(CLA.Count - 2) <> "" Then
                        If Val(CLA(CLA.Count - 1)) < Val(CLA(CLA.Count - 2)) Then
                            Console.WriteLine("错误:上限小于下限！")
                            CurrErrCode = 2
                            Exit Select
                        End If

                        Randomize()
                        If CurrArgs.isInt Then
                            Console.WriteLine(Round((CLA(CLA.Count - 1) - CLA(CLA.Count - 2) + 1) * Rnd() + CLA(CLA.Count - 2)).ToString)
                        Else
                            Console.WriteLine((CLA(CLA.Count - 1) - CLA(CLA.Count - 2) + 1) * Rnd() + CLA(CLA.Count - 2).ToString)
                        End If
                        CurrErrCode = 0
                    End If
                ElseIf ArgStartPos = CLA.Count - 3 Then '===三个参数，即包含上下限和种子的取随机数
                    If Val(CLA(CLA.Count - 2)) < Val(CLA(CLA.Count - 3)) Then
                        Console.WriteLine("错误:上限小于下限！")
                        CurrErrCode = 2
                        Exit Select
                    End If

                    If CLA(CLA.Count - 1) <> "" And CLA(CLA.Count - 2) <> "" And CLA(CLA.Count - 3) <> "" Then
                        Randomize()
                        If CurrArgs.isInt Then
                            Console.WriteLine(Round((CLA(CLA.Count - 2) - CLA(CLA.Count - 3) + 1) * Rnd(CLA.Count - 1) + CLA(CLA.Count - 3)).ToString)
                        Else
                            Console.WriteLine((CLA(CLA.Count - 2) - CLA(CLA.Count - 3) + 1) * Rnd(CLA.Count - 1) + CLA(CLA.Count - 3).ToString)
                        End If
                        CurrErrCode = 0
                    End If
                Else
                    Console.WriteLine("错误:参数个数过多！")
                    CurrErrCode = 1
                End If


            Case "/sqr"     '===平方根函数
                Dim CurrArgs As RndArgs = ParseRNDCLA(CLA) '===获取输入的参数'
                Dim ArgStartPos As Integer = 0
                ArgStartPos = 2 + CurrArgs.ArgsCount

                If CurrArgs.isInt Then
                    If Not CheckDefinitionArea("sqr", True, CLA(3)) Then Exit Select

                    Console.WriteLine(Int(Sqrt(CLA(3))).ToString)
                    CurrErrCode = 0

                Else
                    If Not CheckDefinitionArea("sqr", True, CLA(2)) Then Exit Select

                    Console.WriteLine(Sqrt(CLA(2)).ToString)
                    CurrErrCode = 0

                End If

            '===三角函数==='
            Case "/sin"
                If CLAS2 = "/deg" Then      '===如果要求是角度制，下同
                    Console.WriteLine(Sin((DegToRad(CLA(3)))).ToString)
                Else
                    Console.WriteLine(Sin(CLA(2)).ToString)
                End If
                CurrErrCode = 0
            Case "/cos"
                If CLAS2 = "/deg" Then
                    Console.WriteLine(Cos((DegToRad(CLA(3)))).ToString)
                Else
                    Console.WriteLine(Cos(CLA(2)).ToString)
                End If
                CurrErrCode = 0
            Case "/tan"
                If CLAS2 = "/deg" Then
                    If (CLA(3)) Mod 90 <> 0 Then
                        Console.WriteLine(Tan((DegToRad(CLA(3)))).ToString)
                        CurrErrCode = 0
                    ElseIf (CLA(3)) Mod 180 = 0 Then
                        Console.WriteLine("0")
                        CurrErrCode = 0
                    Else
                        Console.WriteLine("错误:角度模式下Tan函数参数不能为90度的2倍数！")
                        CurrErrCode = 2
                    End If
                Else
                    If (CLA(2)) Mod (PI / 2) <> 0 Then
                        Console.WriteLine(Tan(CLA(2)).ToString)
                        CurrErrCode = 0
                    ElseIf (CLA(2)) Mod PI = 0 Then
                        Console.WriteLine("0")
                        CurrErrCode = 0
                    Else
                        Console.WriteLine("错误:弧度模式下Tan函数参数不能为pi/2的2倍数！")
                        CurrErrCode = 2
                    End If
                End If
            Case "/cot"
                If CLAS2 = "/deg" Then
                    If (CLA(3)) Mod 180 <> 0 Then
                        Console.WriteLine((1 / (Tan((DegToRad(CLA(3)))))).ToString)
                        CurrErrCode = 0
                    Else
                        Console.WriteLine("错误:角度模式下Cot函数参数不能为180度的倍数！")
                        CurrErrCode = 2
                    End If
                Else
                    If (CLA(2)) Mod PI <> 0 Then
                        Console.WriteLine((1 / (Tan(CLA(2)))).ToString)
                        CurrErrCode = 0
                    Else
                        Console.WriteLine("错误:弧度模式下Cot函数参数不能为pi的倍数！")
                        CurrErrCode = 2
                    End If
                End If
            Case "/sec"
                If CLAS2 = "/deg" Then
                    If (CLA(3)) Mod 90 <> 0 Then
                        Console.WriteLine((1 / (Cos((DegToRad(CLA(3)))))).ToString)
                        CurrErrCode = 0
                    Else
                        Console.WriteLine("错误:角度模式下Sec函数参数不能为90度的倍数！")
                        CurrErrCode = 2
                    End If
                Else
                    If (CLA(2)) Mod (PI / 2) <> 0 Then
                        Console.WriteLine((1 / (Cos(CLA(2)))).ToString)
                        CurrErrCode = 0
                    Else
                        Console.WriteLine("错误:弧度模式下Sec函数参数不能为pi/2的倍数！")
                        CurrErrCode = 2
                    End If
                End If
            Case "/csc"
                If CLAS2 = "/deg" Then
                    If (CLA(3)) Mod 180 <> 0 Then
                        Console.WriteLine((1 / (Sin((DegToRad(CLA(3)))))).ToString)
                        CurrErrCode = 0
                    Else
                        Console.WriteLine("错误:角度模式下Csc函数参数不能为180度的倍数！")
                        CurrErrCode = 2
                    End If
                Else
                    If (CLA(2)) Mod PI <> 0 Then
                        Console.WriteLine((1 / (Sin(CLA(2)))).ToString)
                        CurrErrCode = 0
                    Else
                        Console.WriteLine("错误:弧度模式下Csc函数参数不能为pi的倍数！")
                        CurrErrCode = 2
                    End If
                End If

            '===双曲三角函数==='
            Case "/hsin"
                If CLAS2 = "/deg" Then
                    Console.WriteLine(HSin(RadToDeg(CLA(3))).ToString)

                Else
                    Console.WriteLine(HSin(CLA(2)).ToString)

                End If
                CurrErrCode = 0
            Case "/hcos"
                If CLAS2 = "/deg" Then
                    Console.WriteLine(HCos(RadToDeg(CLA(3))).ToString)

                Else
                    Console.WriteLine(HCos(CLA(2)).ToString)

                End If
                CurrErrCode = 0
            Case "/htan"
                If CLAS2 = "/deg" Then
                    Console.WriteLine(HTan(RadToDeg(CLA(3))).ToString)

                Else
                    Console.WriteLine(HTan(CLA(2)).ToString)

                End If
                CurrErrCode = 0
            Case "/hcot"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("hcot", True, CLA(3)) Then Exit Select
                    Console.WriteLine(HCot(RadToDeg(CLA(3))).ToString)

                Else
                    If Not CheckDefinitionArea("hcot", False, CLA(2)) Then Exit Select
                    Console.WriteLine(HCot(CLA(2)).ToString)

                End If
                CurrErrCode = 0
            Case "/hsec"
                If CLAS2 = "/deg" Then
                    Console.WriteLine(HSec(RadToDeg(CLA(3))).ToString)

                Else
                    Console.WriteLine(HSec(CLA(2)).ToString)

                End If
                CurrErrCode = 0
            Case "/hcsc"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("hcsc", True, CLA(3)) Then Exit Select
                    Console.WriteLine(HCsc(RadToDeg(CLA(3))).ToString)

                Else
                    If Not CheckDefinitionArea("hcsc", False, CLA(2)) Then Exit Select
                    Console.WriteLine(HCsc(CLA(2)).ToString)

                End If
                CurrErrCode = 0

            '===反三角函数==='
            Case "/arcsin"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("arcsin", True, CLA(3)) Then Exit Select     '===检查定义域，防止报错，下同
                    Console.WriteLine(RadToDeg(Asin(CLA(3))).ToString)

                Else
                    If Not CheckDefinitionArea("arcsin", False, CLA(2)) Then Exit Select
                    Console.WriteLine(Asin(CLA(2)).ToString)

                End If
                CurrErrCode = 0
            Case "/arccos"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("arccos", True, CLA(3)) Then Exit Select
                    Console.WriteLine(RadToDeg(Acos(CLA(3))).ToString)

                Else
                    If Not CheckDefinitionArea("arccos", False, CLA(2)) Then Exit Select
                    Console.WriteLine(Acos(CLA(2)).ToString)

                End If
                CurrErrCode = 0
            Case "/arctan"
                If CLAS2 = "/deg" Then
                    Console.WriteLine(RadToDeg(Atan(CLA(3)).ToString))
                Else
                    Console.WriteLine(Atan(CLA(2)).ToString)
                End If
                CurrErrCode = 0
            Case "/arccot"
                If CLAS2 = "/deg" Then
                    Console.WriteLine(RadToDeg(ArcCot(CLA(3))).ToString)
                Else
                    Console.WriteLine((ArcCot(CLA(2))).ToString)
                End If
                CurrErrCode = 0
            Case "/arcsec"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("arcsec", True, CLA(3)) Then Exit Select
                    Console.WriteLine(RadToDeg(ArcSec(CLA(3))).ToString)

                Else
                    If Not CheckDefinitionArea("arcsec", False, CLA(2)) Then Exit Select
                    Console.WriteLine((ArcSec(CLA(2))).ToString)

                End If
                CurrErrCode = 0

            Case "/arccsc"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("arccsc", True, CLA(3)) Then Exit Select
                    Console.WriteLine(RadToDeg(ArcCsc(CLA(3))).ToString)

                Else
                    If Not CheckDefinitionArea("arccsc", False, CLA(2)) Then Exit Select
                    Console.WriteLine((ArcCsc(CLA(2))).ToString)

                End If
                CurrErrCode = 0

            '===反双曲三角函数
            Case "/harcsin"
                If CLAS2 = "/deg" Then
                    Console.WriteLine(RadToDeg(HArcsin(CLA(3))).ToString)

                Else
                    Console.WriteLine(HArcsin(CLA(2)).ToString)

                End If
                CurrErrCode = 0
            Case "/harccos"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("harccos", True, CLA(3)) Then Exit Select

                    Console.WriteLine(RadToDeg(HArccos(CLA(3)).ToString))
                Else
                    If Not CheckDefinitionArea("harccos", False, CLA(2)) Then Exit Select

                    Console.WriteLine(HArccos(CLA(2)).ToString)
                End If
                CurrErrCode = 0
            Case "/harctan"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("harctan", True, CLA(3)) Then Exit Select
                    Console.WriteLine(RadToDeg(HArctan(CLA(3))).ToString)

                Else
                    If Not CheckDefinitionArea("harctan", False, CLA(2)) Then Exit Select
                    Console.WriteLine((HArctan(CLA(2))).ToString)

                End If
                CurrErrCode = 0
            Case "/harccot"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("harccot", True, CLA(3)) Then Exit Select
                    Console.WriteLine(RadToDeg(HArccot(CLA(3))).ToString)
                Else
                    If Not CheckDefinitionArea("harccot", False, CLA(2)) Then Exit Select
                    Console.WriteLine((HArccot(CLA(2))).ToString)
                End If
                CurrErrCode = 0
            Case "/harcsec"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("harcsec", True, CLA(3)) Then Exit Select
                    Console.WriteLine(RadToDeg(ArcCsc(CLA(3))).ToString)

                Else
                    If Not CheckDefinitionArea("harcsec", False, CLA(2)) Then Exit Select
                    Console.WriteLine((ArcCsc(CLA(2))).ToString)

                End If
                CurrErrCode = 0
            Case "/harccsc"
                If CLAS2 = "/deg" Then
                    If Not CheckDefinitionArea("harccsc", True, CLA(3)) Then Exit Select
                    Console.WriteLine(RadToDeg(ArcCsc(CLA(3))).ToString)

                Else
                    If Not CheckDefinitionArea("harccsc", False, CLA(2)) Then Exit Select
                    Console.WriteLine((ArcCsc(CLA(2))).ToString)

                End If
                CurrErrCode = 0

            '===其他函数==='
            Case "/lg"
                If Not CheckDefinitionArea("log", True, CLA(2)) Then Exit Select
                Console.WriteLine(Log10(CLA(2)).ToString)
                CurrErrCode = 0

            Case "/ln"
                If Not CheckDefinitionArea("log", True, CLA(2)) Then Exit Select

                Console.WriteLine(Log(CLA(2)).ToString)
                CurrErrCode = 0

            Case "/log"
                If Not CheckDefinitionArea("log", True, CLA(2)) Then Exit Select
                If Not CheckDefinitionArea("log", True, CLA(3)) Then Exit Select

                Console.WriteLine((Log(CLA(3), CLA(2))).ToString)
                CurrErrCode = 0

                '===输入错误==='
            Case Else
                Output("错误:未找到对应命令，无法执行！", 0， AlertColor)
                CurrErrCode = 3
        End Select
    End Sub

    Private Function CheckDefinitionArea(Func As String, isDeg As Boolean, X As Double) As Boolean      '===检查函数定义域
        Select Case Func
            Case "sqr"
                If X < 0 Then
                    Output("错误:非复数运算，根号内要大于0！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "log"
                If X <= 0 Then
                    Output("错误:Log类型函数的参数应该大于0！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "sec"
            Case "csc"
            Case "hcot"
                If X = 0 Then
                    Output("错误:HCot 函数定义域为{x|x≠0}！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "hcsc"
                If X = 0 Then
                    Output("错误:HCsc 函数定义域为{x|x≠0}！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "arcsin"
                If X > 1 Or X < -1 Then
                    Output("错误:Arcsin 函数定义域为[-1，1]！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "arccos"
                If X > 1 Or X < -1 Then
                    Output("错误:Arccos 函数定义域为[-1，1]！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "arcsec"
                If X > -1 And X < 1 Then
                    Output("错误:Arcsec 函数定义域为（-∞，-1]U[1，+∞）！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "arccsc"
                If X > -1 And X < 1 Then
                    Output("错误:Arccsc 函数定义域为（-∞，-1]U[1，+∞）！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "harccos"
                If X < 1 Then
                    Output("错误:HArccos 函数定义域为 [1，+∞)！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "harctan"
                If X >= 1 Or X <= -1 Then
                    Output("错误:HArctan 函数定义域为（-1，1）！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "harccot"
                If X <= 1 And X >= -1 Then
                    Output("错误:HArccot 函数定义域为（-∞，-1）U（1，+∞）！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "harcsec"
                If X > 1 Or X <= 0 Then
                    Output("错误:HArcsec 函数定义域为（0，1]！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
            Case "harccsc"
                If X = 0 Then
                    Output("错误:HArccsc 函数定义域为{x|x≠0}！", 0, AlertColor)
                    CurrErrCode = 2
                    Return False
                Else
                    CurrErrCode = 0
                    Return True
                End If
        End Select
    End Function
    '===辅助函数==='
    Public Function DegToRad(DegNum As Double） As Double
        DegToRad = DegNum * (PI / 180)
    End Function
    Public Function RadToDeg(RadNum As Double） As Double
        RadToDeg = RadNum * (180 / PI)
    End Function

    Public Function ParseRNDCLA(CLA As ObjectModel.ReadOnlyCollection(Of String)) As RndArgs
        Dim MathArg As RndArgs              '===处理参数的函数==='
        MathArg.ArgsCount = 0
        MathArg.isInt = False

        For N = 2 To CLA.Count - 1          ';遍历所有参数
            Dim CurrCLA As String = CLA(N).ToLower '.Replace("-", "/")'处理大小写问题，全转换为小写
            If Left(CurrCLA, 1) = "/" Then  '如果是开关
                If CurrCLA = "/int" Then    '如果开关是要求取整
                    MathArg.isInt = True
                Else
                    Output("警告:未知开关:" + CurrCLA, 0, AlertColor)
                End If
                MathArg.ArgsCount += 1      '开关计数
            End If
        Next
        Return MathArg
    End Function



    '===其他辅助三角函数===
    Function ArcSec(X As Double) As Double  '===反正割
        ArcSec = Acos(1 / X)
    End Function
    Function ArcCsc(X As Double) As Double  '===反余割
        ArcCsc = Asin(1 / X)
    End Function
    Function ArcCot(X As Double) As Double  '===反余切
        ArcCot = Atan(1 / X)
    End Function
    Function HSin(X As Double) As Double    '===双曲正弦
        HSin = (Exp(X) - Exp(-X)) / 2
    End Function
    Function HCos(X As Double) As Double    '===双曲余弦
        HCos = (Exp(X) + Exp(-X)) / 2
    End Function
    Function HTan(X As Double) As Double    '===双曲正切
        HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
    End Function
    Function HSec(X As Double) As Double    '===双曲正割
        HSec = 2 / (Exp(X) + Exp(-X))
    End Function
    Function HCsc(X As Double) As Double    '===双曲余割 
        HCsc = 2 / (Exp(X) - Exp(-X))
    End Function
    Function HCot(X As Double) As Double    '===双曲余切
        HCot = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
    End Function
    Function HArcsin(X As Double) As Double '===反双曲正弦
        HArcsin = Log(X + Sqrt(X * X + 1))
    End Function
    Function HArccos(X As Double) As Double '反双曲余弦
        HArccos = Log(X + Sqrt(X * X - 1))
    End Function
    Function HArctan(X As Double) As Double '===反双曲正切
        HArctan = Log((1 + X) / (1 - X)) / 2
    End Function
    Function HArcsec(X As Double) As Double '===反双曲正割
        HArcsec = Log((Sqrt(-X * X + 1) + 1) / X)
    End Function
    Function HArccsc(X As Double) As Double '===反双曲余割
        HArccsc = Log((Sign(X) * Sqrt(X * X + 1) + 1) / X)
    End Function
    Function HArccot(X As Double) As Double '===反双曲余切
        HArccot = Log((X + 1) / (X - 1)) / 2
    End Function
End Module
