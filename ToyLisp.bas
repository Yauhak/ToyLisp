Attribute VB_Name = "Module1"
'业余写的一个类似lisp语法的脚本语言的解释器
'业余写的一个类似Lisp语法的脚本语言的解释器
'名称：ToyLisp（真就是个玩具） 文件后缀：.LSP（我不是LSP）
'不过该有的东西，比如输入输出、条件循环、数学运算、数组操作、字符串操作、文件读写、公有变量、函数程序、错误提示等都有（^V^）
'语法很简单，就是([函数名/操作符] [参数列表])
'如(out 1 2 3 (+ 4 5))
'每一个程序都只有一个入口main。比如(main (out "Hello World!"))
'如果函数重复定义则以最先定义的函数为标准
'在写好你自己的ToyLisp程序后，可以在选择打开方式那选择ToyLisp.exe运行
'具体语法例子请看注释
'如有疏（ba）漏（ge）请多多包涵
'也感谢你的阅读（话说真有人愿意看完这一大坨代码吗（悲））
'（如果有人看到的话，玩玩也好呗，我就最会写解释器了）
'(。-ω-)zzz
'Yauhak Chen - QQ3953814837
Dim vars()
'变量
Dim funcs()
'定义函数
Dim sourcecode
'读取的源代码
Dim codelen, src, idx, curlv, ifabort, rt, ewhile
'源代码长度，源代码，当前解析位置，当前作用域等级，是否退出函数、循环，函数返回值，是否结束循环
Dim block, crlf '匹配括号，当前解析行数
'修改数组变量
Sub opar(tag, mat, ide, val)
    If ide = UBound(mat) Then
        tag(rec(mat(ide))) = val
    Else
        opar tag(rec(mat(ide))), mat, ide + 1, val
    End If
End Sub
'错误信息输出
'这个功能不完善
Sub cerr(exp)
    MsgBox exp, vbCritical, "ToyLisp"
    End
End Sub
Function rec(ary)
'执行被解析成列表的代码
On Error Resume Next
'不想写参数判断了，太累了
tmp = crlf
If IsArray(ary) Then
If UBound(ary) = -1 Then Exit Function
If UBound(ary) > 0 Then
If Not IsArray(ary(0)) Then
    operate_name = LCase(ary(0))
    If Not operate_name = "def" And Not operate_name = "do" And Not operate_name = "return" And Not operate_name = "m" _
    And Not operate_name = "array" And Not operate_name = "out" And Not operate_name = "in" And Not operate_name = "#" _
    And Not operate_name = "size" And Not operate_name = "m" Then fst_n = rec(ary(1))
    'fst_n 这个变量，我的初衷是记录数学运算语句的第一个值，方便+-*/这样的连续型数学运算，不然要定义好多遍
    Select Case operate_name
    Case "+":
        For i = 2 To UBound(ary)
            fst_n = 1 * fst_n + 1 * rec(ary(i))
        Next
    '(+ 1 2 3 (+ 4 5))
    Case "&":
        For i = 2 To UBound(ary)
            fst_n = fst_n & rec(ary(i))
        Next
    '(& "114" "514")
    '这个是字符串连接
    Case "-":
        For i = 2 To UBound(ary)
            fst_n = fst_n - rec(ary(i))
        Next
    '(- 1 2 3 (+ 4 5))
    Case "*":
        For i = 2 To UBound(ary)
            fst_n = fst_n * rec(ary(i))
        Next
    '(* 1 2 3 (+ 4 5))
    Case "/":
        For i = 2 To UBound(ary)
            fst_n = fst_n / rec(ary(i))
        Next
    '(/ 1 2 3 (+ 4 5))
    Case "and":
        For i = 2 To UBound(ary)
            fst_n = fst_n And rec(ary(i))
        Next
    Case "or":
        For i = 2 To UBound(ary)
            fst_n = fst_n Or rec(ary(i))
        Next
    '(and (or (= 1 1)(= 1 2))(= 2 2))
    Case "%":
        fst_n = fst_n Mod rec(ary(2))
    Case "^":
        fst_n = fst_n ^ rec(ary(2))
    '取余与乘方
    Case "sqrt": '开方
        fst_n = Sqr(fst_n)
    Case "sin":
        fst_n = Sin(fst_n)
    Case "cos":
        fst_n = Cos(fst_n)
    Case "tan":
        fst_n = Tan(fst_n)
    Case "atan":
        fst_n = Atn(fst_n)
    '三角函数
    Case "abs":
        fst_n = Abs(fst_n)
    Case "int":
        fst_n = Int(fst_n)
    Case "fix":
        fst_n = Fix(fst_n)
    'Int 和 Fix 都会删除 number 的小数部份而返回剩下的整数。
    'Int 和 Fix 的不同之处在于，如果 number 为负数，则 Int 返回小于或等于 number 的第一个负整数，
    '而 Fix 则会返回大于或等于 number 的第一个负整数。例如，Int 将 -8.4 转换成 -9，而 Fix 将 -8.4 转换成 -8。
    Case "sgn":
        fst_n = Sgn(fst_n)
    '>0 1;=0 0;<0 -1
    Case "log":
        fst_n = Log(fst_n)
    Case "rand":
        Randomize
        fst_n = Rnd() * fst_n
    Case "exp":
        fst_n = exp(fst_n)
    Case "round":
        fst_n = Round(fst_n)
    Case "=":
        fst_n = IIf(Trim(fst_n) = Trim(rec(ary(2))), 1, 0)
    Case ">":
        fst_n = IIf(1 * fst_n > 1 * rec(ary(2)), 1, 0)
    Case "<":
        fst_n = IIf(1 * fst_n < 1 * rec(ary(2)), 1, 0)
    Case ">=":
        fst_n = IIf(1 * fst_n >= 1 * rec(ary(2)), 1, 0)
    Case "<=":
        fst_n = IIf(1 * fst_n <= 1 * rec(ary(2)), 1, 0)
    Case "!":
        fst_n = IIf(fst_n <> rec(ary(2)), 1, 0)
    '条件判断
    Case "out":
        oput = ""
        For x = 1 To UBound(ary)
            oput = oput & rec(ary(x))
        Next
        MsgBox oput, , "ToyLisp"
        Exit Function
    '输出（不限参数）
    Case "in":
        fst_n = InputBox(rec(ary(1)), "ToyLisp")
    '输入
    Case "asc":
        fst_n = Asc(rec(ary(1)))
    Case "chr":
        fst_n = Chr(rec(ary(1)))
    'asc码与字符的相互转化
    Case "len":
        fst_n = Len(fst_n)
    '字符串长度
    Case "size":
        rec = UBound(rec(ary(1)))
        Exit Function
    '数组长度
    Case "split":
        rec = Split(fst_n, rec(ary(2)))
        Exit Function
    '字符串分割
    Case "substr":
        rec = Mid(fst_n, rec(ary(2)), rec(ary(3)))
        Exit Function
    '(substr "string" 1 1)从"string"字符串第一个位置截取一个长度为一的字符串
    Case "def":
        If UBound(vars) = 0 Then
            vars(0) = Array(ary(1), rec(ary(2)), curlv)
            ReDim Preserve vars(UBound(vars) + 1)
            Exit Function
        End If
        For v = 0 To UBound(vars) - 1
            If vars(v)(0) = ary(1) And (vars(v)(2) = curlv Or vars(v)(2) = -1) Then
                vars(v)(1) = rec(ary(2))
                Exit Function
            End If
        Next
        vars(UBound(vars)) = Array(ary(1), rec(ary(2)), curlv)
        ReDim Preserve vars(UBound(vars) + 1)
        Exit Function
    '定义一个变量
    Case "list":
        Dim llst()
        ReDim llst(0)
        llst(0) = fst_n
        For x = 2 To UBound(ary)
            ReDim Preserve llst(x - 1)
            llst(x - 1) = rec(ary(x))
        Next
        rec = llst
        Exit Function
    '返回一个列表（数组）
    '(def x (list 1 2 3 (list 4 5)))
    Case "m":
        If IsArray(rec(ary(1))) Then
            xp = rec(ary(1))
            For idx = 2 To UBound(ary)
                xp = xp(rec(ary(idx)))
            Next
            rec = xp: Exit Function
        End If
        For v = 0 To UBound(vars) - 1
            If vars(v)(0) = ary(1) And (vars(v)(2) = curlv Or vars(v)(2) = -1) Then
                xp = vars(v)(1)(rec(ary(2)))
                For idx = 3 To UBound(ary)
                    xp = xp(rec(ary(idx)))
                Next
                rec = xp
                Exit Function
            End If
        Next
    '返回数组内容
    '(def x (list 1 2 3 (list 4 5)))
    '(out (m x 3 0)),第三个参数及以后表下标
    Case "array":
        If IsArray(ary(1)) Then
            r = rec(ary(3))
            opar rec(ary(1)), ary(2), 0, r
            Exit Function
        End If
        For v = 0 To UBound(vars) - 1
            If vars(v)(0) = ary(1) And (vars(v)(2) = curlv Or vars(v)(2) = -1) Then
                r = rec(ary(3))
                opar vars(v)(1), ary(2), 0, r
                Exit Function
            End If
        Next
    '操作数组
    '(def x (list 1 2 3 (list 4 5)))
    '(array x (3 0) "H")，中间括号里的表下标
    Case "read":
        nHandle = FreeFile
        readt = ""
        Open fst_n For Input As #nHandle
        Do Until EOF(nHandle)
            Line Input #nHandle, newline
            readt = readt & newline & vbCrLf
        Loop
        Close nHandle
        rec = readt: Exit Function
    Case "outfile":
        nHandle = FreeFile
        Open fst_n For Output As #nHandle
        Print #nHandle, rec(ary(2))
        Close #nHandle
    '文件数据读写
    Case "alloc":
        Dim memspace()
        ReDim memspace(fst_n)
        rec = memspace
        Exit Function
    '返回一个指定大小的数组（元素都为空）
    '(def x (alloc 100))
    Case "do":
        For x = 1 To UBound(ary)
            todo = ary(x)
            Call rec(todo)
        Next
        Exit Function
    '这个可要可不要
    Case "if":
    If rec(ary(1)) = 1 Then
        Call rec(ary(2))
        Exit Function
    Else
        If UBound(ary) = 3 Then
            Call rec(ary(3))
        Else
            Exit Function
        End If
    End If
    'if([表达式][操作1][操作2（可选）])
    '注意操作表达式多余一个时要用do或括号括起来
    'while也一样
    '如(if(= 1 1)((out 1)(out"y"))(out 0))
    Case "while":
        Do While rec(ary(1)) = 1 And ifabort = False And ewhile = False
            rec (ary(2))
        Loop
        If ewhile = True Then ewhile = False
        Exit Function
    '(while (= 1 1)((out 1)(out 2)))
    Case "break":
        '跳出循环
        ewhile = True
        Exit Function
    Case "return":
        '返回自定义函数值
        rt = rec(ary(1))
        ifabort = True
        Exit Function
    Case "#": '注释（！！！注意！注释也被视为函数，要记得空格，且对于if和while语句你要小心！！！）
        Exit Function
    Case Else: '自定义函数
        If UBound(funcs) = 0 And IsEmpty(funcs(0)) Then cerr "警告：函数未定义" & vbCrLf & "描述：" & operate_name: Exit Function
        For Each fn In funcs
            If LCase(fn(1)) = operate_name Then
                bound = UBound(vars)
                If UBound(fn(2)) > 0 Or (UBound(fn(2)) = 0 And fn(2) <> "") Then
                    Dim newsp(): ReDim newsp(0)
                    For i = 2 To UBound(ary)
                        newsp(i - 2) = rec(ary(i))
                        ReDim Preserve newsp(i - 1)
                    Next
                    curlv = curlv + 1
                    vars(UBound(vars)) = Array(fn(2)(0), fst_n, curlv)
                    ReDim Preserve vars(UBound(vars) + 1)
                    For i = 1 To UBound(fn(2))
                        vars(UBound(vars)) = Array(fn(2)(i), newsp(i - 1), curlv)
                        ReDim Preserve vars(UBound(vars) + 1)
                    Next
                End If '自定义函数传参，curlv是当前作用域等级，防止参数访问越界
                '！！！注意：对于无参数的函数应在函数名后面随意写一个参数（"()"也行，比如"(break())"）,否则会被解释器认为是变量
                For exec = 3 To UBound(fn)
                    this = rec(fn(exec))
                    If ifabort = True Then ifabort = False: rec = rt: rt = "": Exit For
                Next
                ReDim Preserve vars(bound) '将函数程序产生的变量销毁
                curlv = curlv - 1
                Exit Function
            End If
         Next
         cerr "警告：函数未定义" & vbCrLf & "描述：" & operate_name
    End Select
    End If
    If IsArray(ary(0)) Then
        For x = 0 To UBound(ary)
            todo = ary(x)
            Call rec(todo)
        Next
        Exit Function
    End If '括号括起来的一系列表达式的执行
    rec = fst_n
Else
    rec = rec(ary(0))
    Exit Function
End If
Else
    If Mid(ary, 1, 1) > "9" Or Mid(ary, 1, 1) < "0" And Not Mid(ary, 1, 1) = "-" And Not Mid(ary, 1, 1) = Chr(34) And Not ary = "" Then
        For v = 0 To UBound(vars) - 1
            x = vars(v)
            If x(0) = ary And (x(2) = curlv Or x(2) = -1) Then
                If IsArray(vars(v)(1)) = False Then
                    rec = Replace(vars(v)(1), Chr(34), "")
                Else
                    rec = vars(v)(1)
                End If
                Exit Function
            End If
        Next
        cerr "警告：变量未定义" & vbCrLf & "描述：" & ary
    Else
        rec = Replace(ary, Chr(34), "")
        Exit Function
    End If
End If '变量、字符串、数值的返回
End Function
Sub parse()
'将源代码解析成列表
    src = sourcecode
    tm = Len(src)
    If tm = 0 Then Exit Sub
    Do While Mid(src, tm, 1) = Chr(13) Or Mid(src, tm, 1) = Chr(10) Or Mid(src, tm, 1) = " " Or Mid(src, tm, 1) = " "
        tm = tm - 1
    Loop
    '这似乎是一个bug，必须把源代码结尾的crlf与空格去掉，否则解析会出问题
    codelen = tm
    r = recursion()
    If Not IsArray(r(0)) Then Exit Sub
    'If ifrun = False Then ifrun = True: Exit Sub
    For Each blocks In r
    If UBound(blocks) > 0 Then
        If blocks(0) = "#" Then
        End If
        If LCase(blocks(0)) = "fn" Then
            funcs(UBound(funcs)) = blocks
            ReDim Preserve funcs(UBound(funcs) + 1)
        End If
        If LCase(blocks(0)) = "public" Then
            c = curlv
            curlv = -1
            For x = 1 To UBound(blocks) - 1 Step 2
                todo = Array("def", blocks(x), blocks(x + 1))
                Call rec(todo)
            Next
            curlv = c
        End If
    End If
    Next '公有变量
    If UBound(funcs) > 0 Then ReDim Preserve funcs(UBound(funcs) - 1)
    For Each blocks In r
    If UBound(blocks) > 0 Then
        If LCase(blocks(0)) = "main" Then
            For x = 1 To UBound(blocks)
                state = rec(blocks(x))
            Next
            Exit Sub
        End If
    End If
    Next '主函数
End Sub
'这便是解析函数了
'将源代码解析成列表（其实就是数组）
Function recursion()
    tmp = crlf
    Dim par()
    ReDim Preserve par(0)
    Do While idx <= codelen
        Do While Mid(src, idx, 1) = " " Or Mid(src, idx, 1) = " " Or Mid(src, idx, 1) = Chr(10) Or Mid(src, idx, 1) = Chr(13)
            If Mid(src, idx, 1) = Chr(13) Then crlf = crlf + 1
            idx = idx + 1
        Loop
        If Mid(src, idx, 1) = "(" Then
        tmp = crlf
            idx = idx + 1
            block = block + 1
            par(UBound(par)) = recursion()
            If Not IsArray(par(UBound(par))(0)) Then
                If (par(UBound(par))(0) = "fn" Or par(UBound(par))(0) = "main") And block > 0 Then cerr "错误：在语句块中定义函数" & vbCrLf & "位于行：" & tmp
            End If
            ReDim Preserve par(UBound(par) + 1)
        ElseIf Mid(src, idx, 1) = ")" Then
            idx = idx + 1
            block = block - 1
            If block < 0 Then cerr "错误：语句块内左括号缺失" & vbCrLf & "位于行：" & tmp
            If UBound(par) > 0 Then ReDim Preserve par(UBound(par) - 1)
            recursion = par
            Exit Function
        Else
            st = ""
            If Mid(src, idx, 1) = Chr(34) Then
                idx = idx + 1: st = st & Chr(34)
                Do Until Mid(src, idx, 1) = Chr(34)
                    st = st & Mid(src, idx, 1)
                    idx = idx + 1
                Loop
                idx = idx + 1
            Else
                Do While Mid(src, idx, 1) <> " " And Mid(src, idx, 1) <> "  " And Mid(src, idx, 1) <> Chr(10) And Mid(src, idx, 1) <> Chr(13) And Mid(src, idx, 1) <> "(" And Mid(src, idx, 1) <> ")" And idx <= codelen
                    st = st & Mid(src, idx, 1)
                    idx = idx + 1
                Loop
            End If
            par(UBound(par)) = st
            ReDim Preserve par(UBound(par) + 1)
        End If
    Loop
    If block > 0 Then cerr "错误：程序无结尾" & vbCrLf & "位于行：" & crlf
    If UBound(par) > 0 Then ReDim Preserve par(UBound(par) - 1)
    recursion = par
End Function
Sub Main() '我是将工程的启动设置为Sub Main()
    f_path = Replace(Command, Chr(34), "")
    nHandle = FreeFile
    Open f_path For Input As #nHandle
    Do Until EOF(nHandle)
        Line Input #nHandle, newline
        sourcecode = sourcecode & newline & vbCrLf
    Loop
    Close nHandle
    codelen = 0
    curlv = 0
    block = 0
    rt = ""
    ifabort = False
    ewhile = False
    idx = 1
    crlf = 1
    ReDim vars(0)
    ReDim funcs(0)
    Call parse
End Sub
'看到这儿啦？
'用心看完啦？
'OK!恭喜你学会了如何写一个简单的解释器！
'请尝试自己动手写一个吧！
