Attribute VB_Name = "Module1"
'ҵ��д��һ������lisp�﷨�Ľű����ԵĽ�����
'ҵ��д��һ������Lisp�﷨�Ľű����ԵĽ�����
'���ƣ�ToyLisp������Ǹ���ߣ� �ļ���׺��.LSP���Ҳ���LSP��
'�������еĶ����������������������ѭ������ѧ���㡢����������ַ����������ļ���д�����б������������򡢴�����ʾ�ȶ��У�^V^��
'�﷨�ܼ򵥣�����([������/������] [�����б�])
'��(out 1 2 3 (+ 4 5))
'ÿһ������ֻ��һ�����main������(main (out "Hello World!"))
'��������ظ������������ȶ���ĺ���Ϊ��׼
'��д�����Լ���ToyLisp����󣬿�����ѡ��򿪷�ʽ��ѡ��ToyLisp.exe����
'�����﷨�����뿴ע��
'�����裨ba��©��ge���������
'Ҳ��л����Ķ�����˵������Ը�⿴����һ��������𣨱�����
'��������˿����Ļ�������Ҳ���£��Ҿ����д�������ˣ�
'����������Ʒ��TG�ļ�����һ����һ����֪������һ�����ļ���
'��ģ�ȫ�ǽ�����(ToyLisp������������GrapLispȥ����ͼ���ܼ��޸���һЩ����Bug�İ汾)
'������û��ToyLisp��ô����ȫ�桢�����ȶ�
'��������Ȥ������ȥ��һ��
'(��-��-)zzz
'Yauhak Chen - QQ3953814837
Dim vars()
'����
Dim funcs()
'���庯��
Dim sourcecode
'��ȡ��Դ����
Dim codelen, src, idx, curlv, ifabort, rt, ewhile
'Դ���볤�ȣ�Դ���룬��ǰ����λ�ã���ǰ������ȼ����Ƿ��˳�������ѭ������������ֵ���Ƿ����ѭ��
Dim block, crlf 'ƥ�����ţ���ǰ��������
'�޸��������
Sub opar(tag, mat, ide, val)
    If ide = UBound(mat) Then
        tag(rec(mat(ide))) = val
    Else
        opar tag(rec(mat(ide))), mat, ide + 1, val
    End If
End Sub
'������Ϣ���
'������ܲ�����
Sub cerr(exp)
    MsgBox exp, vbCritical, "ToyLisp"
    End
End Sub
Function rec(ary)
'ִ�б��������б�Ĵ���
On Error Resume Next
'����д�����ж��ˣ�̫����
tmp = crlf
If IsArray(ary) Then
If UBound(ary) = -1 Then Exit Function
If UBound(ary) > 0 Then
If Not IsArray(ary(0)) Then
    operate_name = LCase(ary(0))
    If Not operate_name = "def" And Not operate_name = "do" And Not operate_name = "return" And Not operate_name = "m" _
    And Not operate_name = "array" And Not operate_name = "out" And Not operate_name = "in" And Not operate_name = "#" _
    And Not operate_name = "size" And Not operate_name = "m" Then fst_n = rec(ary(1))
    'fst_n ����������ҵĳ����Ǽ�¼��ѧ�������ĵ�һ��ֵ������+-*/��������������ѧ���㣬��ȻҪ����ö��
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
    '������ַ�������
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
    'ȡ����˷�
    Case "sqrt": '����
        fst_n = Sqr(fst_n)
    Case "sin":
        fst_n = Sin(fst_n)
    Case "cos":
        fst_n = Cos(fst_n)
    Case "tan":
        fst_n = Tan(fst_n)
    Case "atan":
        fst_n = Atn(fst_n)
    '���Ǻ���
    Case "abs":
        fst_n = Abs(fst_n)
    Case "int":
        fst_n = Int(fst_n)
    Case "fix":
        fst_n = Fix(fst_n)
    'Int �� Fix ����ɾ�� number ��С�����ݶ�����ʣ�µ�������
    'Int �� Fix �Ĳ�֮ͬ�����ڣ���� number Ϊ�������� Int ����С�ڻ���� number �ĵ�һ����������
    '�� Fix ��᷵�ش��ڻ���� number �ĵ�һ�������������磬Int �� -8.4 ת���� -9���� Fix �� -8.4 ת���� -8��
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
    '�����ж�
    Case "out":
        oput = ""
        For x = 1 To UBound(ary)
            oput = oput & rec(ary(x))
        Next
        MsgBox oput, , "ToyLisp"
        Exit Function
    '��������޲�����
    Case "in":
        fst_n = InputBox(rec(ary(1)), "ToyLisp")
    '����
    Case "asc":
        fst_n = Asc(rec(ary(1)))
    Case "chr":
        fst_n = Chr(rec(ary(1)))
    'asc�����ַ����໥ת��
    Case "len":
        fst_n = Len(fst_n)
    '�ַ�������
    Case "size":
        rec = UBound(rec(ary(1)))
        Exit Function
    '���鳤��
    Case "split":
        rec = Split(fst_n, rec(ary(2)))
        Exit Function
    '�ַ����ָ�
    Case "substr":
        rec = Mid(fst_n, rec(ary(2)), rec(ary(3)))
        Exit Function
    '(substr "string" 1 1)��"string"�ַ�����һ��λ�ý�ȡһ������Ϊһ���ַ���
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
    '����һ������
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
    '����һ���б����飩
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
    '������������
    '(def x (list 1 2 3 (list 4 5)))
    '(out (m x 3 0)),�������������Ժ���±�
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
    '��������
    '(def x (list 1 2 3 (list 4 5)))
    '(array x (3 0) "H")���м�������ı��±�
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
    '�ļ����ݶ�д
    Case "alloc":
        Dim memspace()
        ReDim memspace(fst_n)
        rec = memspace
        Exit Function
    '����һ��ָ����С�����飨Ԫ�ض�Ϊ�գ�
    '(def x (alloc 100))
    Case "do":
        For x = 1 To UBound(ary)
            todo = ary(x)
            Call rec(todo)
        Next
        Exit Function
    '�����Ҫ�ɲ�Ҫ
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
    'if([���ʽ][����1][����2����ѡ��])
    'ע��������ʽ����һ��ʱҪ��do������������
    'whileҲһ��
    '��(if(= 1 1)((out 1)(out"y"))(out 0))
    Case "while":
        Do While rec(ary(1)) = 1 And ifabort = False And ewhile = False
            rec (ary(2))
        Loop
        If ewhile = True Then ewhile = False
        Exit Function
    '(while (= 1 1)((out 1)(out 2)))
    Case "break":
        '����ѭ��
        ewhile = True
        Exit Function
    Case "return":
        '�����Զ��庯��ֵ
        rt = rec(ary(1))
        ifabort = True
        Exit Function
    Case "#": 'ע�ͣ�������ע�⣡ע��Ҳ����Ϊ������Ҫ�ǵÿո��Ҷ���if��while�����ҪС�ģ�������
        Exit Function
    Case Else: '�Զ��庯��
        If UBound(funcs) = 0 And IsEmpty(funcs(0)) Then cerr "���棺����δ����" & vbCrLf & "������" & operate_name: Exit Function
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
                End If '�Զ��庯�����Σ�curlv�ǵ�ǰ������ȼ�����ֹ��������Խ��
                '������ע�⣺�����޲����ĺ���Ӧ�ں�������������дһ��������"()"Ҳ�У�����"(break())"��,����ᱻ��������Ϊ�Ǳ���
                For exec = 3 To UBound(fn)
                    this = rec(fn(exec))
                    If ifabort = True Then ifabort = False: rec = rt: rt = "": Exit For
                Next
                ReDim Preserve vars(bound) '��������������ı�������
                curlv = curlv - 1
                Exit Function
            End If
         Next
         cerr "���棺����δ����" & vbCrLf & "������" & operate_name
    End Select
    End If
    If IsArray(ary(0)) Then
        For x = 0 To UBound(ary)
            todo = ary(x)
            Call rec(todo)
        Next
        Exit Function
    End If '������������һϵ�б��ʽ��ִ��
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
        cerr "���棺����δ����" & vbCrLf & "������" & ary
    Else
        rec = Replace(ary, Chr(34), "")
        Exit Function
    End If
End If '�������ַ�������ֵ�ķ���
End Function
Sub parse()
'��Դ����������б�
    src = sourcecode
    tm = Len(src)
    If tm = 0 Then Exit Sub
    Do While Mid(src, tm, 1) = Chr(13) Or Mid(src, tm, 1) = Chr(10) Or Mid(src, tm, 1) = " " Or Mid(src, tm, 1) = " "
        tm = tm - 1
    Loop
    '���ƺ���һ��bug�������Դ�����β��crlf��ո�ȥ������������������
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
    Next '���б���
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
    Next '������
End Sub
'����ǽ���������
'��Դ����������б���ʵ�������飩
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
                If (par(UBound(par))(0) = "fn" Or par(UBound(par))(0) = "main") And block > 0 Then cerr "�����������ж��庯��" & vbCrLf & "λ���У�" & tmp
            End If
            ReDim Preserve par(UBound(par) + 1)
        ElseIf Mid(src, idx, 1) = ")" Then
            idx = idx + 1
            block = block - 1
            If block < 0 Then cerr "����������������ȱʧ" & vbCrLf & "λ���У�" & tmp
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
    If block > 0 Then cerr "���󣺳����޽�β" & vbCrLf & "λ���У�" & crlf
    If UBound(par) > 0 Then ReDim Preserve par(UBound(par) - 1)
    recursion = par
End Function
Sub Main() '���ǽ����̵���������ΪSub Main()
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
'�����������
'���Ŀ�������
'OK!��ϲ��ѧ�������дһ���򵥵Ľ�������
'�볢���Լ�����дһ���ɣ�
