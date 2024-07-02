Attribute VB_Name = "RedBlackTreeTest"
Option Explicit

Sub CompareValues(ByRef expected() As Long, ByRef cur As Long, ByRef buf() As RedBlackTreeTemplate.NodeTypeTemplate, ByVal root As Long)
    If root = -1 Then Exit Sub
    If buf(root).rbChild(0) <> -1 Then
        CompareValues expected, cur, buf, buf(root).rbChild(0)
    End If
    If buf(root).valueTemplate <> expected(cur) Then Err.Raise 3001 Else cur = cur + 1
    If buf(root).rbChild(1) <> -1 Then
        CompareValues expected, cur, buf, buf(root).rbChild(1)
    End If
End Sub

Sub CompareAscending(ByRef cur As Long, ByRef buf() As RedBlackTreeTemplate.NodeTypeTemplate, ByVal root As Long)
    If root = -1 Then Exit Sub
    If buf(root).rbChild(0) <> -1 Then
        CompareAscending cur, buf, buf(root).rbChild(0)
    End If
    If buf(root).valueTemplate <> cur Then Err.Raise 3001 Else cur = cur + 1
    If buf(root).rbChild(1) <> -1 Then
        CompareAscending cur, buf, buf(root).rbChild(1)
    End If
End Sub

Function CheckBlackHeight(ByRef buf() As RedBlackTreeTemplate.NodeTypeTemplate, ByVal root As Long) As Long
    Dim x As Long, y As Long
    
    If root = -1 Then
        CheckBlackHeight = 0
        Exit Function
    End If
    
    x = CheckBlackHeight(buf, buf(root).rbChild(0))
    y = CheckBlackHeight(buf, buf(root).rbChild(1))
    If x <> y Then Err.Raise 3001
    CheckBlackHeight = x - buf(root).rbIsBlack
End Function

Sub CheckNoRedViolation( _
    ByRef buf() As RedBlackTreeTemplate.NodeTypeTemplate, ByVal root As Long, _
    Optional ByVal parentIsBlack = True _
)
    Dim xOk As Boolean, yOk As Boolean
    
    If root = -1 Then
        Exit Sub
    End If
    
    If Not buf(root).rbIsBlack And Not parentIsBlack Then
        Err.Raise 3002
        Exit Sub
    End If
    
    CheckNoRedViolation buf, buf(root).rbChild(0), buf(root).rbIsBlack
    CheckNoRedViolation buf, buf(root).rbChild(1), buf(root).rbIsBlack
End Sub

Sub CheckLinkConsistency( _
    ByRef buf() As RedBlackTreeTemplate.NodeTypeTemplate, ByVal root As Long, _
    Optional ByVal parent = -1 _
)
    If root = -1 Then Exit Sub
    
    With buf(root)
        If .rbParent <> parent Then Err.Raise 3003
        If parent <> -1 Then
            If buf(parent).rbChild(0) <> root And buf(parent).rbChild(1) <> root Then Err.Raise 3003
        End If
        If .rbChild(0) <> -1 Then
            If buf(.rbChild(0)).rbParent <> root Then Err.Raise 3003
        End If
        If .rbChild(1) <> -1 Then
            If buf(.rbChild(1)).rbParent <> root Then Err.Raise 3003
        End If
        CheckLinkConsistency buf, .rbChild(0), root
        CheckLinkConsistency buf, .rbChild(1), root
    End With
End Sub

Sub RunTests(ByRef msg As String)
    Dim i As Long, n As Long
    
    TestAscending msg
    TestDescending msg
    TestRandomSequence1 msg
    TestRandomSequence2 msg
    TestRandomSequence3 msg
    
    n = 500
    msg = msg & vbCrLf & "Performed " + CStr(n) + " random tests."

    For i = 1 To n
        PerformRandomTest 1, 10000
    Next
End Sub
    
Sub TestAscending(ByRef msg As String)
    Dim nodes(0 To 299) As RedBlackTreeTemplate.NodeTypeTemplate
    Dim expected(0 To 299) As Long
    Dim i As Long
    
    msg = msg & vbCrLf & "Tested ascending sequence."
    
    For i = 0 To 299
        nodes(i).valueTemplate = i: expected(i) = i
    Next
    
    RunTest expected, nodes
End Sub

Sub TestDescending(ByRef msg As String)
    Dim nodes(0 To 299) As RedBlackTreeTemplate.NodeTypeTemplate
    Dim expected(0 To 299) As Long
    Dim i As Long
    
    msg = msg & vbCrLf & "Tested descending sequence."
    For i = 0 To 299
        nodes(i).valueTemplate = 299 - i: expected(i) = i
    Next
    
    RunTest expected, nodes
End Sub

Sub MakeArray(ByRef ary() As Long, ParamArray v())
    Dim i As Long, lb As Long, ub As Long
    lb = LBound(v): ub = UBound(v)
    ReDim ary(0 To ub - lb) As Long
    For i = lb To ub
        ary(i - lb) = v(i)
    Next
End Sub

Sub TestRandomSequence1(ByRef msg As String)
    Dim nodes() As RedBlackTreeTemplate.NodeTypeTemplate
    Dim values() As Long, expected() As Long, ub As Long, i As Long
    MakeArray values, 476, 303, 344, 586, 701, 918, 902, 132, 952, 948, 915, 740, 514, 88, 44, 906, 884, 211, 108, 594, 659, 319, 465, 10, 870, 390, 278, 695, 683, 156
    MakeArray expected, 10, 44, 88, 108, 132, 156, 211, 278, 303, 319, 344, 390, 465, 476, 514, 586, 594, 659, 683, 695, 701, 740, 870, 884, 902, 906, 915, 918, 948, 952
    msg = msg & vbCrLf & "Tested random sequence 1."
    
    ub = UBound(values)
    ReDim nodes(0 To ub) As RedBlackTreeTemplate.NodeTypeTemplate
    For i = 0 To ub: nodes(i).valueTemplate = values(i): Next
    
    RunTest expected, nodes
End Sub

Sub TestRandomSequence2(ByRef msg As String)
    Dim nodes() As RedBlackTreeTemplate.NodeTypeTemplate
    Dim values() As Long, expected() As Long, ub As Long, i As Long
    MakeArray values, 230, 540, 166, 654, 454, 936, 133, 4, 146, 803, 963, 127, 247, 941, 574, 331, 56, 184, 499, 562, 408, 175, 326, 49, 635, 547, 117, 751, 180, 911
    MakeArray expected, 4, 49, 56, 117, 127, 133, 146, 166, 175, 180, 184, 230, 247, 326, 331, 408, 454, 499, 540, 547, 562, 574, 635, 654, 751, 803, 911, 936, 941, 963

    msg = msg & vbCrLf & "Tested random sequence 2."
    
    ub = UBound(values)
    ReDim nodes(0 To ub) As RedBlackTreeTemplate.NodeTypeTemplate
    For i = 0 To ub: nodes(i).valueTemplate = values(i): Next
    
    RunTest expected, nodes
End Sub

Sub TestRandomSequence3(ByRef msg As String)
    Dim nodes() As RedBlackTreeTemplate.NodeTypeTemplate
    Dim values() As Long, expected() As Long, ub As Long, i As Long
    MakeArray values, 929, 605, 482, 116, 249, 264, 114, 273, 458, 266, 864, 8, 598, 369, 438, 130, 576, 357, 128, 798, 999, 40, 36, 817, 996, 925, 739, 443, 804, 289, 429, 83, 911, 861, 550, 849, 381, 855, 966, 396, 360, 104, 38, 657, 755, 339, 40, 858, 862, 452
    MakeArray expected, 8, 36, 38, 40, 83, 104, 114, 116, 128, 130, 249, 264, 266, 273, 289, 339, 357, 360, 369, 381, 396, 429, 438, 443, 452, 458, 482, 550, 576, 598, 605, 657, 739, 755, 798, 804, 817, 849, 855, 858, 861, 862, 864, 911, 925, 929, 966, 996, 999
    msg = msg & vbCrLf & "Tested random sequence 3."
    
    ub = UBound(values)
    ReDim nodes(0 To ub) As RedBlackTreeTemplate.NodeTypeTemplate
    For i = 0 To ub: nodes(i).valueTemplate = values(i): Next
    
    RunTest expected, nodes
End Sub

Sub RunTest(ByRef expected() As Long, ByRef nodes() As RedBlackTreeTemplate.NodeTypeTemplate)
    Dim root As Long
    Dim rightDir As Boolean
    Dim foundIdx As Long
    Dim parent As Long
    Dim cur As Long
    Dim i As Long
    
    root = -1 ' tree initially empty
    For i = 0 To UBound(nodes)
        foundIdx = RedBlackTreeTemplate.RedBlackFindPosition(parent, rightDir, nodes, root, nodes(i).valueTemplate)
        If foundIdx = -1 Then RedBlackTreeTemplate.RedBlackInsert nodes, root, i, parent, rightDir
    Next
    
    cur = 0
    CompareValues expected, cur, nodes, root
    If cur <> UBound(expected) + 1 Then Err.Raise 3001

    CheckBlackHeight nodes, root
    CheckNoRedViolation nodes, root
    CheckLinkConsistency nodes, root
End Sub

Sub PerformRandomTest(ByVal minArySize As Long, ByVal maxArySize As Long)
    Dim nodes() As RedBlackTreeTemplate.NodeTypeTemplate
    Dim n As Long
    Dim i As Long, j As Long
    Dim tmp As RedBlackTreeTemplate.NodeTypeTemplate
    Dim cur As Long
    
    Dim root As Long
    Dim foundIdx As Long
    Dim parent As Long
    Dim rightDir As Boolean
    
    n = Int(((maxArySize - minArySize) * Rnd) + minArySize)

    Debug.Print "n = ", n
    
    ReDim nodes(0 To n - 1) As RedBlackTreeTemplate.NodeTypeTemplate
    For i = 0 To n - 1
        nodes(i).valueTemplate = i
    Next
    ' Shuffle
    For i = 0 To n - 1
        j = Int((n - 1 - i) * Rnd + i)
        tmp = nodes(i): nodes(i) = nodes(j): nodes(j) = tmp
    Next
    
    root = -1 ' tree initially empty
    For i = 0 To UBound(nodes)
        foundIdx = RedBlackTreeTemplate.RedBlackFindPosition(parent, rightDir, nodes, root, nodes(i).valueTemplate)
        If foundIdx = -1 Then RedBlackTreeTemplate.RedBlackInsert nodes, root, i, parent, rightDir
    Next
    
    cur = 0
    CompareAscending cur, nodes, root
    If cur <> n Then Err.Raise 3010
    
    CheckBlackHeight nodes, root
    CheckNoRedViolation nodes, root
    CheckLinkConsistency nodes, root
End Sub

