Attribute VB_Name = "RedBlackTreeTemplate"
Type NodeTypeTemplate
    rbParent As Long
    rbChild(0 To 1) As Long
    rbIsBlack As Boolean
    valueTemplate As Long
End Type

Function RedBlackComparatorTemplate(ByRef v1 As Long, ByRef v2 As Long) As Long
    If v1 < v2 Then
        RedBlackComparatorTemplate = -1
    ElseIf v1 = v2 Then
        RedBlackComparatorTemplate = 0
    Else
        RedBlackComparatorTemplate = 1
    End If
End Function

' assumption: root <> -1
Function RedBlackFind(ByRef n As Long, ByRef rightDir As Boolean, ByRef buf() As NodeTypeTemplate, ByVal root As Long, ByVal v As Long) As Boolean
    Dim cmp As Long, cur As Long, parent As Long, d As Boolean
    cur = root: parent = -1
    Do
        cmp = RedBlackComparatorTemplate(v, buf(cur).valueTemplate)
        If cmp < 0 Then
            parent = cur: d = False
            cur = buf(cur).rbChild(0)
        ElseIf cmp = 0 Then
            n = cur
            RedBlackFind = True
            Exit Function
        Else
            parent = cur: d = True
            cur = buf(cur).rbChild(1)
        End If
    Loop Until cur = -1
    n = parent: rightDir = d: RedBlackFind = False
End Function

Private Sub RedBlackRotate(ByRef buf() As NodeTypeTemplate, ByRef root As Long, ByVal p As Long, ByVal rightDir As Boolean)
    Dim g As Long, s As Long, c As Long

    With buf(p)
        g = .rbParent
        s = .rbChild(1 + rightDir)
    End With

    With buf(s)
        c = .rbChild(-rightDir)
        .rbChild(-rightDir) = p
        .rbParent = g
    End With

    With buf(p)
        .rbChild(1 + rightDir) = c
        .rbParent = s
    End With

    If c <> -1 Then buf(c).rbParent = p

    If g <> -1 Then
        With buf(g): .rbChild(-(p = .rbChild(1))) = s: End With
    Else
        root = s
    End If
End Sub

Sub RedBlackInsert(ByRef buf() As NodeTypeTemplate, ByRef root As Long, ByVal n As Long, ByVal p As Long, ByVal rightDir As Boolean)
    With buf(n)
        .rbIsBlack = False
        .rbChild(0) = -1
        .rbChild(1) = -1
        .rbParent = p
    End With
    If p = -1 Then
        root = n
        Exit Sub
    End If
    buf(p).rbChild(-rightDir) = n
    Do
        If buf(p).rbIsBlack Then Exit Sub
        ' From now on P is red.
        g = buf(p).rbParent
        If g = -1 Then ' P red and root
            buf(p).rbIsBlack = True
            Exit Sub
        End If
        ' P is red and not root (G exists)
        ' rightDir is True if P is the right-hand child of G and False otherwise
        rightDir = (buf(buf(p).rbParent).rbChild(1) = p)
        U = buf(g).rbChild(1 + rightDir)
        If U = -1 Then GoTo CaseI56
        If buf(U).rbIsBlack Then GoTo CaseI56

        ' P and U red, G exists
        buf(p).rbIsBlack = True
        buf(U).rbIsBlack = True
        buf(g).rbIsBlack = False
        n = g
        p = buf(n).rbParent
    Loop Until p = -1
    Exit Sub

CaseI56: ' P red and U black (or does not exist), G exists
    If n = buf(p).rbChild(1 + rightDir) Then
        ' CaseI5 (P red and U black and N inner grandchild of G)
        RedBlackRotate buf, root, p, rightDir ' P is never the root, so param root is meaningless here
        n = p ' new current node
        p = buf(g).rbChild(-rightDir)  ' new parent of N
        ' fall through to CaseI6
    End If

    ' CaseI6 (P red and U black and N outer grandchild of G)
    RedBlackRotate buf, root, g, 1 + rightDir ' G may be the root
    buf(p).rbIsBlack = True
    buf(g).rbIsBlack = False
End Sub
