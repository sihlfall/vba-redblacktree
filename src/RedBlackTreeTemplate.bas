Attribute VB_Name = "RedBlackTreeTemplate"
Option Explicit

Type NodeTypeTemplate
    rbParent As Long
    rbChild(0 To 1) As Long ' 0 = left, 1 = right
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

' returns
' * index of node if node was found; in that case out parameters will remain unchanged
' * -1 otherwise, then the out parameters will indicate where the new node would have to be inserted
Function RedBlackFindPosition(ByRef outParent As Long, ByRef outAsRightHandChild As Boolean, ByRef buf() As NodeTypeTemplate, ByVal root As Long, ByVal v As Long) As Long
    Dim cmp As Long, cur As Long, p As Long, rhc As Boolean
    
    cur = root: p = -1: rhc = False
    Do Until cur = -1
        cmp = RedBlackComparatorTemplate(v, buf(cur).valueTemplate)
        If cmp < 0 Then
            p = cur: rhc = False
            cur = buf(cur).rbChild(0)
        ElseIf cmp = 0 Then
            RedBlackFindPosition = cur
            Exit Function
        Else
            p = cur: rhc = True
            cur = buf(cur).rbChild(1)
        End If
    Loop
    outParent = p: outAsRightHandChild = rhc: RedBlackFindPosition = -1
End Function

' Algorithm adapted from https://en.wikipedia.org/w/index.php?title=Red%E2%80%93black_tree&oldid=1150140777
Sub RedBlackInsert(ByRef buf() As NodeTypeTemplate, ByRef outRoot As Long, ByVal newNode As Long, ByVal parent As Long, ByVal asRightHandChild As Boolean)
    Dim g As Long, u As Long, p As Long, n As Long, pIsRhc As Boolean
    Dim gg As Long, b As Long, c As Long, x As Long, y As Long, z As Long, nIsRhc As Boolean
    
    With buf(newNode)
        .rbIsBlack = False
        .rbChild(0) = -1
        .rbChild(1) = -1
        .rbParent = parent
    End With
    If parent = -1 Then
        outRoot = newNode
        Exit Sub
    End If
    
    buf(parent).rbChild(-asRightHandChild) = newNode
    
    n = newNode: p = parent
    Do
        If buf(p).rbIsBlack Then Exit Sub
        ' p red
        g = buf(p).rbParent
        If g = -1 Then ' p red and root
            buf(p).rbIsBlack = True
            Exit Sub
        End If
        ' p red and not root (g exists)
        ' u is supposed to refer to the brother of p
        pIsRhc = buf(g).rbChild(1) = p
        u = buf(g).rbChild(1 + pIsRhc)
        If u = -1 Then GoTo ExitWithRotation
        If buf(u).rbIsBlack Then GoTo ExitWithRotation

        ' p and u red, g exists
        buf(p).rbIsBlack = True
        buf(u).rbIsBlack = True
        buf(g).rbIsBlack = False
        n = g
        p = buf(n).rbParent
    Loop Until p = -1
    Exit Sub

ExitWithRotation: ' p red and u black (or does not exist), g exists
    ' For an explanation of the following, see
    '   https://en.wikibooks.org/w/index.php?title=F_Sharp_Programming/Advanced_Data_Structures&oldid=4052491 ,
    '   Section 3.1 ("Red Black Trees"), second diagram (following the sentence "The center tree is the balanced version.").
    nIsRhc = buf(p).rbChild(1) = n
    If pIsRhc = nIsRhc Then ' outer child
        y = p
        If pIsRhc Then
            b = buf(p).rbChild(0): c = buf(n).rbChild(0): x = g: z = n
        Else
            b = buf(n).rbChild(1): c = buf(p).rbChild(1): x = n: z = g
        End If
    Else ' inner child
        y = n: With buf(n): b = .rbChild(0): c = .rbChild(1): End With
        If pIsRhc Then
            x = g: z = p
        Else
            x = p: z = g
        End If
    End If
    
    gg = buf(g).rbParent
    
    With buf(x): .rbIsBlack = False: .rbParent = y: .rbChild(1) = b: End With
    With buf(y): .rbIsBlack = True: .rbParent = gg: .rbChild(0) = x: .rbChild(1) = z: End With
    With buf(z): .rbIsBlack = False: .rbParent = y: .rbChild(0) = c: End With
    
    If b <> -1 Then buf(b).rbParent = x
    If c <> -1 Then buf(c).rbParent = z
    
    If gg = -1 Then
        outRoot = y
    Else
        With buf(gg): .rbChild(-(.rbChild(1) = g)) = y: End With
    End If
End Sub
