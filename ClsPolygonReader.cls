VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPolygonReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ent As AcadEntity
Private index As Integer    '현재위치
Private coord() As Double
Private size As Integer '포인트 크기(/3 - x,y,z)
Private sidx As Integer '시작지점
Private reverse As Boolean
Const C_ST = 2

Public Function init(e As AcadEntity)
    Set ent = e
    
    coord = ent.Coordinates
    size = UBound(coord)
    
    reverse = False
End Function

Public Function setIndex(i As Integer)

    index = i

End Function


Public Function setIndexByPoint(p As clsPoint)
    
    sidx = getIndexByPoint(p)
    

End Function


Public Function getIndexByPoint(p As clsPoint) As Integer

    sidx = -1


    Debug.Print "BSE: " & p.x & p.y & "==========="
    

    For i = 0 To size Step C_ST
    
        x = coord(i)
        y = coord(i + 1)
        
        'Debug.Print "idx: " & i & " > "  & x & y & "===="
        
        If Int(p.x * 100) = Int(x * 100) And Int(p.y * 100) = Int(y * 100) Then
            sidx = i
            getIndexByPoint = sidx
            Debug.Print "idx: " & i & " > " & x & y & "===="
            Exit Function
        End If
    
    Next i
    
    
    Debug.Print "result: getIndexByPoint: " & sidx
    
    If sidx = -1 Then
        Debug.Print "getIndexByPoint: " & sidx
    End If
    

End Function





Public Function checkDirection(p As clsPoint)

    Dim cp As New clsPoint
    Dim np As New clsPoint
    Dim angle As Double
    
    
    setIndex sidx
    Set cp = getPoint(sidx)
    Set np = nextPoint
    


    angle = getAngleEx(cp.x, cp.y, np.x, np.y, p.x, p.y)
    
    
    If Abs(angle) > PI_2 Then
        reverse = True
    End If
    
    
    
    



End Function

Private Function nextPoint() As clsPoint
    Dim p As New clsPoint
    
    
    If reverse = False Then
        index = index + C_ST
    
        If index > size Then
            index = 0
        End If
    Else
        index = index - C_ST
    
        If index < 0 Then
        index = size - C_ST + 1
        End If
    
    End If
    
    
    p.x = coord(index)
    p.y = coord(index + 1)
    
    
    Set nextPoint = p
End Function

Private Function prevPoint() As clsPoint
    Dim p As New clsPoint
    
    If reverse = False Then
        index = index - C_ST
        
        If index < 0 Then
            index = size - C_ST + 1
        End If
    Else
        index = index + C_ST
    
        If index > size Then
            index = 0
        End If
    End If
    
    
    p.x = coord(index)
    p.y = coord(index + 1)
    
    Set prevPoint = p
End Function


Public Function getPoint(idx As Integer) As clsPoint
    Dim p As New clsPoint
    
    p.x = coord(idx)
    p.y = coord(idx + 1)
    
    Set getPoint = p
End Function



Public Function splitInnerLine(startp As clsPoint, endp As clsPoint) As Double()

    Dim p As New clsPoint
    Dim np As New clsPoint
    Dim count As Integer
    Dim ret() As Double
    Dim idx As Integer

    setIndexByPoint startp
    setIndex (sidx)
    
    Set p = getPoint(sidx)
    Set np = nextPoint()
    
    count = 1
    
    Do While (np.isEquals(endp) = False And size > count)
        count = count + 1
        Set np = nextPoint()
        
    Loop
    
    
    idx = 0
    ReDim ret(0 To count * 2 + 1) As Double
    
    setIndexByPoint startp
    setIndex (sidx)
    
    Set p = getPoint(sidx)
    
    ret(idx) = p.x
    idx = idx + 1
    ret(idx) = p.y
    idx = idx + 1
    
    Set np = nextPoint()
    
    ret(idx) = np.x
    idx = idx + 1
    ret(idx) = np.y
    idx = idx + 1
    
    count = 1
    
    Do While (np.isEquals(endp) = False And size > count)
        count = count + 1
        Set np = nextPoint()
        ret(idx) = np.x
        idx = idx + 1
        ret(idx) = np.y
        idx = idx + 1
    Loop
    
    splitInnerLine = ret
    
End Function


Public Function splitOutLine(startp As clsPoint, endp As clsPoint) As Double()

    Dim p As New clsPoint
    Dim np As New clsPoint
    Dim count As Integer
    Dim ret() As Double
    Dim idx As Integer

    setIndexByPoint startp
    setIndex (sidx)
    
    Set p = getPoint(sidx)
    Set np = nextPoint()
    
    count = 1
    
    Do While (np.isEquals(endp) = False And size > count)
        count = count + 1
        Set np = nextPoint()
        
    Loop
    
    
    idx = 0
    ReDim ret(0 To count * 2 + 1) As Double
    
    setIndexByPoint startp
    setIndex (sidx)
    
    Set p = getPoint(sidx)
    
    ret(idx) = p.x
    idx = idx + 1
    ret(idx) = p.y
    idx = idx + 1
    
    Set np = nextPoint()
    
    ret(idx) = np.x
    idx = idx + 1
    ret(idx) = np.y
    idx = idx + 1
    
    count = 1
    
    Do While (np.isEquals(endp) = False And size > count)
        count = count + 1
        Set np = nextPoint()
        ret(idx) = np.x
        idx = idx + 1
        ret(idx) = np.y
        idx = idx + 1
    Loop
    
    splitInnerLine = ret
    
End Function
