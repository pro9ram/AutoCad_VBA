Attribute VB_Name = "MyMath"
Function getLinearEQuationA(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    Dim dx As Double
    Dim ret As Double
    dx = x2 - x1
    If dx = 0 Then
        dx = 1
    End If
    
    ret = (y2 - y1) / dx
    
    getLinearEQuationA = ret

End Function

Function getLinearEQuationC(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double

    getLinearEQuationC = -getLinearEQuationA(x1, y1, x2, y2) * x1 + y1

End Function


Function getShortestDistance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x As Double, ByVal y As Double) As Double

    Dim a, b, c, d As Double

    a = getLinearEQuationA(x1, y1, x2, y2)
    b = -1
    c = getLinearEQuationC(x1, y1, x2, y2)
    
    d = Abs((a * x + b * y + c)) / Sqr(a * a + b * b)
     
    getShortestDistance = d


End Function


Function getDistance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double

    getDistance = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))

End Function



Function getAngleEx(ByVal cx As Double, ByVal cy As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double

    getAngleEx = ArcTan2((x2 - cx), (y2 - cy)) - ArcTan2((x1 - cx), (y1 - cy))
    
End Function


Function getAngle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double

    Dim dx, dy, rad As Double

    dx = x2 - x1
    dy = y2 - y1

    
    rad = ArcTan2(dx, dy)
    
    getAngle = rad
    
End Function

Function crossXY(ByVal cx As Double, ByVal cy As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double()

    Dim d, r As Double
    Dim th, rad, radex As Double
    Dim nx, ny As Double
    Dim dx, dy As Double
    Dim ret(0 To 2) As Double
    
 
    d = getShortestDistance(x1, y1, x2, y2, cx, cy)
    
    rad = getAngle(x1, y1, x2, y2)
    
    rad = Abs(rad)
    
    If rad > 0 And rad > PI_2 Then
        rad = PI - rad
    End If
    
    If rad < 0 And rad < -PI_2 Then
        rad = PI + rad
    End If
    
    
    
    radex = getAngleEx(x1, y1, cx, cy, x2, y2)
    
    radex = Abs(radex)
    
    r = getDistance(x1, y1, cx, cy)
    
    
    'r = d / Tan(radex)
    dx = r * Cos(rad)
    dy = r * Sin(rad)

    If x2 > x1 Then
        nx = x1 + dx
    Else
        nx = x1 - dx
    End If
    
    
    If y2 > y1 Then
        ny = y1 + dy
    Else
        ny = y1 - dy
    End If
    

    'addDonut2 cx, cy
    'addDonut2 x1, y1
    'addDonut2 x2, y2
    'addDonut2 nx, ny
    
    
    ret(0) = nx
    ret(1) = ny
    
    crossXY = ret
    

End Function

 

Sub testnewxy()
    
    Dim p1() As Double
    Dim ent2 As AcadObject
    Dim ent2p() As Double
    Dim sel As AcadSelectionSet
    Dim ddd() As Double
    Dim ret() As Double
    
    With ThisDrawing
        On Error Resume Next
        .SelectionSets("CurrentSelection").Delete
        Set sel = .SelectionSets.add("CurrentSelection")
        
        p1 = ThisDrawing.Utility.getPoint(, "Enter a point: ")
        ThisDrawing.Utility.getEntity ent2, ent2p, "Enter a Entity: "
        
        ddd = ent2.Coordinates
        
        ret = crossXY(p1(0), p1(1), ddd(0), ddd(1), ddd(2), ddd(3))
        
        
        ThisDrawing.ModelSpace.addLine p1, ret
        
        
    End With
    
    'NewXY 0, 2, 4, 4, 1, 5
    
    
End Sub

Sub testdistance()

    Dim d, r As Double
    Dim th, rad, radex As Double
    Dim nx, ny As Double
 
    d = getShortestDistance(0, 2, 4, 4, 1, 5)
    
    rad = getAngle(0, 2, 4, 4)
    radex = getAngleEx(0, 2, 4, 4, 1, 5)
    
    
    r = d / Tan(radex)

    nx = r * Cos(rad)
    ny = r * Sin(rad)

    Debug.Print ""
    
End Sub





