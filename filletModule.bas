Attribute VB_Name = "filletModule"
Private Const PI As Double = 3.14159265358979
Private Const PI_2 As Double = 1.5707963267949

Public Const SHOW_CIRCLE = True

Function getArcBy3P(p1() As Double, p2() As Double, p3() As Double) As AcadArc
    Dim centerPoint(0 To 2) As Double
    Dim retOjb As AcadArc
    
    x1 = p1(0)
    y1 = p1(1)
    
    x2 = p2(0)
    y2 = p2(1)
     
    x3 = p3(0)
    y3 = p3(1)
    
    d1 = (x2 - x1) / (y2 - y1)
    d2 = (x3 - x2) / (y3 - y2)
    
    cx = ((y3 - y1) + (x2 + x3) * d2 - (x1 + x2) * d1) / (2 * (d2 - d1))
    cy = -d1 * (cx - (x1 + x2) / 2) + (y1 + y2) / 2
     
    r = VBA.Sqr((x1 - cx) ^ 2 + (y1 - cy) ^ 2)
    
    centerPoint(0) = cx
    centerPoint(1) = cy
    
    StartAngle = ArcTan2((x1 - cx), (y1 - cy))
    EndAngle = ArcTan2((x3 - cx), (y3 - cy))
    
    If StartAngle >= 0 And EndAngle >= 0 Then
        If StartAngle > EndAngle Then
            tmp = StartAngle
            StartAngle = EndAngle
            EndAngle = tmp
        End If
    ElseIf StartAngle < 0 And EndAngle < 0 Then
        If StartAngle > EndAngle Then
            tmp = StartAngle
            StartAngle = EndAngle
            EndAngle = tmp
        End If
    ElseIf StartAngle > 0 And EndAngle < 0 Then
        If (PI - StartAngle) + (PI + EndAngle) > PI Then
            tmp = StartAngle
            StartAngle = EndAngle
            EndAngle = tmp
        End If
    ElseIf StartAngle < 0 And EndAngle > 0 Then
        If Abs(StartAngle) + EndAngle > PI Then
            tmp = StartAngle
            StartAngle = EndAngle
            EndAngle = tmp
        End If
            
    End If
    
    
    Set retOjb = ThisDrawing.ModelSpace.AddArc(centerPoint, r, StartAngle, EndAngle)
     
    
    Set getArcBy3P = retOjb

End Function



Function arc2line(myArc As AcadArc) As AcadEntity

    Dim retObj As AcadEntity
    Dim delta As Double
    
    delta = myArc.EndAngle - myArc.StartAngle
    If delta < 0 Then delta = delta + (2 * PI)
    Dim numOfSegments As Integer
    Dim points() As Double
    'adjust below for reality
    numOfSegments = CInt(myArc.ArcLength) * 10 ' length of segment = 1, last segment = remainder
    ReDim points(0 To 2 * numOfSegments + 1)
    ang = 1 / (myArc.Radius * 10)
    points(0) = myArc.startPoint(0)
    points(1) = myArc.startPoint(1)
    adir = ang
    For X = 2 To UBound(points) - 2 Step 2
        mypolarpoint = ThisDrawing.Utility.PolarPoint(myArc.Center, myArc.StartAngle + adir, myArc.Radius)
        adir = adir + ang
        points(X) = mypolarpoint(0)
        points(X + 1) = mypolarpoint(1)
    Next X
    points(UBound(points) - 1) = myArc.endPoint(0)
    points(UBound(points) - 0) = myArc.endPoint(1)
    Set retObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
    
    Set arc2line = retObj

End Function


Function findVertex(ent As AcadEntity, ent2 As AcadEntity)

    Dim ddd() As Double
    Dim filletpoint(0 To 7) As Double   '8
        
    ddd = ent2.Coordinates
    
    count = UBound(ddd)
    
    vx = ddd(count - 1)
    vy = ddd(count)
    
    
    pntIntersec = ent.IntersectWith(ent2, acExtendNone)
    
    newp = addVertex(ent, vx, vy)
    
    If newp(0) <> 0 And newp(1) <> 0 Then
    
        filletpoint(index) = newp(0)
        index = index + 1
        filletpoint(index) = newp(1)
        index = index + 1
         
    End If
    
    
    vx = ddd(0)
    vy = ddd(1)
    
    newp = addVertex(ent, vx, vy)
    
    If newp(0) <> 0 And newp(1) <> 0 Then
    
        filletpoint(index) = newp(0)
        index = index + 1
        filletpoint(index) = newp(1)
        index = index + 1
         
    End If

End Function

Function addVertex(ent As AcadEntity, ByVal vx As Double, ByVal vy As Double) As Double()

    Dim newv(0 To 1) As Double
    Dim ddd() As Double
    Dim count As Integer
    Dim distance As Double

    Dim ret As Boolean
    
    newv(0) = vx: newv(1) = vy
    ret = False
    
    addDonut2 vx, vy
    ddd = ent.Coordinates
    count = UBound(ddd)
    
    x1 = ddd(count - 1)
    y1 = ddd(count)
    
    For ii = 0 To count Step 2
    
        x2 = ddd(ii)
        y2 = ddd(ii + 1)
        
        'ff = (y2 - y1) / (x2 - x1)
        
        'ff1 = (vy - y1) / (vx - x1)
        'ff2 = (y2 - vy) / (x2 - vx)
        
        f = getTan(x1, x2, y1, y2)
        f1 = getTan(x1, vx, y1, vy)
        f2 = getTan(vx, x2, vy, y2)
         
        fd = XYDistance(x1, y1, x2, y2)
        fd1 = XYDistance(x1, y1, vx, vy) + XYDistance(x2, y2, vx, vy)
        
        addDonut x2, y2
        addDonut x1, y1
        
        If Abs(fd - fd1) < 0.01 Then
            
            'addDonut x2, y2
            'addDonut x1, y1
            addDonut vx, vy
            ent.addVertex ii / 2, newv
            ent.Update
            ret = True
        
        End If
        
        x1 = x2
        y1 = y2
        
    Next ii
    
    Debug.Print " "
    
    If ret = False Then
        newv(0) = 0
        newv(1) = 0
    
    End If
    
    
    addVertex = newv
        
        
    
End Function

Function getTan(ByVal x1 As Double, ByVal x2 As Double, ByVal y1 As Double, ByVal y2 As Double) As Double
    
    Dim t As Double
    
    If x1 = x2 Then
    
        t = 1.79769313486232E+307
    Else
        t = (y2 - y1) / (x2 - x1)
    
    End If
    
    getTan = t

End Function






Function XYDistance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double

  'Returns the distance between two points

  Dim dblDist As Double

  Dim dblXSl As Double

  Dim dblYSl As Double

  Dim varErr As Variant

  On Error GoTo Err_Control

  'Calc distance

  dblXSl = (x1 - x2) ^ 2

  dblYSl = (y1 - y2) ^ 2

  dblDist = Sqr(dblXSl + dblYSl)

  'Return Distance

  XYDistance = dblDist

Exit_Here:

  Exit Function

Err_Control:

  Select Case Err.Number

  'Add additional Case selections here

    Case Else

    MsgBox Err.Description

    Err.Clear

    Resume Exit_Here

  End Select

End Function



Function addDonut(ByVal x1 As Double, ByVal y1 As Double)

    
    Dim circleObj As AcadCircle
    Dim centerPoint(0 To 2) As Double
    
 
    If SHOW_CIRCLE = True Then
        centerPoint(0) = x1
        centerPoint(1) = y1
        
        Set circleObj = ThisDrawing.ModelSpace.AddCircle(centerPoint, 1)
        circleObj.color = acBlue
        
        circleObj.Update
    End If

End Function

Function addDonut2(ByVal x1 As Double, ByVal y1 As Double)

    Dim circleObj As AcadCircle
    Dim centerPoint(0 To 2) As Double
    
 
    If SHOW_CIRCLE = True Then
        centerPoint(0) = x1
        centerPoint(1) = y1
        
        Set circleObj = ThisDrawing.ModelSpace.AddCircle(centerPoint, 0.5)
        circleObj.color = acRed
        circleObj.Update
    End If
    

End Function

Function addDonut3(ByVal x1 As Double, ByVal y1 As Double)

    Dim circleObj As AcadCircle
    Dim centerPoint(0 To 2) As Double
    
    If SHOW_CIRCLE = True Then
        centerPoint(0) = x1
        centerPoint(1) = y1
        
        Set circleObj = ThisDrawing.ModelSpace.AddCircle(centerPoint, 0.5)
        circleObj.color = acYellow
        
        circleObj.Update
    End If
    

End Function
