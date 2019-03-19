Attribute VB_Name = "filletModule"
Private Const PI As Double = 3.14159265358979
Private Const PI_2 As Double = 1.5707963267949

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
    
    If StartAngle > 0 And EndAngle > 0 Then
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
        If StartAngle < PI_2 Then
            tmp = StartAngle
            StartAngle = EndAngle
            EndAngle = tmp
        End If
    ElseIf StartAngle < 0 And EndAngle > 0 Then
        If EndAngle > PI_2 Then
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
    numOfSegments = CInt(myArc.ArcLength) ' length of segment = 1, last segment = remainder
    ReDim points(0 To 2 * numOfSegments + 1)
    ang = 1 / myArc.Radius
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
