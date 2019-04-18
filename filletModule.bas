Attribute VB_Name = "filletModule"
Public Const PI As Double = 3.14159265358979
Public Const PI_2 As Double = 1.5707963267949






Function arc2line(myArc As AcadArc) As AcadEntity

    Dim retObj As AcadEntity
    Dim delta As Double
    
    delta = myArc.EndAngle - myArc.StartAngle
    If delta < 0 Then delta = delta + (2 * PI)
    Dim numOfSegments As Integer
    Dim points() As Double
    'adjust below for reality
    numOfSegments = CInt(myArc.ArcLength * 10) ' length of segment = 1, last segment = remainder
    ReDim points(0 To 2 * numOfSegments + 1)
    ang = 1 / (myArc.Radius * 10)
    points(0) = myArc.startPoint(0)
    points(1) = myArc.startPoint(1)
    adir = ang
    For x = 2 To UBound(points) - 2 Step 2
        mypolarpoint = ThisDrawing.Utility.PolarPoint(myArc.Center, myArc.StartAngle + adir, myArc.Radius)
        adir = adir + ang
        points(x) = mypolarpoint(0)
        points(x + 1) = mypolarpoint(1)
    Next x
    points(UBound(points) - 1) = myArc.endPoint(0)
    points(UBound(points) - 0) = myArc.endPoint(1)
    Set retObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
    
    Set arc2line = retObj

End Function


Function addCrossVertext(src As AcadEntity, dst As AcadEntity) As Variant

    Dim ret As Variant
    Dim size As Integer
    
    ret = src.IntersectWith(dst, acExtendOtherEntity)   'find vertex
    size = 0
    
    
    size = UBound(ret)
    
    If size > 0 Then
        For i = 0 To size Step 3
            addDonut2 ret(i), ret(i + 1)
            
            addVertex dst, ret(i), ret(i + 1)     ' add vertext
        Next
    Else
    
    End If
    
    addCrossVertext = ret

End Function


Function findVertex(ent As AcadEntity, ent2 As AcadEntity) As Double()

    Dim ddd() As Double
    Dim filletpoint(0 To 7) As Double   '8
        
    ddd = ent2.Coordinates
    
    count = UBound(ddd)
    
    vx = ddd(count - 1)
    vy = ddd(count)
    
    newp = addVertex(ent, vx, vy)
    
    If newp(0) <> 0 And newp(1) <> 0 Then
    
        filletpoint(index) = newp(0)
        index = index + 1
        filletpoint(index) = newp(1)
        index = index + 1
         
    Else
        vx = ddd(0)
        vy = ddd(1)
        
        newp = addVertex(ent, vx, vy)
        
        If newp(0) <> 0 And newp(1) <> 0 Then
        
            filletpoint(index) = newp(0)
            index = index + 1
            filletpoint(index) = newp(1)
            index = index + 1
        End If
    End If
    
    findVertex = newp

End Function



Function getNearestVertex(ent2 As AcadEntity, ref As clsPoint) As Double()

    Dim ent2d() As Double
    Dim x1, y1, x2, y2 As Double
    
    Dim dist1, dist2 As Double
    Dim ret(0 To 1) As Double
    
    ent2d = ent2.Coordinates
    count = UBound(ent2d)
    
    
    x1 = ent2d(0)
    y1 = ent2d(1)
    
    x2 = ent2d(count - 1)
    y2 = ent2d(count)
    
    
    'addDonutA x1, y1
    'addDonutA x2, y2
        
    dist1 = getDistance(x1, y1, ref.x, ref.y)
    dist2 = getDistance(x2, y2, ref.x, ref.y)
    
    If dist1 > dist2 Then
        ret(0) = x2
        ret(1) = y2
    Else
        ret(0) = x1
        ret(1) = y1
    End If
    
    getNearestVertex = ret


End Function

Function findVertex2(ent As AcadEntity, ent2 As AcadEntity, pt As clsPoint) As Double()

    Dim filletpoint(0 To 7) As Double   '8
    Dim nv() As Double
    
    nv = getNearestVertex(ent2, pt)
    vx = nv(0)
    vy = nv(1)
    
    newp = addVertex2(ent, vx, vy)
    
    If newp(0) <> 0 And newp(1) <> 0 Then
    
        filletpoint(index) = newp(0)
        index = index + 1
        filletpoint(index) = newp(1)
        index = index + 1
   
    End If
    
    findVertex2 = newp

End Function


Function addVertex(ent As AcadEntity, ByVal vx As Double, ByVal vy As Double) As Double()

    Dim newv(0 To 1) As Double
    Dim ddd() As Double
    Dim count As Integer
    Dim distance As Double

    Dim ret As Boolean
    Dim newp() As Double
    
    Dim selidx As Integer
    Dim mindist As Double
    Dim exists As Boolean
    
    newv(0) = vx: newv(1) = vy
    ret = False
    
    
    exists = isContains(ent, vx, vy)
    
    If exists = False Then
        
        mindist = 999999
        
        'addDonutA vx, vy
        ddd = ent.Coordinates
        count = UBound(ddd)
        
        x1 = ddd(count - 1)
        y1 = ddd(count)
        
        For ii = 0 To count Step 2
        
            x2 = ddd(ii)
            y2 = ddd(ii + 1)
            
            F = getTan(x1, x2, y1, y2)
            f1 = getTan(x1, vx, y1, vy)
            f2 = getTan(vx, x2, vy, y2)
             
            fd = XYDistance(x1, y1, x2, y2)
            fd1 = XYDistance(x1, y1, vx, vy) + XYDistance(x2, y2, vx, vy)
            
            'addDonutA x2, y2
            'addDonut x1, y1
            
            
            If Abs(fd - fd1) < mindist Then
                mindist = Abs(fd - fd1)
                newp = crossXY(vx, vy, x1, y1, x2, y2)
                selidx = ii
            
            End If
            
            'If Abs(fd - fd1) < PD_TOL Then
                
                'addDonut x2, y2
                'addDonut x1, y1
                
                
            '    newp = crossXY(vx, vy, x1, y1, x2, y2)
                
                'addDonut vx, vy
                
            '    newv(0) = newp(0)
            '    newv(1) = newp(1)
            '    addDonut vx, vy
            '    ent.addVertex ii / 2, newv
            '    ent.Update
            '    ret = True
            '    Exit For
            
            'End If
            
            x1 = x2
            y1 = y2
            
        Next ii
        
        If mindist < PD_TOL Then
            newv(0) = newp(0)
            newv(1) = newp(1)
            'addDonutA vx, vy
            ent.addVertex selidx / 2, newv
            ent.Update
            ret = True
        End If
        
        If ret = False Then
            newv(0) = 0
            newv(1) = 0
            Debug.Print "================NOT FOUND================"
        End If
    End If
    
    addVertex = newv
        
        
    
End Function


Function addVertex2(ent As AcadEntity, ByVal vx As Double, ByVal vy As Double) As Double()

    Dim newv(0 To 1) As Double
    Dim ddd() As Double
    Dim count As Integer
    Dim distance As Double

    Dim ret As Boolean
    Dim newp() As Double
    
    Dim mindist As Double
    Dim exists As Boolean
    mindist = 999999
    
    newv(0) = vx: newv(1) = vy
    ret = False
    
    
    exists = isContains(ent, vx, vy)
    If exists = False Then
        'addDonutA vx, vy
        ddd = ent.Coordinates
        count = UBound(ddd)
        
        x1 = ddd(count - 1)
        y1 = ddd(count)
        
        
        
        
        For ii = 0 To count Step 2
        
            x2 = ddd(ii)
            y2 = ddd(ii + 1)
            
            addDonutY x2, y2
            addDonutY x1, y1
            
            F = getTan(x1, x2, y1, y2)
            f1 = getTan(x1, vx, y1, vy)
            f2 = getTan(vx, x2, vy, y2)
             
            fd = XYDistance(x1, y1, x2, y2)
            fd1 = XYDistance(x1, y1, vx, vy) + XYDistance(x2, y2, vx, vy)
            
            
            If Abs(fd - fd1) < mindist Then
                mindist = Abs(fd - fd1)
                xx1 = x1
                yy1 = y1
                xx2 = x2
                yy2 = y2
                iii = ii
            End If
            
            x1 = x2
            y1 = y2
            
        Next ii
        
        
        Debug.Print "===========================>>>> " & mindist & ", " & PD_TOL
        
        
                
            'addDonut x2, y2
            'addDonut x1, y1
            
            
        newp = crossXY(vx, vy, xx1, yy1, xx2, yy2)
        
        'addDonut vx, vy
        
        newv(0) = newp(0)
        newv(1) = newp(1)
        'addDonutA vx, vy
        ent.addVertex iii / 2, newv
        ent.Update
        ret = True
            
        
        
        
        If ret = False Then
            newv(0) = 0
            newv(1) = 0
            Debug.Print "================NOT FOUND================"
        End If
    End If
    
    addVertex2 = newv
        
        
    
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
    Dim dblDist As Double
    Dim dblXSl As Double
    Dim dblYSl As Double

    Dim varErr As Variant
    On Error GoTo Err_Control

    dblXSl = (x1 - x2) ^ 2
    dblYSl = (y1 - y2) ^ 2
    dblDist = Sqr(dblXSl + dblYSl)

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

 
