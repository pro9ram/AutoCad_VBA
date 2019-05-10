Attribute VB_Name = "topology"

Function getFarPoint1(list() As clsPoint, pt As clsPoint) As clsPoint

    Dim p As clsPoint
    Dim dd As Double
    Dim max As Double
    Dim idx As Integer
    
    max = 0
    
    For i = 0 To UBound(list)
    
        p = list(0)
        dd = pt.distance(p)
        
        If dd > max Then
            max = dd
            idx = i
            
        End If
        
    Next


    getFarPoint1 = list(idx)


End Function

Function getFarPoint2(list As Variant, pt As clsPoint) As clsPoint

    Dim p As New clsPoint
    Dim dd As Double
    Dim idx As Integer
    Dim x As Double
    Dim y As Double
    
    max = 0
    
     
    
    For i = 0 To UBound(list) Step 3
    
        x = list(i)
        y = list(i + 1)
    
        p.initXy x, y 'x, y
        
        
        dd = pt.distance(p)
        
        If dd > max And dd < FAR_OVER Then
            max = dd
            idx = i
            
        End If
        
    Next

    p.initXy list(idx), list(idx + 1)


    Set getFarPoint2 = p


End Function





'intersectwith에서 잘못된점이 있기때문에 생성
Function getFarPoint2Ex(ent As AcadEntity, list As Variant, pt As clsPoint) As clsPoint

    Dim p As New clsPoint
    Dim dd As Double
    Dim idx As Integer
    Dim x As Double
    Dim y As Double
    Dim exists As Boolean
    
    Dim cpt As New clsPoint
    Dim newpr As New clsPolygonReader
    
    Dim ddd() As Double
    Dim d1 As Double
    Dim d2 As Double
    
    Dim tmp As AcadLWPolyline
    
    max = 0
    
     
    newpr.init ent
    newpr.setIndexByPoint pt
    
    'addDonut5 pt.x, pt.y
    
    For i = 0 To UBound(list) Step 3
    
        x = list(i)
        y = list(i + 1)
    
        
    
        exists = isContains(ent, x, y)
    
        If exists = True Then
            
            p.initXy x, y 'x, y
            'addDonutA x, y
            
            newpr.setEIndexByPoint p
            
            ddd = newpr.getLongLine2
            'Set tmp = ThisDrawing.ModelSpace.AddLightWeightPolyline(ddd)
            d1 = getArrayDistance(ddd)
            
            ddd = newpr.getShortLine2
            'Set tmp = ThisDrawing.ModelSpace.AddLightWeightPolyline(ddd)
            d2 = getArrayDistance(ddd)
            
            
            If d1 > d2 Then
                dd = d2
            Else
                dd = d1
            End If
            
            'Debug.Print ">>>>>> " & d1 & ", " & d2 & ">>>>> " & dd
            
            'dd = pt.distance(p)
            
            If dd > max And dd < FAR_OVER Then
                max = dd
                idx = i
                
            End If
        End If
        
    Next

    p.initXy list(idx), list(idx + 1)


    Set getFarPoint2Ex = p


End Function



Function getNearPoint2(list As Variant, pt As clsPoint) As clsPoint

    Dim p As New clsPoint
    Dim dd As Double
    Dim idx As Integer
    Dim x As Double
    Dim y As Double
    
    min = 99999999
    
    For i = 0 To UBound(list) Step 3
    
        x = list(0)
        y = list(1)
    
        p.initXy list(i), list(i + 1) 'x, y
        
        dd = pt.distance(p)
        
        If dd < min Then
            min = dd
            idx = i
            
        End If
        
    Next

    p.initXy list(idx), list(idx + 1)


    Set getNearPoint2 = p


End Function


Function getLineByEntity(ent As AcadEntity, ByVal vx As Double, ByVal vy As Double) As AcadLine

    Dim newv(0 To 1) As Double
    Dim ddd() As Double
    Dim count As Integer
    Dim distance As Double

    Dim ret As Boolean
    Dim newp() As Double
    Dim line As AcadLine
    
    newv(0) = vx: newv(1) = vy
    ret = False
    
    addDonutR vx, vy
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
        
        If Abs(fd - fd1) < PD_TOL Then
            
            newp = crossXY(vx, vy, x1, y1, x2, y2)
            
            
            Dim sp(0 To 2) As Double
            Dim ep(0 To 2) As Double
            
            
            sp(0) = x1
            sp(1) = y1
            ep(0) = x2
            ep(1) = y2
            
            Set line = ThisDrawing.ModelSpace.addLine(sp, ep)
      
            
            addDonutY vx, vy
            
            Exit For
        
        End If
        
        x1 = x2
        y1 = y2
        
    Next ii
    
    If ret = False Then
        newv(0) = 0
        newv(1) = 0
    
    End If
        
   
    
    
    Set getLineByEntity = line
        
        
    
End Function


Function searchCrossVertext(src As AcadEntity, dst As AcadEntity) As Double()

    
    searchCrossVertext = src.IntersectWith(dst, acExtendNone)   'find vertex
     
    
End Function

Function searchCrossVertextEx(src As AcadEntity, dst As AcadEntity) As Double()

    Dim size As Integer
    Dim d3 As Variant
    Dim ret() As Double
    Dim ret2() As Double
    Dim vx, vy As Variant
    Dim b1, b2 As Boolean
    Dim idx As Integer
    Dim cnt As Integer
    
    idx = 0
    
    d3 = src.IntersectWith(dst, acExtendNone)   'find vertex
    
    size = UBound(d3)
    ReDim ret(size) As Double
    
    For i = 0 To size Step 3
    
        vx = d3(i)
        vy = d3(i + 1)
        
        b1 = isContains(src, CDbl(vx), CDbl(vy))
        b2 = isContains(dst, CDbl(vx), CDbl(vy))
        
        If b1 = True And b2 = True Then
        
            ret(idx) = vx
            ret(idx + 1) = vy
            ret(idx + 2) = 0
            
            idx = idx + 3
            cnt = cnt + 1
        
        End If
    
    Next
    
    ReDim ret2(0 To cnt * 3 - 1) As Double
    
    For i = 0 To cnt * 3 - 1 Step 3
    
        ret2(i) = ret(i)
        ret2(i + 1) = ret(i + 1)
        ret2(i + 2) = ret(i + 2)
    
    Next
    
    
    
    searchCrossVertextEx = ret2
    
End Function



Function isContains(src As AcadEntity, vx As Double, vy As Double) As Boolean

    Dim ddd() As Double
    Dim x, y As Double
    Dim ret As Boolean
    Dim b As Boolean
    
    ret = False
    ddd = src.Coordinates
   
    
    For i = 0 To UBound(ddd) - 1 Step 2
    
        x = ddd(i)
        y = ddd(i + 1)
        
        
        'addDonut5 x, y
        
        b = isEqualsDouble(CDbl(x), y, vx, vy)
        
        'If x = vx And y = vy Then
        If b = True Then
            ret = True
            Exit For
        End If
    
    
    Next
    
    isContains = ret


End Function


Function getArrayDistance(d() As Double) As Double

    Dim ret As Double
    Dim size As Integer
    Dim x1, y1, x2, y2 As Double
        
    ret = 0
    size = UBound(d)
    
    If size = 0 Then
        Return
    End If

    x1 = d(0)
    y1 = d(1)
    
    For i = 2 To size Step 2
    
        x2 = d(i)
        y2 = d(i + 1)
        
        ret = ret + getDistance(x1, y1, x2, y2)
        
        x1 = x2
        y1 = y2
       
    Next
    
    getArrayDistance = ret

End Function

Function trimArrayDouble(ddd() As Double) As Double()

    Dim size As Integer
    Dim idx As Integer
    Dim ret() As Double
    
    idx = -1
    size = UBound(ddd)
    
    
    For i = 0 To size - 1
    
        If ddd(i) = 0 And ddd(i + 1) = 0 Then
            idx = i
            Exit For
        End If
        
    Next
    

    
    If idx = -1 Then
    
        trimArrayDouble = ddd
    Else
    
    
        ReDim ret(idx - 1) As Double
        
        For i = 0 To idx - 1
        
            ret(i) = ddd(i)
        
        Next
    
        trimArrayDouble = ret
    End If
    


End Function


Function trimPolyline(ent As AcadLWPolyline) As AcadLWPolyline

    Dim retent As AcadLWPolyline

    Dim ddd() As Double
    Dim ddd2() As Double
    Dim ddd3() As Double
    
    Dim x1, y1, x2, y2 As Double
    Dim size As Integer
    Dim idx As Integer
    
    idx = 0
    ddd = ent.Coordinates
    size = UBound(ddd)
    
    ReDim ddd2(size) As Double
    
    x1 = ddd(0)
    y1 = ddd(1)
    
    ddd2(0) = x1
    ddd2(1) = y1
    
    idx = idx + 2
    
    For i = 2 To UBound(ddd) Step 2
    
        x2 = ddd(i)
        y2 = ddd(i + 1)
    
        If x1 = x2 And y1 = y2 Then
            Debug.Print ""
            
        Else
            ddd2(idx) = x2
            ddd2(idx + 1) = y2
        
            idx = idx + 2
        End If
        
        x1 = x2
        y1 = y2
    
    
    Next
    
    
    ddd3 = trimArrayDouble(ddd2)
    
    Debug.Print ""
    
    Set retent = ThisDrawing.ModelSpace.AddLightWeightPolyline(ddd3)
    retent.Closed = True
    
    Set trimPolyline = retent

End Function


Function polyline2lwpolyline(src As AcadEntity) As AcadLWPolyline


    Dim size As Integer
    Dim ret As AcadLWPolyline
    Dim ddd() As Double
    Dim ddd2() As Double
    Dim idx As Integer
    
    ddd = src.Coordinates
    size = UBound(ddd)
    
    
    ReDim ddd2(size * 2 / 3) As Double
    
    For i = 0 To size Step 3
    
        ddd2(idx) = ddd(i)
        ddd2(idx + 1) = ddd(i + 1)
        
        idx = idx + 2
    
    Next i
    
    Set ret = ThisDrawing.ModelSpace.AddLightWeightPolyline(ddd2)
    If ret.Closed = False Then
        ret.Closed = True
    End If
    
    

End Function


Function d3ToPolyline(ddd As Variant) As AcadLWPolyline


    Dim size As Integer
    Dim ret As AcadLWPolyline
    Dim ddd2() As Double
    Dim idx As Integer
    
     
    size = UBound(ddd)
    
    
    ReDim ddd2(size * 2 / 3) As Double
    
    For i = 0 To size Step 3
    
        ddd2(idx) = ddd(i)
        ddd2(idx + 1) = ddd(i + 1)
        
        idx = idx + 2
    
    Next i

    Set ret = ThisDrawing.ModelSpace.AddLightWeightPolyline(ddd2)
    If ret.Closed = False Then
        ret.Closed = True
    End If
    
    Set d3ToPolyline = ret
    

End Function



Function expoldeRegion(reg As Variant)

    Dim has As Boolean
    Dim conv As New ClsConverter
    Dim Punto As Variant

    has = hasSubRegion(reg)
    
    If has = True Then
    
        Punto = reg.Explode
        reg.Delete
        
        For i = 0 To UBound(Punto)
            conv.reg2polyline Punto(i)
        Next i
        
        For i = 0 To UBound(Punto)
            Punto(i).Delete
        Next i
    
    Else
    
        conv.reg2polyline reg
        reg.Delete
    
    End If
    


End Function


Function hasSubRegion(reg As Variant) As Boolean

    Dim arr As Variant
    Dim size As Integer
    Dim ret As Boolean

    ret = True

    arr = reg.Explode
    size = UBound(arr)
    
    If size > 0 Then
        If TypeOf arr(0) Is AcadLine Then
            ret = False
        
        End If
    End If
    
    
    For i = 0 To size
        arr(i).Delete
    
    Next
    
    hasSubRegion = ret

End Function


Function sublwpolyline(sx As Double, sy As Double, ent As AcadLWPolyline) As AcadLWPolyline


    Dim ddd() As Double
    Dim size As Integer
    Dim sidx As Integer
    Dim x, y As Double
    
    Dim ddd2() As Double
    
    Dim retent As AcadLWPolyline
    
    ddd = ent.Coordinates
    size = UBound(ddd)
    
    sidx = 0
    
    For i = 0 To size Step 2
        x = ddd(i)
        y = ddd(i + 1)
        
        If sx = x And sy = y Then
            Exit For
        End If
    
        sidx = sidx + 2
    Next i
    
    
    ReDim ddd2(size - sidx) As Double
    
    For i = sidx To size Step 2
        ddd2(i - sidx) = ddd(i)
        ddd2(i - sidx + 1) = ddd(i + 1)
        
    Next i
    
    Set retent = ThisDrawing.ModelSpace.AddLightWeightPolyline(ddd2)
    
    
    Set sublwpolyline = retent


End Function





Sub topologyTest()

    Dim p1 As New clsPoint
    Dim p2 As New clsPoint
    Dim distance As Double
    
    p1.initXy 0, 0
    p2.initXy 1, 1
    
    
    distance = p1.distance(p2)
    
    

End Sub


