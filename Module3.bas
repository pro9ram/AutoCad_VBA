Attribute VB_Name = "Module3"
Sub Step1()
    Dim arcObjs() As AcadEntity
    Dim plroad() As AcadEntity
    Dim plObjs() As AcadEntity
    Dim ddd() As Double
    
    Dim ent As AcadEntity
    Dim ent2 As AcadEntity
    
    Dim vx, vy As Double
    Dim count As Integer
    
    Dim filletpoint(0 To 7) As Double   '8
    
    Dim dout1() As Double
    Dim dout2() As Double
    
    Dim df1() As Double
    Dim df2() As Double
    
    Dim dresult() As Double
    Dim plResult As AcadLWPolyline
    
    
    
    plroad = selectPolylineObs
    arcObjs = selectArcObs
    
    plObjs = arc2lines(arcObjs)
    'ddd = plObjs(0).Coordinates
    
    index = 0
    For i = 0 To UBound(plroad)
        Set ent = plroad(i)
        
        For j = 0 To UBound(plObjs)
        
            Set ent2 = plObjs(j)
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
        
        Next
    Next
    
    
    
    dout1 = searchOutLine(plroad(0), plObjs(0), plObjs(1))
    
    ThisDrawing.ModelSpace.AddLightWeightPolyline (dout1)
    
    dout2 = searchInnerLine(plroad(1), plObjs(0), plObjs(1))
    
    ThisDrawing.ModelSpace.AddLightWeightPolyline (dout2)
    Debug.Print " "
    
    Set ent = plObjs(0)
    df1 = ent.Coordinates
    Set ent = plObjs(1)
    df2 = ent.Coordinates
    
    dresult = mergeAll(dout1, df1, dout2, df2)
    
    
    Set plResult = ThisDrawing.ModelSpace.AddLightWeightPolyline(dresult)
    plResult.Closed = True
    plResult.Update
    
    
    
End Sub

Function mergeAll(d1() As Double, d2() As Double, d3() As Double, d4() As Double) As Double()

    Dim dout() As Double
    Dim count As Integer
    
    
    count = UBound(d1) + UBound(d2) + UBound(d3) + UBound(d4) + 3 - 2
    ReDim dout(count) As Double
    
    idx = 0
    
    For i = 0 To UBound(d1)
        dout(idx) = d1(i)
        idx = idx + 1
    Next
    
    If d2(0) = d1(UBound(d1) - 1) Then
        For i = 2 To UBound(d2)
            dout(idx) = d2(i)
            idx = idx + 1
        Next
    Else
        For i = UBound(d2) - 1 To 0 Step -2
            dout(idx) = d2(i)
            idx = idx + 1
            dout(idx) = d2(i + 1)
            idx = idx + 1
        Next
    End If
    
    
    If d3(0) = d2(UBound(d2) - 1) Then
        For i = 2 To UBound(d3)
            dout(idx) = d3(i)
            idx = idx + 1
        Next
    Else
        For i = UBound(d3) - 1 To 0 Step -2
            dout(idx) = d3(i)
            idx = idx + 1
            dout(idx) = d3(i + 1)
            idx = idx + 1
        Next
    End If
    
     
    If d4(0) = d3(UBound(d3) - 1) Then
        For i = 2 To UBound(d4) - 2
            dout(idx) = d4(i)
            idx = idx + 1
        Next
    Else
        For i = UBound(d4) - 1 To 2 Step -2
            dout(idx) = d4(i)
            idx = idx + 1
            dout(idx) = d4(i + 1)
            idx = idx + 1
        Next
    End If
    
    
    mergeAll = dout
    
    

End Function



Function searchInnerLine(plobj As AcadEntity, f1obj As AcadEntity, f2obj As AcadEntity) As Double()
    searchInnerLine = searchCrossLine(plobj, f1obj, f2obj, True)
End Function

Function searchOutLine(plobj As AcadEntity, f1obj As AcadEntity, f2obj As AcadEntity) As Double()

    searchOutLine = searchCrossLine(plobj, f1obj, f2obj, False)

End Function


Function searchCrossLine(plobj As AcadEntity, f1obj As AcadEntity, f2obj As AcadEntity, isinner As Boolean) As Double()

    Dim dpl() As Double
    Dim df1() As Double
    Dim df2() As Double
    Dim point As Double
    
    Dim f1idx, f1idxn As Integer    '필렛1 접점 인덱스
    Dim f2idx, f2idxn As Integer    '필렛2 접점 인덱스
    
    Dim index1, index1n As Integer   '폴리라인 접점 인덱스
    
    
    Dim dret() As Double    '리턴좌표
    
    
    'Set ret = Nothing
    
    dpl = plobj.Coordinates '폴리라인
    df1 = f1obj.Coordinates '필렛1
    df2 = f2obj.Coordinates '필렛2

    x = df1(0)
    y = df1(1)
    
    index1 = searchIndex(dpl, x, y)
    
    If index1 = -1 Then
        count = UBound(df1)
        
        x = df1(count - 1)
        y = df1(count)
        
        index1 = searchIndex(dpl, x, y)
        f1idx = count - 1
        f1idxn = count - 3
    Else
        f1idx = 0
        f1idxn = 2
        
    
    End If
    
    x = df2(0)
    y = df2(1)
    
    index2 = searchIndex(dpl, x, y)
    
    If index2 = -1 Then
        count = UBound(df2)
        
        x = df2(count - 1)
        y = df2(count)
        
        index2 = searchIndex(dpl, x, y)
        f2idx = count - 1
        f2idxn = count - 3
    Else
        f2idx = 0
        f2idxn = 2
    End If
    


    If index1 + 2 > UBound(dpl) Then
        index1n = index1 - 2
    Else
        index1n = index1 + 2
    End If


    cx = dpl(index1)
    cy = dpl(index1 + 1)
    
    x1 = dpl(index1n)
    y1 = dpl(index1n + 1)
    
    x2 = df1(f1idxn)
    y2 = df1(f1idxn + 1)

    addDonut2 cx, cy
    addDonut2 x1, y1
    addDonut2 x2, y2
    


    rad = Abs(Atn((y2 - cy) / (x2 - cx)) - Atn((y1 - cy) / (c1 - cx)))
    
    If rad < 1.571 = isinner Then '삭제영역에 해당함
        If index2 > index1 Then     '올바른 방향임
            count = index2 - index1 + 1
            ReDim dret(0 To count) As Double
            
            For i = index1 To index2 + 1
                dret(i - index1) = dpl(i)
            Next
        
        Else
            count = index1 - index2 + 1
            ReDim dret(0 To count) As Double
            
            For i = index2 To index1 + 1
                dret(i - index2) = dpl(i)
            Next
        
        
        End If
        
    Else
        If index2 > index1 Then
        
            count = index1 + UBound(dpl) - index2 + 2   'y좌표 2개라서 +2
            ReDim dret(0 To count) As Double
            idx = 0
        
            For i = index2 To UBound(dpl)   '개수만큼 찾기때문에 y가 포함됨 +1 안함
                dret(idx) = dpl(i)
                idx = idx + 1
            Next
                    
            For i = 0 To index1 + 1 'y좌표 포함해야하므로 +1
                dret(idx) = dpl(i)
                idx = idx + 1
            Next
            
            
            Debug.Print
            
            
            
        Else
        
        End If
    
    End If
    
    
    
    searchCrossLine = dret
    

End Function

Function searchIndex(d() As Double, ByVal x As Double, ByVal y As Double) As Integer

    Dim index As Integer
    
    index = -1
    
    For i = 0 To UBound(d) Step 2
    
        If d(i) = x And d(i + 1) = y Then
            index = i
        End If
    
    Next


    searchIndex = index
    
End Function





Function selectArcObs() As AcadEntity()
    Dim UsersSelection As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim arcObjs() As AcadEntity
 
     With ThisDrawing
        On Error Resume Next
        
        .SelectionSets("CurrentSelection").Delete
        Set UsersSelection = .SelectionSets.Add("CurrentSelection")
        
        FilterType(0) = 0
        FilterData(0) = "Arc"
        
        UsersSelection.SelectOnScreen FilterType, FilterData
        
        
        count = getEntitySize(UsersSelection, "AcDbArc")
        
        ReDim arcObjs(0 To count - 1) As AcadEntity
        
        For i = 0 To count
            Set arcObjs(i) = UsersSelection.Item(i)
        Next i
        
        
    End With
    
    selectArcObs = arcObjs

End Function


Function selectPolylineObs() As AcadEntity()
    Dim UsersSelection As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim arcObjs() As AcadEntity
 
     With ThisDrawing
        On Error Resume Next
        
        .SelectionSets("CurrentSelection").Delete
        Set UsersSelection = .SelectionSets.Add("CurrentSelection")
        
        FilterType(0) = 0
        FilterData(0) = "LWPolyline"
        
        UsersSelection.SelectOnScreen FilterType, FilterData
        
        
        count = getEntitySize(UsersSelection, "AcDbPolyline")
        
        ReDim arcObjs(0 To count - 1) As AcadEntity
        
        For i = 0 To count
            Set arcObjs(i) = UsersSelection.Item(i)
        Next i
        
        
    End With
    
    selectPolylineObs = arcObjs

End Function

Function addVertex(ent As AcadEntity, ByVal vx As Double, ByVal vy As Double) As Double()

    Dim newv(0 To 1) As Double
    Dim ddd() As Double
    Dim count As Integer
    Dim distance As Double

    Dim ret As Boolean
    
    newv(0) = vx: newv(1) = vy
    ret = False
    
    'addDonut vx, vy
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
        

        
        If Abs(f1 - f2) < 0.00001 And Abs(fd - fd1) < 0.00001 Then
            
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

Sub SelectRawData()
    Dim i, ii, x, y As Integer
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim UsersSelection As AcadSelectionSet
    Dim selObj() As AcadEntity
    Dim ret As Integer
    Dim arcObjs() As AcadEntity
    Dim ddd() As Double
    Dim x2, y2 As Variant
    Dim count As Integer
    Dim newv(0 To 1) As Double
    
    
    With ThisDrawing
        On Error Resume Next
        .SelectionSets("CurrentSelection").Delete
        Set UsersSelection = .SelectionSets.Add("CurrentSelection")
        
        
        FilterType(0) = 0
        FilterData(0) = "LWPolyline"
        
        UsersSelection.SelectOnScreen FilterType, FilterData
        
        
        count = getEntitySize(UsersSelection, "AcDbPolyline") - 1 'AcDbArc
        
        ReDim selObj(0 To count) As AcadEntity
        
        For i = 0 To count
            Set selObj(i) = UsersSelection.Item(i)
            Debug.Print " "
             
        Next i
        
        
    End With
    
    
    vx = 1194.35534891423
    vy = 2999.85414787302
    newv(0) = vx: newv(1) = vy
    
    For i = 0 To count
        ddd = selObj(i).Coordinates
        count = UBound(ddd)
        
        x1 = ddd(count - 1)
        y1 = ddd(count)
        
        For ii = 0 To count Step 2
        
            x2 = ddd(ii)
            y2 = ddd(ii + 1)
            
            f = (y2 - y1) / (x2 - x1)
            
            f1 = (vy - y1) / (vx - x1)
            f2 = (y2 - vy) / (x2 - vx)
            
            If Abs(f1 - f2) < 0.00001 Then
                addDonut x2, y2
                addDonut x1, y1
                
                selObj(i).addVertex ii / 2, newv
                selObj(i).Update
            
            End If
            
            x1 = x2
            y1 = y2
            
        Next ii
    Next i
    
    
    
    Debug.Print " "
    
    
End Sub


Function addDonut(ByVal x1 As Double, ByVal y1 As Double)

    Dim circleObj As AcadCircle
    Dim centerPoint(0 To 2) As Double
    
    centerPoint(0) = x1
    centerPoint(1) = y1
    
    Set circleObj = ThisDrawing.ModelSpace.AddCircle(centerPoint, 1)
    circleObj.Update
    
    'Donut x1, y1, 5
    

End Function

Function addDonut2(ByVal x1 As Double, ByVal y1 As Double)

    Dim circleObj As AcadCircle
    Dim centerPoint(0 To 2) As Double
    
    centerPoint(0) = x1
    centerPoint(1) = y1
    
    Set circleObj = ThisDrawing.ModelSpace.AddCircle(centerPoint, 0.5)
    circleObj.Update
    
    'Donut x1, y1, 5
    

End Function



Function getEntitySize(ss As AcadSelectionSet, text As String) As Integer

    Dim count As Integer
    
    count = 0

    For Each ent In ss
        If ent.ObjectName = text Then
            count = count + 1
        End If
    Next ent
        
    getEntitySize = count
        


End Function

Function getArcSize(ss As AcadSelectionSet) As Integer

    Dim count As Integer

    count = 0

    For Each ent In ss
        If ent.ObjectName = "AcDbArc" Then
            count = count + 1
        End If
    Next ent
        
    getArcSize = count
        


End Function


Function arc2lines(arcs() As AcadEntity) As AcadEntity()

    Dim i As Integer
    Dim myArc As AcadArc
    Dim objSel As AcadEntity
    Dim myPL As AcadLWPolyline
    Dim mypolarpoint
    Dim bulge() As Double
    Dim legs As Integer
    Const PI = 3.14159265358979
    Dim delta As Double
    Dim count As Integer
    Dim retObj() As AcadEntity
    
    count = UBound(arcs)
    ReDim retObj(0 To count) As AcadEntity
        
    For i = 0 To UBound(arcs)
        Set myArc = arcs(i)
 
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
        For x = 2 To UBound(points) - 2 Step 2
            mypolarpoint = ThisDrawing.Utility.PolarPoint(myArc.Center, myArc.StartAngle + adir, myArc.Radius)
            adir = adir + ang
            points(x) = mypolarpoint(0)
            points(x + 1) = mypolarpoint(1)
        Next x
        points(UBound(points) - 1) = myArc.endPoint(0)
        points(UBound(points) - 0) = myArc.endPoint(1)
        Set retObj(i) = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
        retObj(i).Update
        
       
        
        myArc.Delete
        
    Next i
    
    arc2lines = retObj

End Function
 


Sub Test()
    
    Dim SelPl(0 To 1) As AcadEntity
    Dim splyObj As AcadLWPolyline
    Dim newObj As AcadLWPolyline
    
    
    Dim UsersSelection As AcadSelectionSet
    
    
    Dim DrawingSelected As AcadEntity 'delete the selection set if it already exists
    Dim intPoint12 As Variant
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim splyCoords() As Double
    
    Dim temp As Variant
    
    Dim index As Integer
    
    
    With ThisDrawing
        On Error Resume Next
        .SelectionSets("CurrentSelection").Delete
        'get selection from user
        MsgBox "Select objects! Hit Enter to finish!"
        Set UsersSelection = .SelectionSets.Add("CurrentSelection")
        
        FilterType(0) = 0
        FilterData(0) = "LWPolyline"
        
        UsersSelection.SelectOnScreen FilterType, FilterData
        Set SelPl(0) = UsersSelection.Item(0)
        Set SelPl(1) = UsersSelection.Item(1)
        
        intPoint12 = SelPl(0).IntersectWith(SelPl(1), acExtendNone)
         
        
        splyCoords = SelPl(0).Coordinates
        splyCoords(7) = 2000
        
        SelPl(0).Update
        
        
        Set newObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(splyCoords)
        newObj.Closed = True
        
        
        Debug.Print " "
        
        
    End With

End Sub

