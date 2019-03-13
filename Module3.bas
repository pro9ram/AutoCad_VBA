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
    
    plroad = selectPolylineObs
    arcObjs = selectArcObs
    
    plObjs = arc2lines(arcObjs)
    'ddd = plObjs(0).Coordinates
    
    
    For i = 0 To UBound(plroad)
        Set ent = plroad(i)
        For j = 0 To UBound(plObjs)
        
            Set ent2 = plObjs(j)
            ddd = ent2.Coordinates
            
            count = UBound(ddd)
            
            vx = ddd(count - 1)
            vy = ddd(count)
            
            addVertex ent, vx, vy
            
            vx = ddd(0)
            vy = ddd(1)
        
            addVertex ent, vx, vy
        
        Next
    Next
    
    
    
    
    Debug.Print " "
    
    
End Sub


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

Function addVertex(ent As AcadEntity, ByVal vx As Double, ByVal vy As Double)

    Dim newv(0 To 1) As Double
    Dim ddd() As Double
    Dim count As Integer
    'Dim distance As Double

    'vx = 1194.35534891423
    'vy = 2999.85414787302
    newv(0) = vx: newv(1) = vy
    
    'addDonut vx, vy
    ddd = ent.Coordinates
    count = UBound(ddd)
    
    x1 = ddd(count - 1)
    y1 = ddd(count)
    
    For ii = 0 To count Step 2
    
        x2 = ddd(ii)
        y2 = ddd(ii + 1)
        
        
        If x1 = x2 Then
            Debug.Print " "
            
        End If
        
        f = (y2 - y1) / (x2 - x1)
        
        f1 = (vy - y1) / (vx - x1)
        f2 = (y2 - vy) / (x2 - vx)
        
         
         
        fd = XYDistance(x1, y1, x2, y2)
        fd1 = XYDistance(x1, y1, vx, vy) + XYDistance(x2, y2, vx, vy)
        

        
        If Abs(f1 - f2) < 0.00001 And Abs(fd - fd1) < 0.00001 Then
            
            addDonut x2, y2
            addDonut x1, y1
            
            ent.addVertex ii / 2, newv
            ent.Update
        
        End If
        
        x1 = x2
        y1 = y2
        
    Next ii
    
    
    
    
    Debug.Print " "
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
    
    Dim Index As Integer
    
    
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

