Attribute VB_Name = "Test3"


Sub test4()

    Dim plroad() As AcadEntity
    Dim ret As Variant
    
    plroad = selectPolylineObs  'select two polylines(one is a source, the ohter is a target entity.
     
    
    ret = addCrossVertext(plroad(0), plroad(1))
    
    
    Set temp = ThisDrawing.ActiveLayer
    Set pp = plroad(0)
    
    Debug.Print "ActiveLayer: " & temp.name
    pp.layer = temp.name
    
    
    

End Sub


Sub test77()

    Dim selec As New ClsSelEntity
    Dim ent As AcadEntity
    Dim pt As New clsPoint
    
    selec.selectPolyline ("Get Polyline 1: ")
    Set ent = selec.getEntity
    Set pt = selec.getPoint
        
    Debug.Print ""
    

End Sub


Sub test78()

    Dim seler As New ClsSelReal
    Dim rad As Double
    Dim test As Variant
    
    rad = seler.selectReal
    
    Debug.Print " "

End Sub

Sub test79()


    Dim seles As New clsSelSets
    Dim ent As AcadEntity
    
    Set ent = seles.selectArc
    

End Sub
 
 
Sub test80()

    newLayer "TEST"

End Sub

 
Sub testSelection()

    Dim newroad As AcadEntity
    Dim basePnt1 As Variant
        
    On Error GoTo errControl
    
Start:
    ThisDrawing.Utility.getEntity newroad, basePnt1, "Get Object1: "
    
    
errControl:
    
    If Err.Description = "Method 'GetEntity' of object 'IAcadUtility' failed" Then
    
        If CheckKey(VK_ESCAPE) = True Then
            Debug.Print "End"
        
            End
        Else
            Debug.Print "Resume"
            Resume
        End If
    Else
        MsgBox Err.Description
    End If
    

End Sub

Sub test56()
    Dim plroad() As AcadEntity
    Dim ret As Variant
    Dim rad As Double
    
    Dim newroad As AcadObject
    Dim basePnt1 As Variant
    
    Dim oldroad As AcadObject
    Dim basePnt2 As Variant
    
    Dim pt1 As New clsPoint
    Dim pt2 As New clsPoint
    
    
    Dim pr1 As New clsPolygonReader
    Dim pr2 As New clsPolygonReader
    
    Dim line1 As AcadLine
    Dim line2 As AcadLine
    Dim np3 As New clsPoint
    Dim result As AcadLine
    
    Dim retArc As AcadEntity
 
    ThisDrawing.Utility.getEntity newroad, basePnt1, "Get Object1: "
    'ret = addCrossVertext(plroad(0), plroad(1))
    
    ThisDrawing.Utility.getEntity oldroad, basePnt2, "Get Object2: "
    'ret = addCrossVertext(plroad(0), plroad(1))
    
    ptIntersection = searchCrossVertext(newroad, oldroad)
    
    pt1.x = basePnt1(0)
    pt1.y = basePnt1(1)
    
    pt2.x = basePnt2(0)
    pt2.y = basePnt2(1)
    
    
    Set np3 = getNearPoint2(ptIntersection, pt2)
    
    'addDonut5 np3.x, np3.y
    
    pr1.init newroad
    Set line1 = pr1.getLine(pt1)
    
    pr2.init oldroad
    Set line2 = pr2.getLine(pt2)
    
    Set result = getShortestLine(line2, pt2, np3)
    
    
    
    'line1.Handle
    'line2.Handle
    
    'Application.ActiveDocument.SendCommand "_chamfer" & vbCr & "D" & vbCr & "1" & vbCr & "1" & vbCr & "(HandEnt """ & line1.Handle & """)" & vbCr & "(HandEnt """ & line2.Handle & """)" & vbCr
    Application.ActiveDocument.SendCommand "_fillet" & vbCr & "T" & vbCr & "N" & vbCr & "r" & vbCr & "3" & vbCr & "(HandEnt """ & line1.Handle & """)" & vbCr & "(HandEnt """ & result.Handle & """)" & vbCr
    
    
    If line2.Handle <> result.Handle Then
        result.Delete
    End If
    
    line1.Delete
    line2.Delete
    
    Dim ssetObj As AcadSelectionSet
    Set ssetObj = ThisDrawing.ActiveSelectionSet
    
    
    
    ssetObj.Select acSelectionSetLast
    
     For Each ent In ssetObj
        Set retArc = ent
     Next
    
    
    
    
    
    
    Debug.Print " "


End Sub


Function getShortestLine(line As AcadLine, ct As clsPoint, pt As clsPoint) As AcadLine

    Dim newline As AcadLine
    
    Dim sp() As Double
    Dim ep() As Double
        
    L12 = line.Length
    
    sp = line.startPoint
    ep = line.endPoint
    
    L13 = getDistance(sp(0), sp(1), pt.x, pt.y)
    L23 = getDistance(ep(0), ep(1), pt.x, pt.y)
    
    L1C = getDistance(sp(0), sp(1), ct.x, ct.y)
    L2C = getDistance(ep(0), ep(1), ct.x, ct.y)
    
    
    If L12 > L13 And L12 > L23 Then
        
        If L1C < L13 Then
            'L13
            sp(0) = sp(0)
            sp(1) = sp(1)
            ep(0) = pt.x
            ep(1) = pt.y
            
        Else        ' L1C < L23
            'L23
            sp(0) = ep(0)
            sp(1) = ep(1)
            ep(0) = pt.x
            ep(1) = pt.y
            
        End If
        
        Set newline = ThisDrawing.ModelSpace.addLine(sp, ep)
        
        'line.Delete
    ElseIf L12 > L23 And L12 > L13 Then
        
       'L13
        sp(0) = sp(0)
        sp(1) = sp(1)
        ep(0) = pt.x
        ep(1) = pt.y
        
        Set newline = ThisDrawing.ModelSpace.addLine(sp, ep)
        'line.Delete
   
    Else
    
        Set newline = line
    
    End If
    
    
    Debug.Print ""
    
    Set getShortestLine = newline


End Function


Sub test55()

    Dim plroad() As AcadEntity
    Dim ret As Variant
    Dim rad As Double
    
    plroad = selectObjs 'select two polylines(one is a source, the ohter is a target entity.
     
    
    'ret = addCrossVertext(plroad(0), plroad(1))
    
    
    
    
    
    
    Set tmp1 = plroad(0)
    Set tmp2 = plroad(1)
    
    
    tmp3 = tmp1.Handle
    tmp4 = tmp2.Handle
    
    Application.ActiveDocument.SendCommand "_chamfer" & vbCr & "D" & vbCr & "1" & vbCr & "1" & vbCr & "(HandEnt """ & tmp1.Handle & """)" & vbCr & "(HandEnt """ & tmp2.Handle & """)" & vbCr
    'Application.ActiveDocument.SendCommand "_fillet" & vbCr & "T" & vbCr & "N" & vbCr & "r" & vbCr & "3" & vbCr & "(HandEnt """ & tmp1.Handle & """)" & vbCr & "(HandEnt """ & tmp2.Handle & """)" & vbCr
    
    
    
    
    Debug.Print " "
    

End Sub


Sub LineFillet()



ThisDrawing.SendCommand "_fillet" & vbCr & "T" & vbCr & "N" & vbCr & "r" & vbCr & "3" & vbCr
ThisDrawing.SendCommand "(HandEnt """ & r.Handle & """)" & vbCr
ThisDrawing.SendCommand "(HandEnt """ & plineObj.Handle & """)" & vbCr


Debug.Print "done"
   'ThisDrawing.SendCommand "fillet T N r 3"
   'ThisDrawing.ModelSpace.ad
           
End Sub

 
Public Sub testFillet()

    Dim doc As Document

    doc = Application.DocumentManager.MdiActiveDocument

    Dim ed As Editor
    ed = doc.Editor

    Dim Db As Database
    Db = doc.Database



    Dim pEntRes1 As PromptEntityResult
    pEntRes1 = ed.getEntity("Select first line to Fillet")

    If pEntRes1.Status <> PromptStatus.OK Then

        Return

    End If



    Dim obj1 As String
    obj1 = pEntRes1.ObjectID.ObjectClass.name



    Dim pEntRes2 As PromptEntityResult
    pEntRes2 = ed.getEntity("Select second line to Fillet")

    If pEntRes2.Status <> PromptStatus.OK Then

        Return

    End If



    obj1 = pEntRes2.ObjectID.ObjectClass.name


    Dim strHandle1 As String
    
    strHandle1 = pEntRes1.ObjectID.Handle.ToString()

    Dim strEntName1 As String
    strEntName1 = "(handent """ & strHandle1 & """)"



    Dim strHandle2 As String
    strHandle2 = pEntRes2.ObjectID.Handle.ToString()

    Dim strEntName2 As String
    strEntName2 = "(handent """ & strHandle2 & """)"

    Dim strCommand As String
    strCommand = "_fillet" + vbCr + "r" + vbCr + "0.495" + vbCr + strEntName1 + vbCr + strEntName2 + vbCr

    doc.AcadDocument.SendCommand (strCommand)



End Sub

Public Sub prob2()
    Dim ssetObj As AcadSelectionSet
    Set ssetObj = ThisDrawing.PickfirstSelectionSet
    ssetObj.Clear
    
    ' asks for an entity and put it into the sset
    Dim returnObj As AcadObject
    Dim basePnt As Variant
    ThisDrawing.Utility.getEntity returnObj, basePnt, "Select object: "
    Dim objlist(0) As AcadObject
    Set objlist(0) = returnObj
    ' difference from previous version
    Dim objlistvar As Variant
    objlistvar = objlist
    ssetObj.AddItems objlistvar
    ' then delete the sset to make it the "previous selection set"
    ssetObj.Delete
    
    Set ssetObj = ThisDrawing.PickfirstSelectionSet ' or whatever ss
    ssetObj.Select acSelectionSetPrevious
End Sub



Public Sub FilletVertex(ByVal LWPline As AcadLWPolyline, _
    ByVal VertexNumber As Integer, ByVal Radius As Double, _
    Optional AsChamfer As Boolean)

    On Error Resume Next: DoEvents
    
    Dim IsConcave As Boolean
    Dim AngleToVertex As Double
    Dim AngleFromVertex As Double
    Dim AngleIncluded As Double
    Dim ptList As Variant
    Dim LastVertex As Integer
    Dim PrevVertex As Integer
    Dim NextVertex As Integer
    Dim Bulge As Double
    Dim pt1 As Variant
    Dim pt2 As Variant
    Dim pt2a As Variant
    Dim pt2b As Variant
    Dim pt3 As Variant
    Dim VertexA(1) As Double
    Dim VertexB(1) As Double
    Dim chamfer As Double
    Dim Util As AcadUtility
    
    Set Util = ThisDrawing.Utility

    If Radius = 0 Then
        Exit Sub
    ElseIf Radius < 0 Then
        Radius = Abs(Radius)
        IsConcave = True
    End If
    
    With LWPline
        ptList = .Coordinates
        
        LastVertex = (UBound(ptList) - 1) / 2
        If VertexNumber > LastVertex Then VertexNumber = 0
        NextVertex = VertexNumber + 1
        PrevVertex = VertexNumber - 1
        If NextVertex > LastVertex Then NextVertex = 0
        If PrevVertex < 0 Then PrevVertex = LastVertex
        
        If NextVertex = PrevVertex Then Exit Sub
        
        pt1 = .Coordinate(PrevVertex)
        pt2 = .Coordinate(VertexNumber)
        pt3 = .Coordinate(NextVertex)
    
        ReDim Preserve pt1(2): pt1(2) = 0
        ReDim Preserve pt2(2): pt2(2) = 0
        ReDim Preserve pt3(2): pt3(2) = 0
        
        AngleToVertex = Util.AngleFromXAxis(pt2, pt1)
        AngleFromVertex = Util.AngleFromXAxis(pt2, pt3)
        
        AngleIncluded = (AngleToVertex - AngleFromVertex)
        If AngleIncluded > PI Then
            AngleIncluded = AngleIncluded - (2 * PI)
        ElseIf AngleIncluded < -PI Then
            AngleIncluded = AngleIncluded + (2 * PI)
        End If
        

            
        If IsConcave Then
            chamfer = Radius
        Else
            chamfer = Radius * _
                Tan((PI - (Abs(AngleIncluded))) / 2)
        End If
        
        pt2b = Util.PolarPoint(pt2, _
            AngleFromVertex, chamfer)
        VertexB(0) = pt2b(0): VertexB(1) = pt2b(1)
        .Coordinate(VertexNumber) = VertexB
        
        pt2a = Util.PolarPoint(pt2, _
            AngleToVertex, chamfer)
        VertexA(0) = pt2a(0): VertexA(1) = pt2a(1)
        .addVertex VertexNumber, VertexA
        
        If Not AsChamfer Then

            If IsConcave Then
                Bulge = Tan((AngleIncluded) / -4#)
            Else
                Bulge = Tan((IIf(AngleIncluded > 0, PI, -PI) - AngleIncluded) / 4#)
            End If
            
             .SetBulge VertexNumber, Bulge

         End If
    End With
End Sub


Sub simplifyObject()    ' ¿À·ù

    Dim Userselection As AcadSelectionSet

    ThisDrawing.SelectionSets("CurrentSelection").Delete
    Set UsersSelection = ThisDrawing.SelectionSets.add("CurrentSelection")
        
    UsersSelection.SelectOnScreen
         
    
    
    tmp3 = UsersSelection.Handle

    ThisDrawing.SendCommand "(setq clean_id (tpm_cleanalloc))"
    ThisDrawing.SendCommand "(setq clean_var_id (tpm_varalloc))"
    ThisDrawing.SendCommand "(setq action_var_id (tpm_varalloc))"
    
    ThisDrawing.SendCommand "(tpm_varset clean_var_id ""CLEAN_TOL"" 2)"
    ThisDrawing.SendCommand "(tpm_varset clean_var_id ""INCLUDEOBJS_AUTOSELECT"" 0)"
    
    ThisDrawing.SendCommand "(tpm_cleanactionlistins clean_var_id 0 128 clean_var_id)"
    ThisDrawing.SendCommand "(setq DrwCleanresult (tpm_cleaninit clean_id clean_var_id (HandEnt """ & tmp1.Handle & """)))"
    ThisDrawing.SendCommand "(setq DrwCleanresult (tpm_cleanstart clean_id))"
    ThisDrawing.SendCommand "(tpm_cleanend clean_id)"
    
   

End Sub



Sub test100()

    ThisDrawing.StartUndoMark
    
    
    
    


End Sub



Sub test101()

    ThisDrawing.EndUndoMark
    
    
    ThisDrawing.SendCommand "Undo" & vbCr & "Back" & vbCr & "Y" & vbCr
    
    


End Sub


    
Sub test102()

    Dim plroad() As AcadEntity
    Dim ent As AcadEntity
    plroad = selectObjs
    
    
    For i = 0 To Bound(plroad)
        ent = plorad(i)
        showLayer ent.layer
    Next


End Sub

  
Sub test103()

    Dim plroad() As AcadEntity
    Dim ent As AcadEntity
    plroad = selectObjs
    
    
    For i = 0 To Bound(plroad)
        ent = plorad(i)
        hideLayer ent.layer
    Next


End Sub

Sub test104()

    Dim tmpfillet As AcadLWPolyline
    Dim tmpoldroad As AcadLWPolyline
    Dim intersec As Variant
    
    
    Dim plroad() As AcadEntity
    Dim ent As AcadEntity
    plroad = selectObjs
    
    intersec = searchCrossVertext(plroad(0), plroad(1))
    
    For i = 0 To UBound(intersec) Step 3
    
        addDonutA intersec(i), intersec(i + 1)
    Next
    
    
    
    Debug.Print ""
    

End Sub


Sub test105()

    Dim plroad As AcadEntity
    Dim ent As AcadEntity
    Dim ret As AcadEntity
    
    Dim ddd() As Double
    Set plroad = selectEntity

    Debug.Print ""
    
    
    Set ret = trimPolyline(plroad)
    
    plroad.Delete
    
    'ddd = ret.Coordinates
    '
    'For i = 0 To UBound(ddd) Step 2
    '    addDonutA ddd(i), ddd(i + 1)
    'Next
    
    
    '1146072.1709
End Sub


