VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFillet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private orgLayer As String
Private oldroad As AcadEntity
Private newroad As AcadEntity
Private fillet As AcadEntity

Private ptIntersection As Variant


Private np1 As New clsPoint
Private np2 As New clsPoint
Private np3 As New clsPoint


Private newroadsplit() As Double
Private oldroadsplit() As Double
Private filletsplit() As Double



Public Function load()

    Dim retArc As AcadEntity
    Dim tmp As Variant
    
    
    Dim tmpnewroad As AcadEntity
    
    Dim splitter As New clsSplitter
    Dim pr As New clsPolygonReader

    
    Dim merge As New ClsMerger
    Dim resultd() As Double
    Dim test As Boolean
    
    Dim pt1 As New clsPoint
    Dim pt3 As New clsPoint
    
    
    Dim pr1 As New clsPolygonReader
    Dim pr2 As New clsPolygonReader
    
    Dim line1 As AcadLine
    Dim line2 As AcadLine
    Dim resultpl As AcadLine
    Dim ssetObj As AcadSelectionSet
    Dim result As AcadEntity

  
    Dim selec1 As New ClsSelEntity
    Dim selec2 As New ClsSelEntity
    Dim undoPoint As AcadEntity
    
    Dim seler As New ClsSelReal
    Dim frad As Double
    
    Dim selecs As New clsSelSets
    
    Dim eidx As Integer
    Dim freezelayer As AcadLayer
    Dim ok As Boolean
    
    selec1.selectPolyline ("#Get Polyline 1: ")
    Set tmpnewroad = selec1.getEntity
    Set pt1 = selec1.getPoint
    
    
    'newFreezeLayer "zzz_Freeze"
    Set undoPoint = ThisDrawing.ModelSpace.AddText("UU-U", selec1.getMovePoindD(5), 2)
    'undoPoint.layer = "zzz_Freeze"
    
    
    ok = selec2.selectPolyline("#Get Polyline 2: ")
    
    If ok = False Then
        undoPoint.Delete
        End
    
    End If
    
    
    Set oldroad = selec2.getEntity
    Set pt3 = selec2.getPoint
    
    Debug.Print "!!! click " & pt3.x & ", " & pt3.y
    
    'pt3.x = 3914.05231561216
    'pt3.y = 1998.90660493234
    
    
    Set newroad = trimPolyline(tmpnewroad)
    newroad.layer = tmpnewroad.layer
    tmpnewroad.Delete
    
    ptIntersection = searchCrossVertext(newroad, oldroad)
    
    Set np3 = getNearPoint2(ptIntersection, pt3)
    
    'addDonut5 np3.x, np3.y
    
    
    pr1.init newroad
    Set line1 = pr1.getLine(pt1)
    
     
    pr2.init oldroad
    Set line2 = pr2.getLine(pt3)
    
     
    
    Set resultpl = getShortestLine(line2, pt3, np3)
    
    
    frad = seler.selectFilletRadius
    
    
    Application.ActiveDocument.SendCommand "_fillet" & vbCr & "T" & vbCr & "N" & vbCr & "r" & vbCr & frad & vbCr & "(HandEnt """ & line1.Handle & """)" & vbCr & "(HandEnt """ & resultpl.Handle & """)" & vbCr
    
    
    Set retArc = selecs.selectArc
    
    'Set retArc = selectArcObj
    
    'Set ssetObj = ThisDrawing.ActiveSelectionSet
    '
    'ssetObj.Select acSelectionSetLast
    '
    'For Each ent In ssetObj
    '   Set retArc = ent
    'Next
    
    If line2.Handle <> resultpl.Handle Then
        resultpl.Delete
    End If
    
    line1.Delete
    line2.Delete
    
    
    
    Set fillet = arc2line(retArc)  '아크를 라인으로 변경
    
    ThisDrawing.SendCommand "sob" & vbCr
    
    filletsplit = fillet.Coordinates
    
    retArc.Delete
    
    
    Debug.Print "fillet created"
    
    
    ptIntersection = searchCrossVertext(newroad, oldroad)
    
    
    'For i = 0 To UBound(ptIntersection) Step 3
    '    addDonutA ptIntersection(i), ptIntersection(i + 1)          '!!!!!!!!!!!!!!!!!!!!!!
    'Next i
    
    newp1 = findVertex2(newroad, fillet, pt1)   '아크와 도로(새도로)의 접점 추가
    newp2 = findVertex2(oldroad, fillet, pt3)   '아크와 도로(구도로)의 접점 추가
   
   
    np1.initXy newp1(0), newp1(1)
    np2.initXy newp2(0), newp2(1)
    
    
    Debug.Print "np1: " & np1.x & ", " & np1.y
    Debug.Print "np2: " & np2.x & ", " & np2.y
    
    
    'If np1.x = 0 Or np1.y = 0 Or np2.x = 0 Or np2.y = 0 Then
    '    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
     
    'End If
    
    
    
    
    'addDonut5 np1.x, np1.y
    'addDonut5 np2.x, np2.y
   
    Set np3 = getFarPoint2Ex(newroad, ptIntersection, np2)     '기존도로 아크 접점에서 먼 지점을 찾는다
    
    'addDonut5 np3.x, np3.y
    
    addVertex oldroad, np3.x, np3.y     '기존도로에 접점추가
    addVertex oldroad, np2.x, np2.y
    addVertex newroad, np1.x, np1.y
    
    Debug.Print "cross point created"
    
    pr.init newroad
    pr.setIndexByPoint np1
    pr.checkDirection pt3
    
    Debug.Print "newroad direction processed"
    
    
    sidx = pr.setIndexByPoint(np1)
    eidx = pr.setEIndexByPoint(np3)
    
    newroadsplit = pr.getLongLine2()
  
 
    Debug.Print "newroad split created"
   
    pr.init oldroad
    pr.setIndexByPoint np2
    pr.checkDirection pt1
    
    Debug.Print "oldroad direction processed"
    
    
    sidx = pr.setIndexByPoint(np2)
    eidx = pr.setEIndexByPoint(np3)
    
    oldroadsplit = pr.getShortLine2()
    
    'oldroadsplit = pr.splitInnerLine(np2, np3)
    'addLine oldroadsplit

    
    Debug.Print "oldroad split created"
        
        
    checkLastPoint
    checkCrossPoint
    
        
    size1 = UBound(newroadsplit)
    size2 = UBound(filletsplit)
    size3 = UBound(oldroadsplit)
    
    merge.setSize size1 + size2 + size3 + 2
    
    
    'ThisDrawing.ModelSpace.AddLightWeightPolyline oldroadsplit
    
    
    merge.add newroadsplit
    merge.add filletsplit
    merge.add oldroadsplit
    
    
    Set result = addLine(merge.getMerged)
    'newroad.Delete
    fillet.Delete
    
    
    'Set temp = ThisDrawing.ActiveLayer
    
    orgLayer = newroad.layer
    result.layer = getAddLayerName(orgLayer)
    newroad.layer = getDelLayerName(orgLayer)
    
    If orgLayer Like LYR_ADD & "*" Then
        newroad.Delete
    End If
    
    
    
    undoPoint.Delete
    
    
End Function


Private Function getArcBy3P(p1() As Double, p2() As Double, p3() As Double) As AcadArc
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


Function checkLastPoint()

    Dim retvalue() As Double
    Dim lastx As Double
    Dim lasty As Double
    Dim size As Integer
    Dim dist1 As Double
    Dim dist2 As Double
    
    size = UBound(filletsplit)
    
    lastx = filletsplit(size - 1)
    lasty = filletsplit(size)
    
    'addDonutA lastx, lasty
    'addDonutA np2.x, np2.y

    dist1 = getDistance(np2.x, np2.y, lastx, lasty)
    
    lastx = filletsplit(0)
    lasty = filletsplit(1)
    
    dist2 = getDistance(np2.x, np2.y, lastx, lasty)
    
    
    If dist1 <> 0 And dist2 <> 0 Then   '겹치지 않을때
    
        If dist1 < 0.1 Then      'dist1 is minvalue
        
            filletsplit(size - 1) = np2.x
            filletsplit(size) = np2.y
        
          '  ReDim retvalue(size - 2) As Double
         '   For i = 0 To size - 2
        '        retvalue(i) = filletsplit(i + 2)
       '
      '      Next i
        
            'filletsplit = retvalue
         
        ElseIf dist2 < 0.1 Then     ' dist2 is minvalue
            
            filletsplit(0) = np2.x
            filletsplit(1) = np2.y
        
    '        ReDim retvalue(size - 2) As Double
     '       For i = 2 To size
      '          retvalue(i - 2) = filletsplit(i)
       '
        '    Next i
        
         '   filletsplit = retvalue
        
        
        End If
    End If
     


End Function
 
 
Function checkCrossPoint()


    Dim tmpfillet As AcadLWPolyline
    Dim tmpoldroad As AcadLWPolyline
    Dim intersec As Variant
    
    Dim selidx As Integer
    Dim mindist As Double
    
    Dim plreader As New ClsPolylineReader
    Dim nx As Double
    Dim ny As Double
    
    selidx = -1
    mindist = 99999
    
    Set tmpfillet = ThisDrawing.ModelSpace.AddLightWeightPolyline(filletsplit)
    Set tmpoldroad = ThisDrawing.ModelSpace.AddLightWeightPolyline(oldroadsplit)
    
    intersec = searchCrossVertext(tmpfillet, oldroad)
    
    For i = 0 To UBound(intersec) Step 3
    
        dist = getDistance(np3.x, np3.y, intersec(i), intersec(i + 1))
        
        If dist < mindist Then
        
            mindist = dist
            selidx = i
        End If
        
    Next i
    
    
    If selidx = -1 Then
        Return
    End If
    
    
    nx = intersec(selidx)
    ny = intersec(selidx + 1)
    
    'addDonutA nx, ny
    
    
    
    
    addVertex tmpfillet, nx, ny
    addVertex tmpoldroad, nx, ny
    
    plreader.init tmpfillet
    filletsplit = plreader.getLongDDLine(nx, ny)
    'ThisDrawing.ModelSpace.AddLightWeightPolyline filletsplit
    
    plreader.init tmpoldroad
    oldroadsplit = plreader.getLongDDLine(nx, ny)
    'ThisDrawing.ModelSpace.AddLightWeightPolyline oldroadsplit
    
    addVertex oldroad, nx, ny
    
    
    tmpfillet.Delete
    tmpoldroad.Delete
    
    
    
    
    Debug.Print ""
    
    
    
End Function
