VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFillet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private orgLayer As String
Private oldroad As AcadEntity
Private newroad As AcadEntity
Private fillet As AcadEntity

Private ptIntersection As Variant


Public Function load()

    Dim pt1 As New clsPoint
    Dim p2() As Double
    Dim pt3 As New clsPoint
    Dim retArc As AcadEntity
    Dim tmp As Variant
    
    
    Dim np1 As New clsPoint
    Dim np2 As New clsPoint
    Dim np3 As New clsPoint
    
    Dim splitter As New clsSplitter
    Dim pr As New clsPolygonReader
    
    Dim newroadsplit() As Double
    Dim oldroadsplit() As Double
    Dim filletsplit() As Double
    
    Dim merge As New ClsMerger
    Dim resultd() As Double
    Dim result As AcadEntity
    Dim test As Boolean
    
    
    
    Dim selec1 As New ClsSelEntity
    Dim selec2 As New ClsSelEntity
    
    
    Dim seler As New ClsSelReal
    Dim frad As Double
    
    Dim selecs As New clsSelSets
    
    selec1.selectPolyline ("#Get Polyline 1: ")
    Set newroad = selec1.getEntity
    Set pt1 = selec1.getPoint
    
    ThisDrawing.SetVariable "osmode", 0
    p2 = ThisDrawing.Utility.getPoint(, "#Enter a point2: ")
    
    selec2.selectPolyline ("#Get Polyline 3: ")
    Set oldroad = selec2.getEntity
    Set pt3 = selec2.getPoint
    
    
    Debug.Print "pt1: " & pt1.x & ", " & pt1.y
    Debug.Print "p2: " & p2(0) & ", " & p2(1)
    Debug.Print "pt3: " & pt3.x & ", " & pt3.y
   
    
    Set retArc = getArcBy3P(pt1.getPoints(), p2, pt3.getPoints()) '세점을 지나는 아크를 생성
    Set fillet = arc2line(retArc)  '아크를 라인으로 변경
    filletsplit = fillet.Coordinates
    
    retArc.Delete
    
    
    Debug.Print "fillet created"
    
    
    ptIntersection = searchCrossVertext(newroad, oldroad)
    
        
    For i = 0 To UBound(ptIntersection) Step 3
        addDonutY ptIntersection(i), ptIntersection(i + 1)
        
    Next i
    
    newp1 = findVertex2(newroad, fillet, pt1)   '아크와 도로(새도로)의 접점 추가
    newp2 = findVertex2(oldroad, fillet, pt3)   '아크와 도로(구도로)의 접점 추가
   
   
    np1.initXy newp1(0), newp1(1)
    np2.initXy newp2(0), newp2(1)
    
    
    Debug.Print "np1: " & np1.x & ", " & np1.y
    Debug.Print "np2: " & np2.x & ", " & np2.y
    
    addDonutB np1.x, np1.y
    addDonutB np2.x, np2.y
    
   
    Set np3 = getFarPoint2(ptIntersection, np2)     '기존도로 아크 접점에서 먼 지점을 찾는다
    
    
    addVertex oldroad, np3.x, np3.y     '기존도로에 접점추가
    
    addDonutG np3.x, np3.y
   
    Debug.Print "cross point created"
    
    pr.init newroad
    pr.setIndexByPoint np1
    pr.checkDirection pt3
    
    'addDonut5 np3.x, np3.y
    'addDonut5 np1.x, np1.y
    Debug.Print "newroad direction processed"
    
    newroadsplit = pr.splitInnerLine(np3, np1)
    'addLine newroadsplit
    
    Debug.Print "newroad split created"
   
    pr.init oldroad
    pr.setIndexByPoint np2
    pr.checkDirection pt1
    
    Debug.Print "oldroad direction processed"
    
    addDonutA pt1.x, pt1.y
    addDonutA np2.x, np2.y
    addDonutA np3.x, np3.y
    
    oldroadsplit = pr.splitInnerLine(np2, np3)
    'addLine oldroadsplit
    
    
    
    Debug.Print "oldroad split created"
        
    size1 = UBound(newroadsplit)
    size2 = UBound(filletsplit)
    size3 = UBound(oldroadsplit)
    
    merge.setSize size1 + size2 + size3 + 2
    
    
    merge.add newroadsplit
    merge.add filletsplit
    merge.add oldroadsplit
    
    
    Set result = addLine(merge.getMerged)
    'newroad.Delete
    fillet.Delete
    
    ThisDrawing.SendCommand "sob" & vbCr
    
    
    'Set temp = ThisDrawing.ActiveLayer
    
    orgLayer = newroad.layer
    result.layer = getAddLayerName(orgLayer)
    newroad.layer = getDelLayerName(orgLayer)
    
    If orgLayer Like LYR_ADD & "*" Then
        newroad.Delete
    End If
    
    
    
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

 
 
