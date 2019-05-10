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
    
Private np1 As New clsPoint
Private np2 As New clsPoint
Private np3 As New clsPoint


Private newroadsplit() As Double
Private oldroadsplit() As Double
Private filletsplit() As Double

Private tmpf As AcadLWPolyline


Private isEdited As Boolean






Public Function load()

    Dim pt1 As New clsPoint
    Dim p2() As Double
    Dim pt3 As New clsPoint
    Dim retArc As AcadEntity
    Dim tmp As Variant
    
    Dim splitter As New clsSplitter
    Dim pr As New clsPolygonReader
    
    Dim merge As New ClsMerger
    Dim resultd() As Double
    Dim result As AcadEntity
    Dim test As Boolean
    
    Dim selec1 As New ClsSelEntity
    Dim selec2 As New ClsSelEntity
    Dim undoPoint As AcadEntity
    
    Dim seler As New ClsSelReal
    Dim frad As Double
    
    Dim selecs As New clsSelSets
    Dim okcnt As Integer
    
    Dim ok As Boolean
    
    On Error GoTo EndOfFunction
    Debug.Print "==================START ================="
    
    isEdited = False
    okcnt = 0
    
    ok = selec1.selectPolyline("#Get Polyline 1: ")
    
    If ok = False Then
        ThisDrawing.Utility.prompt "선택에 오류가 발생했습니다."
        GoTo EndOfFunction
    
    End If
    
    Set newroad = selec1.getEntity
    Set pt1 = selec1.getPoint
    
    Set undoPoint = ThisDrawing.ModelSpace.AddText("UU-U", selec1.getMovePoindD(5), 2)
    
    ThisDrawing.SetVariable "osmode", 0
    p2 = ThisDrawing.Utility.getPoint(, "#Enter a point2: ")
    
    ok = selec2.selectPolyline("#Get Polyline 3: ")
    
    If ok = False Then
        ThisDrawing.Utility.prompt "선택에 오류가 발생했습니다."
        GoTo EndOfFunction
    
    End If
    
    Set oldroad = selec2.getEntity
    Set pt3 = selec2.getPoint
    
    
    isEdited = hasProperties(newroad)
    
    
    
    Debug.Print "pt1.x = " & pt1.x
    Debug.Print "pt1.y = " & pt1.y
    Debug.Print "p2(0) = " & p2(0)
    Debug.Print "p2(1) = " & p2(1)
    Debug.Print "pt3.x = " & pt3.x
    Debug.Print "pt3.y = " & pt3.y
   
   
    'pt1.x = 1111079.17911234
    'pt1.y = 1684396.52927265
    'p2(0) = 1111081.64487035
    'p2(1) = 1684391.02156739
    'pt3.x = 1111089.6705744
    'pt3.y = 1684388.55760551
   
     
    ptIntersection = searchCrossVertext(newroad, oldroad)
        
    For i = 0 To UBound(ptIntersection) Step 3
        'addDonutA ptIntersection(i), ptIntersection(i + 1)          '!!!!!!!!!!!!!!!!!!!!!!
        
        ok = isContains(newroad, CDbl(ptIntersection(i)), CDbl(ptIntersection(i + 1)))
        If ok = True Then
            okcnt = okcnt + 1
        End If
    Next i
    
    
    If okcnt < 2 Then
        ThisDrawing.Utility.prompt "인접부분 확인바랍니다 - 접점: " & okcnt
        GoTo EndOfFunction
    
    End If
   
    
    Set retArc = getArcBy3P(pt1.getPoints(), p2, pt3.getPoints()) '세점을 지나는 아크를 생성
    Set fillet = arc2line(retArc)  '아크를 라인으로 변경
    
    ThisDrawing.SendCommand "sob" & vbCr
    
    filletsplit = fillet.Coordinates
    
    retArc.Delete
    
    Debug.Print "fillet created"
    
  
    
  
    
    
    newp1 = findVertex2(newroad, fillet, pt1)   '아크와 도로(새도로)의 접점 추가
    newp2 = findVertex2(oldroad, fillet, pt3)   '아크와 도로(구도로)의 접점 추가
   
   
    np1.initXy newp1(0), newp1(1)
    np2.initXy newp2(0), newp2(1)
    
    
    Debug.Print "np1: " & np1.x & ", " & np1.y
    Debug.Print "np2: " & np2.x & ", " & np2.y
    
    'addDonutB np1.x, np1.y
    'addDonutB np2.x, np2.y
    
   
    Set np3 = getFarPoint2Ex(newroad, ptIntersection, np1)     '기존도로 아크 접점에서 먼 지점을 찾는다
    
    ptIntersection = searchCrossVertext(newroad, fillet)
  
    If UBound(ptIntersection) > 0 Then
        Set np1 = getNearPoint2(ptIntersection, np3)
        addVertex fillet, np1.x, np1.y
        'addDonutA np1.x, np1.y
    End If
    
    ptIntersection = searchCrossVertext(oldroad, fillet)
  
    If UBound(ptIntersection) > 0 Then
        Set np2 = getNearPoint2(ptIntersection, np3)
        addVertex fillet, np2.x, np2.y
        'addDonutA np2.x, np2.y
    End If
    
    filletsplit = fillet.Coordinates
    
    
    addVertex oldroad, np3.x, np3.y     '기존도로에 접점추가
    addVertex oldroad, np2.x, np2.y
    addVertex newroad, np1.x, np1.y
    
    'addDonutA np1.x, np1.y
    'addDonutA np2.x, np2.y
    'addDonutA np3.x, np3.y
   
    Debug.Print "cross point created"
    
    pr.init newroad
    pr.setIndexByPoint np1
    pr.checkDirection pt3
    
    'addDonut5 np3.x, np3.y
    'addDonut5 np1.x, np1.y
    Debug.Print "newroad direction processed"
    
    sidx = pr.setIndexByPoint(np1)
    eidx = pr.setEIndexByPoint(np3)
    
    newroadsplit = pr.getLongLine2()
    
    Debug.Print "newroad split created"
   
    pr.init oldroad
    pr.setIndexByPoint np2
    pr.checkDirection pt1
    
    Debug.Print "oldroad direction processed"
    
    'addDonutA pt1.x, pt1.y
    'addDonutA np2.x, np2.y
    'addDonutA np3.x, np3.y
    
    sidx = pr.setIndexByPoint(np2)
    eidx = pr.setEIndexByPoint(np3)
    
    oldroadsplit = pr.getShortLine2()
    
    'oldroadsplit = pr.splitInnerLine(np2, np3)
    'addLine oldroadsplit
    
    
    
    Debug.Print "oldroad split created"
    
    checkTailPointEx np2
    checkLastPoint
    'oldroadsplit = checkCrossPoint3Ex(oldroad, oldroadsplit, np3)
    checkCrossPoint
    
    Debug.Print "oldroad split checked"
    
    checkTailPointEx np1
    checkLastPoint2
    'newroadsplit = checkCrossPoint4Ex(newroad, newroadsplit, np2)
    checkCrossPoint2
    
    Debug.Print "newroad split checked"
        
    size1 = UBound(newroadsplit)
    size2 = UBound(filletsplit)
    size3 = UBound(oldroadsplit)
    
    merge.setSize size1 + size2 + size3 + 2
    
    
    merge.add newroadsplit
    merge.add filletsplit
    merge.add oldroadsplit
    
    
    Set result = addLine(merge.getMerged)
    result.Closed = True
    'newroad.Delete
    fillet.Delete
    
    
    'Set temp = ThisDrawing.ActiveLayer
    
    orgLayer = newroad.layer
    result.layer = getAddLayerName(orgLayer)
    newroad.layer = getDelLayerName(orgLayer)
    
    If (orgLayer Like LYR_ADD & "*") Or isEdited = False Then
        newroad.Delete
    End If
    
EndOfFunction:
    If undoPoint Is Nothing = False Then
        undoPoint.Delete
    End If
    Debug.Print "DONE"
    
    
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


Function checkTailPointEx(pt As clsPoint)
    
    Dim idx As Integer
    Dim lastx As Double
    Dim lasty As Double
    Dim size As Integer
    Dim dist1 As Double
    Dim dist2 As Double
    Dim x As Double
    Dim y As Double
    Dim ret() As Double
    
    size = UBound(filletsplit)
    
    lastx = filletsplit(size - 1)
    lasty = filletsplit(size)

    dist1 = getDistance(pt.x, pt.y, lastx, lasty)
    
    lastx = filletsplit(0)
    lasty = filletsplit(1)
    
    dist2 = getDistance(pt.x, pt.y, lastx, lasty)
    
    
    'addDonutA np1.x, np1.y
    
    If dist1 < dist2 Then
        For i = size - 1 To 0 Step -2
            x = filletsplit(i)
            y = filletsplit(i + 1)
            
            If x = pt.x And y = pt.y Then
            
                Debug.Print ""
                Exit For
                
            End If
            idx = idx - 2
            
        Next i
    
    Else
        
        
        For i = 0 To size Step 2
            x = filletsplit(i)
            y = filletsplit(i + 1)
            
            'addDonut5 x, y
            
            'Debug.Print "*" & pt.x & ", " & pt.y
            'Debug.Print "+" & x & ", " & y
            
            b1 = (x = pt.x)
            b2 = (y = pt.y)
            
            If isEqualsDouble(x, y, pt.x, pt.y) Then
            
                Debug.Print ""
                Exit For
                
            End If
            idx = idx + 2
        
        Next i
        
        If idx < size Then
            ReDim ret(size - idx) As Double
            
            For i = idx To size Step 2
            
                ret(i - idx) = filletsplit(i)
                ret(i - idx + 1) = filletsplit(i + 1)
                Debug.Print ""
            Next i
            
            filletsplit = ret
        End If
    
    
    End If
  
    
End Function

Function checkTailPoint()
    
    Dim idx As Integer
    Dim lastx As Double
    Dim lasty As Double
    Dim size As Integer
    Dim dist1 As Double
    Dim dist2 As Double
    Dim x As Double
    Dim y As Double
    Dim ret() As Double
    
    size = UBound(filletsplit)
    
    lastx = filletsplit(size - 1)
    lasty = filletsplit(size)

    dist1 = getDistance(np1.x, np1.y, lastx, lasty)
    
    lastx = filletsplit(0)
    lasty = filletsplit(1)
    
    dist2 = getDistance(np1.x, np1.y, lastx, lasty)
    
    
    'addDonutA np1.x, np1.y
    
    If dist1 < dist2 Then
        For i = size - 1 To 0 Step -2
            x = filletsplit(i)
            y = filletsplit(i + 1)
            
            If x = np1.x And y = np1.y Then
            
                Debug.Print ""
                Exit For
                
            End If
            idx = idx - 2
            
        Next i
    
    Else
        
        
        For i = 0 To size Step 2
            x = filletsplit(i)
            y = filletsplit(i + 1)
            
            'addDonut5 x, y
            
            Debug.Print "*" & np1.x & ", " & np1.y
            Debug.Print "+" & x & ", " & y
            
            b1 = (x = np1.x)
            b2 = (y = np1.y)
            
            If isEqualsDouble(x, y, np1.x, np1.y) Then
            
                Debug.Print ""
                Exit For
                
            End If
            idx = idx + 2
        
        Next i
        
        If idx < size Then
            ReDim ret(size - idx) As Double
            
            For i = idx To size Step 2
            
                ret(i - idx) = filletsplit(i)
                ret(i - idx + 1) = filletsplit(i + 1)
                Debug.Print ""
            Next i
            
            filletsplit = ret
        End If
    
    
    End If
  
    
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
    
    'addDonut5 lastx, lasty
    'addDonutA np2.x, np2.y

    dist1 = getDistance(np2.x, np2.y, lastx, lasty)
    
    lastx = filletsplit(0)
    lasty = filletsplit(1)
    
    dist2 = getDistance(np2.x, np2.y, lastx, lasty)
    
    
    If dist1 <> 0 And dist2 <> 0 Then   '겹치지 않을때
    
        If dist1 < dist2 Then
            filletsplit(size - 1) = np2.x
            filletsplit(size) = np2.y
        
        Else
            filletsplit(0) = np2.x
            filletsplit(1) = np2.y
        
        End If
  
    End If
  
End Function



Function checkLastPoint2()  '

    Dim retvalue() As Double
    Dim firstx As Double
    Dim firsty As Double
    Dim size As Integer
    Dim dist1 As Double
    Dim dist2 As Double
    
    size = UBound(filletsplit)
    
    firstx = filletsplit(size - 1)
    firsty = filletsplit(size)
    
    'addDonut5 firstx, firsty
    'addDonutA np1.x, np1.y

    dist1 = getDistance(np1.x, np1.y, firstx, firsty)
    
    firstx = filletsplit(0)
    firsty = filletsplit(1)
    
    dist2 = getDistance(np1.x, np1.y, firstx, firsty)
    
    
    If dist1 <> 0 And dist2 <> 0 Then   '겹치지 않을때
    
        If dist1 < dist2 Then
            filletsplit(size - 1) = np1.x
            filletsplit(size) = np1.y
        
        Else
            filletsplit(0) = np1.x
            filletsplit(1) = np1.y
        
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
    
    Dim x, y As Double
    Dim nx As Double
    Dim ny As Double
    
    Dim b1, b2 As Boolean
    
    selidx = -1
    mindist = 99999
    
    Set tmpfillet = ThisDrawing.ModelSpace.AddLightWeightPolyline(filletsplit)
    Set tmpoldroad = ThisDrawing.ModelSpace.AddLightWeightPolyline(oldroadsplit)
    
    intersec = searchCrossVertext(tmpfillet, oldroad)
    
    'addDonutA np3.x, np3.y
    
    For i = 0 To UBound(intersec) Step 3
    
        x = intersec(i)
        y = intersec(i + 1)
    
        b1 = isContains(tmpfillet, CDbl(x), CDbl(y))
        b2 = isContains(tmpoldroad, CDbl(x), CDbl(y))
        
        
        'addDonut5 intersec(i), intersec(i + 1)
        
        If b1 = False Or b2 = False Then
            
        Else
        
            dist = getDistance(np3.x, np3.y, intersec(i), intersec(i + 1))
            
            
            
            If dist < mindist Then
            
                mindist = dist
                selidx = i
            End If
        
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
    'Set tmpf = ThisDrawing.ModelSpace.AddLightWeightPolyline(filletsplit)
    'tmpf.color = acBlue
    
    plreader.init tmpoldroad
    oldroadsplit = plreader.getLongDDLine(nx, ny)
    'Set tmpf = ThisDrawing.ModelSpace.AddLightWeightPolyline oldroadsplit
    
    addVertex oldroad, nx, ny
    
    
    tmpfillet.Delete
    tmpoldroad.Delete
    
    
    Debug.Print ""
    
    
    
End Function

Function checkCrossPointEx()


    Dim tmpfillet As AcadLWPolyline
    Dim tmpoldroad As AcadLWPolyline
    Dim intersec As Variant
    
    Dim selidx As Integer
    Dim mindist As Double
    
    Dim plreader As New ClsPolylineReader
    
    Dim x, y As Double
    Dim nx As Double
    Dim ny As Double
    
    Dim b1, b2 As Boolean
    
    selidx = -1
    mindist = 99999
    
    Set tmpfillet = ThisDrawing.ModelSpace.AddLightWeightPolyline(filletsplit)
    Set tmpoldroad = ThisDrawing.ModelSpace.AddLightWeightPolyline(oldroadsplit)
    
    intersec = searchCrossVertext(tmpfillet, oldroad)
    
    For i = 0 To UBound(intersec) Step 3
    
        x = intersec(i)
        y = intersec(i + 1)
    
        'b1 = isContains(tmpfillet, CDbl(x), CDbl(y))
        'b2 = isContains(tmpoldroad, CDbl(x), CDbl(y))
        
        
        'addDonut5 intersec(i), intersec(i + 1)
        
                
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
    
    'b1 = isContains(tmpfillet, CDbl(x), CDbl(y))
    'b2 = isContains(tmpoldroad, CDbl(x), CDbl(y))
    
    
    np2.x = nx
    np2.y = ny
    
    addVertex tmpfillet, nx, ny
    addVertexT tmpoldroad, nx, ny
   
    'Set tmpf = sublwpolyline(nx, ny, tmpoldroad)
   
    plreader.init tmpfillet
    filletsplit = plreader.getLongDDLine(nx, ny)
    
    'Set tmpf = ThisDrawing.ModelSpace.AddLightWeightPolyline(filletsplit)
    'tmpf.color = acBlue
    
    plreader.init tmpoldroad
    oldroadsplit = plreader.getLongDDLine(nx, ny)
    'Set tmpf = ThisDrawing.ModelSpace.AddLightWeightPolyline(oldroadsplit)
    'tmpf.color = acRed
    
    tmpfillet.Delete
    tmpoldroad.Delete
    
    
    Debug.Print ""
    
    
    
End Function



Function checkCrossPoint2()


    Dim tmpfillet As AcadLWPolyline
    Dim tmpnewroad As AcadLWPolyline
    Dim intersec As Variant
    
    Dim selidx As Integer
    Dim mindist As Double
    
    Dim plreader As New ClsPolylineReader
    
    Dim x, y As Double
    Dim nx As Double
    Dim ny As Double
    
    Dim b1, b2 As Boolean
    
    selidx = -1
    mindist = 99999
    
    Set tmpfillet = ThisDrawing.ModelSpace.AddLightWeightPolyline(filletsplit)
    Set tmpnewroad = ThisDrawing.ModelSpace.AddLightWeightPolyline(newroadsplit)
    
    intersec = searchCrossVertext(tmpfillet, newroad)
    
    'addDonut5 np2.x, np2.y
    
    For i = 0 To UBound(intersec) Step 3
    
        x = intersec(i)
        y = intersec(i + 1)
    
        b1 = isContains(tmpfillet, CDbl(x), CDbl(y))
        b2 = isContains(tmpnewroad, CDbl(x), CDbl(y))
        
        If b1 = False Or b2 = False Then
            
        Else
        
            dist = getDistance(np2.x, np2.y, intersec(i), intersec(i + 1))
            
            'addDonutA intersec(i), intersec(i + 1)
            
            If dist < mindist Then
            
                mindist = dist
                selidx = i
            End If
        
        End If
    Next i
    
    
    If selidx = -1 Then
        Return
    End If
    
    
    nx = intersec(selidx)
    ny = intersec(selidx + 1)
    
    'addDonutA nx, ny
    
    
    
    
    addVertex tmpfillet, nx, ny
    addVertex tmpnewroad, nx, ny
    
    plreader.init tmpfillet
    filletsplit = plreader.getLongDDLine(nx, ny)
    'ThisDrawing.ModelSpace.AddLightWeightPolyline filletsplit
    
    plreader.init tmpnewroad
    newroadsplit = plreader.getLongDDLine(nx, ny)
    'ThisDrawing.ModelSpace.AddLightWeightPolyline oldroadsplit
    
    addVertex newroad, nx, ny
    
    
    tmpfillet.Delete
    tmpnewroad.Delete
    
    
    Debug.Print ""
    
    
    
End Function



Function checkCrossPoint3Ex(road As AcadEntity, roadsplit() As Double, np As clsPoint) As Double()


    Dim tmpfillet As AcadLWPolyline
    Dim tmproad As AcadLWPolyline
    Dim intersec As Variant
    
    Dim selidx As Integer
    Dim mindist As Double
    
    Dim plreader As New ClsPolylineReader
    
    Dim x, y As Double
    Dim nx As Double
    Dim ny As Double
    
    Dim b1, b2 As Boolean
    
    selidx = -1
    mindist = 99999
    
    Set tmpfillet = ThisDrawing.ModelSpace.AddLightWeightPolyline(filletsplit)
    Set tmproad = ThisDrawing.ModelSpace.AddLightWeightPolyline(roadsplit)
    
    intersec = searchCrossVertext(tmpfillet, road)
    
    'addDonut5 np.x, np.y
    
    For i = 0 To UBound(intersec) Step 3
    
        x = intersec(i)
        y = intersec(i + 1)
    
        'b1 = isContains(tmpfillet, CDbl(x), CDbl(y))
        'b2 = isContains(tmproad, CDbl(x), CDbl(y))
        
        
        dist = getDistance(np.x, np.y, intersec(i), intersec(i + 1))
            
        'addDonutA intersec(i), intersec(i + 1)
            
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
    
    np2.x = nx
    np2.y = ny
    
    
    
    addVertex tmpfillet, nx, ny
    addVertexT2 tmproad, nx, ny
    
    plreader.init tmpfillet
    filletsplit = plreader.getLongDDLine(nx, ny)
    'ThisDrawing.ModelSpace.AddLightWeightPolyline filletsplit
    
    plreader.init tmproad
    roadsplit = plreader.getLongDDLine(nx, ny)
    'ThisDrawing.ModelSpace.AddLightWeightPolyline oldroadsplit
    
    'addVertex newroad, nx, ny
    
    
    tmpfillet.Delete
    tmproad.Delete
    
    
    Debug.Print ""
    
    checkCrossPoint3Ex = roadsplit
    
End Function


Function checkCrossPoint4Ex(road As AcadEntity, roadsplit() As Double, np As clsPoint) As Double()


    Dim tmpfillet As AcadLWPolyline
    Dim tmproad As AcadLWPolyline
    Dim intersec As Variant
    
    Dim selidx As Integer
    Dim mindist As Double
    
    Dim plreader As New ClsPolylineReader
    
    Dim x, y As Double
    Dim nx As Double
    Dim ny As Double
    
    Dim b1, b2 As Boolean
    
    selidx = -1
    mindist = 99999
    
    Set tmpfillet = ThisDrawing.ModelSpace.AddLightWeightPolyline(filletsplit)
    Set tmproad = ThisDrawing.ModelSpace.AddLightWeightPolyline(roadsplit)
    
    intersec = searchCrossVertext(tmpfillet, road)
    
    'addDonut5 np.x, np.y
    
    For i = 0 To UBound(intersec) Step 3
    
        x = intersec(i)
        y = intersec(i + 1)
    
        'b1 = isContains(tmpfillet, CDbl(x), CDbl(y))
        'b2 = isContains(tmproad, CDbl(x), CDbl(y))
        
        
        dist = getDistance(np.x, np.y, intersec(i), intersec(i + 1))
            
        'addDonutA intersec(i), intersec(i + 1)
            
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
    
    addDonutA nx, ny
    
    np2.x = nx
    np2.y = ny
    
    
    
    addVertex tmpfillet, nx, ny
    addVertexT2 tmproad, nx, ny
    
    plreader.init tmpfillet
    filletsplit = plreader.getLongDDLine(nx, ny)
    'Set tmpf = ThisDrawing.ModelSpace.AddLightWeightPolyline(filletsplit)
    'tmpf.color = acBlue
    
    plreader.init tmproad
    roadsplit = plreader.getLongDDLine(nx, ny)
    'Set tmpf = ThisDrawing.ModelSpace.AddLightWeightPolyline(oldroadsplit)
    'tmpf.color = acRed
    
    'addVertex newroad, nx, ny
    
    
    tmpfillet.Delete
    tmproad.Delete
    
    
    Debug.Print ""
    
    checkCrossPoint4Ex = roadsplit
    
End Function
