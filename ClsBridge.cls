VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ent As AcadEntity
Private polygon As New clsPolygonReader
Private sidx As Integer
Private eidx As Integer


'새 라인 정보
Private new1d() As Double
Private new2d() As Double

Private pts As New clsPoint
Private pte As New clsPoint


Private line As AcadLWPolyline
Private coords() As Double

Private tailLEN As Integer



Function createBridge()


    Dim selec1 As New ClsSelEntity
    Dim selec2 As New ClsSelEntity
    
    
    Dim pt1 As clsPoint
    Dim pt2 As clsPoint
    
    Dim p1() As Double
    Dim p2() As Double
    
    tailLEN = 5
    
    selec1.selectPolyline ("Get Polyline 1: ")
    Set ent = selec1.getEntity
    Set pt1 = selec1.getPoint

    selec2.selectPolyline ("Get Polyline 2: ")
    Set ent = selec2.getEntity
    Set pt2 = selec2.getPoint
    
    
    new1d = addVertex2(ent, pt1.x, pt1.y)
    new2d = addVertex2(ent, pt2.x, pt2.y)
    
    pts.x = new1d(0)
    pts.y = new1d(1)
    
    pte.x = new2d(0)
    pte.y = new2d(1)
    
    
    polygon.init ent
    sidx = polygon.getIndexByPoint(pts)
    eidx = polygon.getIndexByPoint(pte)
    
    Set line = getShortLine
    
    p1 = addHead()
    line.addVertex 0, p1
    line.Update
    
    p2 = addTail()
    size = (UBound(line.Coordinates) + 1) / 2
    
    line.addVertex size, p2
    
    Debug.Print ""
    

End Function


Function getShortLine() As AcadLWPolyline

    Dim line1 As AcadLWPolyline
    Dim line2 As AcadLWPolyline
    
    polygon.setIndex sidx
    Set line1 = getLine1
    
    polygon.setIndex sidx
    polygon.setReverse
    Set line2 = getLine1
    
    
    len1 = getLWPDistance(line1)
    len2 = getLWPDistance(line2)
    
    
    If len1 < len2 Then
        'LINE 1
        line2.Delete
        polygon.setReverse
    Else
        'LINE 2
        line1.Delete
        Set line1 = line2
    End If
    
    Set getShortLine = line1
    
End Function

    
Public Function addHead() As Double()

    Dim ptx As New clsPoint
    Dim ptn As New clsPoint
    
    Dim p1(0 To 1) As Double
    
    polygon.setIndex sidx
    Set ptn = polygon.nextPoint
    
    ptx.x = pts.x + 5
    ptx.y = pts.y
    
    'addDonutA pts.x, pts.y
    'addDonutA ptn.x, ptn.y
    'addDonutA ptx.x, ptx.y
    
    ang = getAngleEx(pts.x, pts.y, ptx.x, ptx.y, ptn.x, ptn.y)
    
    
    ang = ang - PI * 3 / 4
    
    xx = pts.x + tailLEN * Math.Cos(ang)
    yy = pts.y + tailLEN * Math.Sin(ang)
    
    
    'addDonutA xx, yy
     
    p1(0) = xx
    p1(1) = yy
    
    addHead = p1 'line1.addVertex 0, p1
     

End Function


  
Public Function addTail() As Double()

    
    Dim ptx As New clsPoint
    Dim ptn As New clsPoint
    
    Dim p1(0 To 1) As Double
    
    
    coords = line.Coordinates
    size = UBound(coords)
    
    ptn.x = coords(size - 3)
    ptn.y = coords(size - 2)
    
    pte.x = coords(size - 1)
    pte.y = coords(size)
    
    ptx.x = ptn.x + 5
    ptx.y = ptn.y
    
    'addDonutA pte.x, pte.y
    'addDonutA ptn.x, ptn.y
    'addDonutA ptx.x, ptx.y
    
    ang = getAngleEx(ptn.x, ptn.y, ptx.x, ptx.y, pte.x, pte.y)
    
    
    ang = ang - PI * 1 / 4
    
    xx = pte.x + tailLEN * Math.Cos(ang)
    yy = pte.y + tailLEN * Math.Sin(ang)
    
    
    'addDonutA xx, yy
     
    p1(0) = xx
    p1(1) = yy
    
    addTail = p1 'line1.addVertex 0, p1
     

End Function


Function getLine1() As AcadLWPolyline

    Dim pline As AcadLWPolyline
    Dim ptc As New clsPoint
    Dim pt As New clsPoint
    Dim ddd() As Double
    Dim size As Integer
    Dim eidx2 As Integer
    
    eidx2 = 0
    size = polygon.getSize
    
    
    Set ptc = polygon.getPoint(sidx)
    'addDonutA ptc.x, ptc.y
        
    For idx = 0 To size
        Set pt = polygon.nextPoint
        
        'addDonutA pt.x, pt.y
        
        If pte.isEquals(pt) = True Then
            eidx2 = idx + 1
            Exit For
        End If
        
    Next
    
    ReDim ddd(0 To eidx2 * 2 + 1) As Double
        
    ddd(0) = ptc.x
    ddd(1) = ptc.y
    
    
    polygon.setIndex sidx
        
    For idx = 1 To eidx2
        Set pt = polygon.nextPoint
        ddd(idx * 2) = pt.x
        ddd(idx * 2 + 1) = pt.y
    Next
    
    Set getLine1 = ThisDrawing.ModelSpace.AddLightWeightPolyline(ddd)
    
    
    
End Function

Public Function getLWPDistance(ent As AcadLWPolyline) As Double


    Dim ddd() As Double
    Dim size As Integer
    Dim x1, y1, x2, y2 As Double
    Dim sum As Double
    
    sum = 0
    ddd = ent.Coordinates
    
    size = UBound(ddd)
    
    
    x1 = ddd(0)
    y1 = ddd(1)
    
    For i = 2 To size Step 2
    
        x2 = ddd(i)
        y2 = ddd(i + 1)
        
        sum = sum + getDistance(x1, y1, x2, y2)
        
        x1 = x2
        y1 = y2
    
    
    Next

    getLWPDistance = sum


End Function



Public Function setCoords()

    coords = line.Coordinates
    

End Function




