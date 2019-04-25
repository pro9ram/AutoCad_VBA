VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsChamfer"
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
    Dim pt3 As New clsPoint
    Dim retLine As AcadLine
    Dim tmp As Variant
    
    Dim sp(0 To 2) As Double
    Dim ep(0 To 2) As Double
    
    Dim np1 As New clsPoint
    Dim np2 As New clsPoint
    Dim np3 As New clsPoint
    
    Dim splitter As New clsSplitter
    Dim pr As New clsPolygonReader
    
    Dim newroadsplit() As Double
    Dim oldroadsplit() As Double
    Dim additionsplit(0 To 3) As Double
    
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
    
    selec2.selectPolyline ("#Get Polyline 2: ")
    Set oldroad = selec2.getEntity
    Set pt3 = selec2.getPoint
    

    
    sp(0) = pt1.x
    sp(1) = pt1.y
    ep(0) = pt3.x
    ep(1) = pt3.y
    
    
    Set retLine = ThisDrawing.ModelSpace.addLine(sp, ep)
    
    'retArc.Delete
    
    
    Debug.Print "fillet created"
    
    
    ptIntersection = searchCrossVertext(newroad, oldroad)
    
    newp1 = findVertexAtLine2(newroad, retLine, pt1)   '아크와 도로(새도로)의 접점 추가
    newp2 = findVertexAtLine2(oldroad, retLine, pt3)   '아크와 도로(구도로)의 접점 추가
   
   
    np1.initXy newp1(0), newp1(1)
    np2.initXy newp2(0), newp2(1)
    
    
    Debug.Print "np1: " & np1.x & ", " & np1.y
    Debug.Print "np2: " & np2.x & ", " & np2.y
    

    additionsplit(0) = np1.x
    additionsplit(1) = np1.y
    additionsplit(2) = np2.x
    additionsplit(3) = np2.y
   
    Set np3 = getFarPoint2(ptIntersection, np2)     '기존도로 아크 접점에서 먼 지점을 찾는다
    
    
    addVertex oldroad, np3.x, np3.y     '기존도로에 접점추가
    

   
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
    
    oldroadsplit = pr.splitInnerLine(np2, np3)
    'addLine oldroadsplit
    
    
    
    Debug.Print "oldroad split created"
        
    size1 = UBound(newroadsplit)
    size2 = UBound(additionsplit)
    size3 = UBound(oldroadsplit)
    
    merge.setSize size1 + size2 + size3 + 2
    
    
    merge.add newroadsplit
    merge.add additionsplit
    merge.add oldroadsplit
    
    
    Set result = addLine(merge.getMerged)
    result.Closed = True
    'newroad.Delete
    retLine.Delete
    
    'ThisDrawing.SendCommand "sob" & vbCr
    
    orgLayer = newroad.layer
    result.layer = getAddLayerName(orgLayer)
    newroad.layer = getDelLayerName(orgLayer)
    
    If orgLayer Like LYR_ADD & "*" Then
        newroad.Delete
    End If
    
    
End Function
 

 
