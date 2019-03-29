VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsChamfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private oldroad As AcadEntity
Private newroad As AcadEntity
Private fillet As AcadEntity

Private ptIntersection As Variant


Public Function load()

    Dim p1 As New clsPoint
    Dim p2() As Double
    Dim p3 As New clsPoint
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
    
    
    ThisDrawing.SetVariable "osmode", 512
    ThisDrawing.Utility.getEntity newroad, tmp, "Enter a point 1: "
        
    p1.initPt tmp
     
    ThisDrawing.SetVariable "osmode", 512
    ThisDrawing.Utility.getEntity oldroad, tmp, "Enter a point 3: "
    
    p3.initPt tmp
    
    
    Debug.Print "p1: " & p1.x & ", " & p1.y
    Debug.Print "p3: " & p3.x & ", " & p3.y
    
    
    addDonut5 p1.x, p1.y
    addDonut5 p3.x, p3.y
    
    Dim vlist(0 To 3) As Double
    
    vlist(0) = p1.x
    vlist(1) = p1.y
    vlist(2) = p3.x
    vlist(3) = p3.y
    
    
    
    Set retArc = ThisDrawing.ModelSpace.AddLightWeightPolyline(vlist)
    
    'retArc.Delete
    
    
    Debug.Print "fillet created"
    
    
    ptIntersection = searchCrossVertext(newroad, oldroad)
    
    newp1 = findVertex(newroad, retArc)   '아크와 도로(새도로)의 접점 추가
    newp2 = findVertex(oldroad, retArc)   '아크와 도로(구도로)의 접점 추가
   
   
    np1.initXy newp1(0), newp1(1)
    np2.initXy newp2(0), newp2(1)
    
    
    Debug.Print "np1: " & np1.x & ", " & np1.y
    Debug.Print "np2: " & np2.x & ", " & np2.y
    
   
    Set np3 = getFarPoint2(ptIntersection, np2)     '기존도로 아크 접점에서 먼 지점을 찾는다
    
    
    addVertex oldroad, np3.x, np3.y     '기존도로에 접점추가
    

   
    Debug.Print "cross point created"
    
    pr.init newroad
    pr.setIndexByPoint np1
    pr.checkDirection p3
    
    'addDonut5 np3.x, np3.y
    'addDonut5 np1.x, np1.y
    Debug.Print "newroad direction processed"
    
    newroadsplit = pr.splitInnerLine(np3, np1)
    'addLine newroadsplit
    
    Debug.Print "newroad split created"
   
    pr.init oldroad
    pr.setIndexByPoint np2
    pr.checkDirection p1
    
    Debug.Print "oldroad direction processed"
    
    oldroadsplit = pr.splitInnerLine(np2, np3)
    'addLine oldroadsplit
    
    
    
    Debug.Print "oldroad split created"
        
     size1 = UBound(newroadsplit)
    size2 = UBound(vlist)
    size3 = UBound(oldroadsplit)
    
    merge.setSize size1 + size2 + size3 + 2
    
    
    merge.add newroadsplit
    merge.add vlist
    merge.add oldroadsplit
    
    
    addLine merge.getMerged
    newroad.Delete
    retArc.Delete
    
    
    
    'ThisDrawing.SendCommand "sob" & vbCr
    
End Function
 

 
