VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsPolylineReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private polyline As AcadLWPolyline
Private ddd() As Double


Public Function init(ent As AcadLWPolyline)

    Set polyline = ent
    
    ddd = polyline.Coordinates
    

End Function



Public Function getLongDDLine(vx As Double, vy As Double) As Double()

    
    Dim selidx As Integer
    Dim size As Integer
    Dim x As Double
    Dim y As Double
    
    Dim d1() As Double
    Dim d2() As Double
    
    Dim dist1 As Double
    Dim dist2 As Double
    
    selidx = -1
    
    size = UBound(ddd)
    
    For i = 0 To size Step 2
    
        x = ddd(i)
        y = ddd(i + 1)
        
        If x = vx And y = vy Then
        
            selidx = i
            Exit For
        
        End If
    
    Next i
    
    If selidx >= 0 Then
    
        ReDim dd1(selidx + 1) As Double
        ReDim dd2(size - (selidx + 1) + 1) As Double
    
        For i = 0 To selidx Step 2
            dd1(i) = ddd(i)
            dd1(i + 1) = ddd(i + 1)
        Next i
        
        For i = selidx To size Step 2
        
            dd2(i - selidx) = ddd(i)
            dd2(i - selidx + 1) = ddd(i + 1)
        
        Next i
    
    
    End If
    
    
    
    dist1 = getArrayDistance(dd1)
    dist2 = getArrayDistance(dd2)
    
    
    If dist1 < dist2 Then
        getLongDDLine = dd2
    
    Else
        getLongDDLine = dd1
    
    End If
    
    

End Function




