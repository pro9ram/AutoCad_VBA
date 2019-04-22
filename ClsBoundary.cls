VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsBoundary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private baseents() As AcadEntity
Private dstents() As AcadEntity

Private baseregs As Variant
Private dstregs As Variant


Private dstcount As Integer
Private idx As Integer


Private regs As Variant








Public Function selects()

    Dim conv As New ClsConverter
    Dim reg As AcadRegion
    Dim reg2 As AcadRegion
    
    idx = 0
    
    

    baseents = selectPLObjs
    dstents = selectPLObjs
    
    If IsArrayEmpty(baseents) = False And IsArrayEmpty(dstents) = False Then
        
        dstcount = UBound(dstents)
        
        
        
        dstregs = ThisDrawing.ModelSpace.AddRegion(dstents)
        
        For i = 0 To UBound(dstregs)
        
            Set reg = dstregs(i)
            
            baseregs = ThisDrawing.ModelSpace.AddRegion(baseents)
            
            For j = 0 To UBound(baseregs)
                Set reg2 = baseregs(j)
                reg.Boolean acSubtraction, reg2
            Next j
            
            If reg.Area > 0 Then
                expoldeRegion reg
            End If
            'conv.reg2polyline reg
            'reg.Delete
            
        Next i
        
        
        For i = 0 To UBound(dstents)
            dstents(i).Delete
        
        Next i
    End If
    
    
    'explodeRegions
    
    Debug.Print ""
    'reg2.Boolean acSubtraction, reg3
    
    
    
    'expoldeRegion reg2
    
    'expoldeRegion reg1
     

End Function

 




Private Function expoldeRegion(reg As Variant)

    Dim has As Boolean
    Dim conv As New ClsConverter
    Dim Punto As Variant

    has = hasSubRegion(reg)
    
    If has = True Then
    
        Punto = reg.Explode
        reg.Delete
        
        For i = 0 To UBound(Punto)
            conv.reg2polyline Punto(i)
        Next i
        
        For i = 0 To UBound(Punto)
            Punto(i).Delete
        Next i
    
    Else
    
        conv.reg2polyline reg
        reg.Delete
    
    End If
    


End Function


Private Function IsArrayEmpty(anArray As Variant) As Boolean

    Dim ret As Boolean
    
    ret = True

    On Error GoTo IS_EMPTY
    If (UBound(anArray) >= 0) Then
    
        ret = False
    
    End If
IS_EMPTY:
        IsArrayEmpty = ret
End Function

Private Function hasSubRegion(reg As Variant) As Boolean

    Dim arr As Variant
    Dim size As Integer
    Dim ret As Boolean

    ret = True

    arr = reg.Explode
    size = UBound(arr)
    
    If size > 0 Then
        If TypeOf arr(0) Is AcadLine Then
            ret = False
        
        End If
    End If
    
    
    For i = 0 To size
        arr(i).Delete
    
    Next
    
    hasSubRegion = ret

End Function
