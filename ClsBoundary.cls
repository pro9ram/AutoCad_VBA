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
    Dim lyrname As String
    Dim dstent As Variant
    Dim result As Boolean
    
    idx = 0
    
    

    baseents = selectPLObjs
    dstents = selectPLObjs
    
    If IsArrayEmpty(baseents) = False And IsArrayEmpty(dstents) = False Then
        
        dstcount = UBound(dstents)
        
        
        For k = 0 To UBound(dstents)
        
            ReDim dstent(0 To 0) As AcadEntity
            Set dstent(0) = dstents(k)              '한개씩 처리
 
            lyrname = dstents(k).layer               '현재레이어 설정
        
            dstregs = ThisDrawing.ModelSpace.AddRegion(dstent)
            
            result = False
            
            For i = 0 To UBound(dstregs)            '하나씩 처리하니 region도 하나겠지?
            
                
                Set reg = dstregs(i)
                
                baseregs = ThisDrawing.ModelSpace.AddRegion(baseents)       'base 도로는 subtract로 잘리기때문에 매번 생성함
                
                For j = 0 To UBound(baseregs)
                    Set reg2 = baseregs(j)
                    reg.Boolean acSubtraction, reg2
                Next j
                
                If reg.Area > 0 Then
                    result = expoldeRegion(reg, lyrname)
                    
          
                    If result = True Then
                        dstent(0).Delete        '올바르게 처리되었을때만 원본을 삭제함
                    Else
                        reg.Delete
                    End If
                End If
                
            Next i
            
                      
        
        Next k
        
    End If
    
    
    'explodeRegions
    
    Debug.Print ""
    'reg2.Boolean acSubtraction, reg3
    
    
    
    'expoldeRegion reg2
    
    'expoldeRegion reg1
     

End Function

 




Private Function expoldeRegion(reg As Variant, lyrname As String) As Boolean

    Dim has As Boolean
    Dim conv As New ClsConverter
    Dim Punto As Variant
    Dim ent As AcadLWPolyline
    Dim ret As Boolean
    Dim arr As Variant
    Dim subreg As Variant

    ret = False
    
    has = hasSubRegion(reg)
    
    If has = True Then
    
        arr = reg.Explode
        size = UBound(arr)
        
        For i = 0 To size
            
            Set subreg = arr(i)
            
            has = hasArc(subreg)
        
            If has = False Then
            
                Set ent = conv.reg2polyline(subreg)
                ent.layer = lyrname
                subreg.Delete
                ret = True
            End If
        
        Next i
        
        reg.Delete
    
    
    Else
    
        has = hasArc(reg)
        
        If has = False Then
        
            Set ent = conv.reg2polyline(reg)
            ent.layer = lyrname
            reg.Delete
            ret = True
        End If
    End If
    

    expoldeRegion = ret

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

    ret = False

    arr = reg.Explode
    size = UBound(arr)
    
    If size > 0 Then
        If TypeOf arr(0) Is AcadRegion Then
            ret = True
        
        End If
    End If
    
    
    For i = 0 To size
        arr(i).Delete
    
    Next
    
    hasSubRegion = ret

End Function


Private Function hasArc(reg As Variant) As Boolean

    Dim arr As Variant
    Dim size As Integer
    Dim ret As Boolean

    ret = False

    arr = reg.Explode
    size = UBound(arr)
    
    For i = 0 To size
        If TypeOf arr(i) Is AcadArc Then
            arr(i).color = acBlue
            ret = True
            Exit For
        End If
    
    Next
    
    For i = 0 To size
        arr(i).Delete
    
    Next
    
    hasArc = ret

End Function

