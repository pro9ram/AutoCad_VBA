Attribute VB_Name = "Excutor"
Sub createFillet()

    Dim fillet As New clsFillet
    
    ThisDrawing.EndUndoMark
    ThisDrawing.StartUndoMark
    fillet.load
    
     

End Sub


Sub createFillet2()
    
    

    Dim fillet2 As New ClsFillet2
    
        
    
    ThisDrawing.EndUndoMark
    ThisDrawing.StartUndoMark
    
    fillet2.load
    

End Sub




Sub createFilletT()

    
    Dim filletT As New ClsFilletT
    
    
    ThisDrawing.EndUndoMark
    ThisDrawing.StartUndoMark
    filletT.load
    
    

End Sub



Sub createChamfer()

    Dim chamfer As New ClsChamfer
    
    
    ThisDrawing.EndUndoMark
    ThisDrawing.StartUndoMark
    chamfer.load
    
     

End Sub


Sub createChamfer2()

    Dim chamfer2 As New clsChamfer2
        
    
    ThisDrawing.EndUndoMark
    ThisDrawing.StartUndoMark
    chamfer2.load
     
    

End Sub



Sub createChamferT()

    Dim chamferT As New ClsChamferT
     
    
    ThisDrawing.EndUndoMark
    ThisDrawing.StartUndoMark
    
    chamferT.load
     
    

End Sub


Sub revert()

    ThisDrawing.EndUndoMark
    'ThisDrawing.SendCommand "Undo" & vbCr & "Back" & vbCr & "N" & vbCr
    
    
End Sub


Sub createBridge()
    
    Dim pp As New ClsBridge
    pp.createBridge
    
End Sub

Sub showSettingsForm()
    Dim myUserForm As New frmSettings

    myUserForm.frmSettings_init
    myUserForm.show


End Sub



Sub splitBoundary()

    Dim token() As String

    ThisDrawing.StartUndoMark

    saveOnLayers
    hideLayerAll
    
    token = getTempSplit2
    
    showLayerLike token
    
    ThisDrawing.SendCommand "bbc" & vbCr
    
    restoreOnLayers
    
    ThisDrawing.EndUndoMark

End Sub


