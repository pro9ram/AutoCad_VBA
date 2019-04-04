VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSelReal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Public Function selectReal()

    Dim rad As Double

    On Error GoTo errControl

    rad = getDefaultFilletRadius
    rad = ThisDrawing.Utility.GetReal("Radius (default: " & rad & ")   ")
    
    setDefaultFilletRadius rad
    
errControl:
    
    If Err.Description = "User input is a keyword" Then
        rad = getDefaultFilletRadius
    Else
    
    End If
    

    selectReal = rad

End Function


Public Function selectFilletRadius()

    Dim rad As Double

    On Error GoTo errControl

    rad = getDefaultFilletRadius
    rad = ThisDrawing.Utility.GetReal("#Radius (default: " & rad & ")   ")
    
    setDefaultFilletRadius rad
    
errControl:
    
    If Err.Description = "User input is a keyword" Then
        rad = getDefaultFilletRadius
    Else
    
    End If
    

    selectFilletRadius = rad

End Function


Public Function selectChamferDistance()

    Dim dist As Double

    On Error GoTo errControl

    dist = getDefaultChamferDistance
    dist = ThisDrawing.Utility.GetReal("#Distance (default: " & dist & ")   ")
    
    setDefaultChamferDistance dist
    
errControl:
    
    If Err.Description = "User input is a keyword" Then
        dist = getDefaultChamferDistance
        
    Else
    
    End If
    

    selectChamferDistance = dist

End Function

