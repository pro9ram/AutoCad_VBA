Attribute VB_Name = "ModInit"
Public Const TEMP_LAYER = "temp_layer"
Public Const SHOW_CIRCLE = False

Public Const P2P_TOL = 100    ' 1~100: 1 ------100 민감
Public Const PD_TOL = 0.1     ' 0.1~1: 민감 0.1 ------ 1
Public Const FAR_OVER = 30   '접점이 여러개인데 아주 멀리 떨어진경우를 찾아 skip. () <- 이런경우


Public Const KEY_TOL = "userr1"
Public Const KEY_RAD = "userr2"
Public Const KEY_CHA = "userr3"
Public Const DEF_TOL = 0.1
Public Const DEF_RAD = 3
Public Const DEF_CHA = 3


Public Const KEY_FAR_DIST = "useri1"
Public Const DEF_FAR_DIST = FAR_OVER



Public Const KEY_BRD = "users1"
Public Const DEF_BRD = "RD_SISUL_L_대교_170701"


Public Const KEY_LAYER_ON = "users4"
Public Const DEF_LAYER_ON = "도로*|!!*"

Public Const KEY_TEMP_SPLIT = "users5"
'Public Const KEY_TEMP_SPLIT2 = "users4"
'Public Const DEF_BRD = "RD_SISUL_L_대교_170701"



Public Const LYR_ADD = "!!추가_"
Public Const LYR_DEL = "!!삭제_"







Public Function initConstant()

    Dim tol As Double
    Dim rad As Double
    Dim fardist As Integer
    
    tol = ThisDrawing.GetVariable(KEY_TOL)
    rad = ThisDrawing.GetVariable(KEY_RAD)
    fardist = ThisDrawing.GetVariable(KEY_FAR_DIST)


    If tol = 0 Then
        tol = DEF_TOL
        ThisDrawing.SetVariable KEY_TOL, tol
    End If


    If rad = 0 Then
        rad = DEF_RAD
        ThisDrawing.SetVariable KEY_RAD, rad
    End If


    If fardist = 0 Then
        fardist = DEF_FAR_DIST
        ThisDrawing.SetVariable KEY_FAR_DIST, fardist
    End If


End Function


Public Function setDefaultFilletRadius(rad As Double)

    ThisDrawing.SetVariable KEY_RAD, rad

End Function

Public Function getDefaultFilletRadius() As Double

    Dim rad As Double
    rad = ThisDrawing.GetVariable(KEY_RAD)
    
    If rad = 0 Then
        rad = DEF_RAD
        ThisDrawing.SetVariable KEY_RAD, rad
    End If
    
    getDefaultFilletRadius = rad
    
End Function


Public Function setDefaultTolerance(tol As Double)

    ThisDrawing.SetVariable KEY_TOL, tol

End Function


Public Function getDefaultTolerance()

    Dim tol As Double
    tol = ThisDrawing.GetVariable(KEY_TOL)
   
    If tol = 0 Then
        tol = DEF_TOL
        ThisDrawing.SetVariable KEY_TOL, tol
    End If

    getDefaultTolerance = tol
    
End Function


Public Function setDefaultChamferDistance(cd As Double)

    ThisDrawing.SetVariable KEY_CHA, cd

End Function


Public Function getDefaultChamferDistance()

    Dim cd As Double
    cd = ThisDrawing.GetVariable(KEY_CHA)
   
    If cd = 0 Then
        cd = DEF_CHA
        ThisDrawing.SetVariable KEY_CHA, cd
    End If

    getDefaultChamferDistance = cd
    
End Function



Public Function setDefaultFarDistance(cd As Double)

    ThisDrawing.SetVariable KEY_FAR_DIST, cd

End Function


Public Function getDefaultFarDistance()

    Dim cd As Double
    cd = ThisDrawing.GetVariable(KEY_FAR_DIST)
   
    If cd = 0 Then
        cd = DEF_FAR_DIST
        ThisDrawing.SetVariable KEY_FAR_DIST, cd
    End If

    getDefaultFarDistance = cd
    
End Function



Public Function setDefaultBridgeLayer(layer As String)

    ThisDrawing.SetVariable KEY_BRD, layer

End Function

Public Function getDefaultBridgeLayer() As String

    Dim layer As String
    layer = ThisDrawing.GetVariable(KEY_BRD)
    
    If layer = Empty Then
        layer = DEF_BRD
        ThisDrawing.SetVariable KEY_BRD, layer
    End If
    
    getDefaultBridgeLayer = layer
    
End Function


Public Function setDefaultLayerOn(layer As String)

    ThisDrawing.SetVariable KEY_LAYER_ON, layer

End Function

Public Function getDefaultLayerOn() As String()

    Dim layer As String
    Dim layers() As String
    layer = ThisDrawing.GetVariable(KEY_LAYER_ON)
    
    If layer = Empty Then
        layer = DEF_LAYER_ON
        ThisDrawing.SetVariable KEY_LAYER_ON, layer
    End If
    
    
    layers() = split(layer, "|")
        
    getDefaultLayerOn = layers
    
End Function


Public Function getDefaultLayerOnString() As String

    Dim layer As String

    layer = ThisDrawing.GetVariable(KEY_LAYER_ON)
    
    If layer = Empty Then
        layer = DEF_LAYER_ON
        ThisDrawing.SetVariable KEY_LAYER_ON, layer
    End If
    
        
    getDefaultLayerOnString = layer
    
End Function



Public Function setTempSplit(result As String)  '1|2|3|

    ThisDrawing.SetVariable KEY_TEMP_SPLIT, result

End Function


Public Function getTempSplit() As String()

    Dim str() As String
    Dim result As String
    result = ThisDrawing.GetVariable(KEY_TEMP_SPLIT)
    
    If result <> Empty Then
        str() = split(result, "|")
        
    End If
    
    getTempSplit = str
    
End Function

