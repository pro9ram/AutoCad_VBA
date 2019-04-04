VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsSelEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const VK_ESCAPE = &H1B, VK_ENTER = &HD, VK_LBUTTON = &H1
'Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public ent As AcadEntity
Public pt As New clsPoint


Public Function selectPolyline(prompt As String)

    Dim v As Variant

   On Error GoTo errControl
    
Start:
    ThisDrawing.SetVariable "osmode", 512
    ThisDrawing.Utility.getEntity ent, v, prompt
    
errControl:
    
    If Err.Description = "Method 'GetEntity' of object 'IAcadUtility' failed" Then
    
        If CheckKey(VK_ESCAPE) = True Or CheckKey(VK_ENTER) Then
            Debug.Print "End"
        
            End
        Else
            Debug.Print "Resume"
            Resume
        End If
    Else
    
        pt.x = v(0)
        pt.y = v(1)
    
        Debug.Print ""
        
        
        
    End If
    

End Function



Public Function getEntity() As AcadEntity

    Set getEntity = ent

End Function


Public Function getPoint() As clsPoint

    Set getPoint = pt

End Function
