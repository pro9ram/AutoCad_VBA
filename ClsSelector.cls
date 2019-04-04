VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ent As AcadEntity
Private size As Integer


Public Function init(e As AcadEntity)
    Set ent = e
    
    coord = ent.Coordinates
    size = UBound(coord)
    
End Function

