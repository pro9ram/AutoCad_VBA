Attribute VB_Name = "selector"

Const VK_ESCAPE = &H1B, VK_ENTER = &HD, VK_LBUTTON = &H1
Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Function CheckKey(lngKey As Long) As Boolean
    If GetAsyncKeyState(lngKey) Then
        CheckKey = True
        Debug.Print "True"
        
    Else
        CheckKey = False
        Debug.Print "False"
    End If
End Function


Function selectArcObs() As AcadEntity()
    Dim UsersSelection As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim arcObjs() As AcadEntity
 
     With ThisDrawing
        On Error Resume Next
        
        .SelectionSets("CurrentSelection").Delete
        Set UsersSelection = .SelectionSets.add("CurrentSelection")
        
        FilterType(0) = 0
        FilterData(0) = "Arc"
        
        UsersSelection.SelectOnScreen FilterType, FilterData
        
        
        count = getEntitySize(UsersSelection, "AcDbArc")
        
        ReDim arcObjs(0 To count - 1) As AcadEntity
        
        For i = 0 To count
            Set arcObjs(i) = UsersSelection.Item(i)
        Next i
        
        
    End With
    
    selectArcObs = arcObjs

End Function



Function selectArcObj() As AcadEntity
    Dim UsersSelection As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim arcObjs As AcadEntity
    Dim count As Integer
 
     With ThisDrawing
        On Error Resume Next
        
        .SelectionSets("CurrentSelection").Delete
        Set UsersSelection = .SelectionSets.add("CurrentSelection")
        
        FilterType(0) = 0
        FilterData(0) = "Arc"
        
        UsersSelection.SelectOnScreen FilterType, FilterData
        
        
        count = getEntitySize(UsersSelection, "AcDbArc")
        
        For i = 0 To count
            Set arcObjs = UsersSelection.Item(i)
            Exit For
        Next i
        
        
    End With
    
    Set selectArcObj = arcObjs

End Function

 

Function selectPolyline()

    Dim newroad As AcadEntity
    Dim basePnt1 As Variant
        
    On Error GoTo Start
    
Start:
    ThisDrawing.Utility.getEntity newroad, basePnt1, "Get Object1: "
    
    
    Debug.Print ""
    
    

End Function


Function selectObjs() As AcadEntity()
    Dim UsersSelection As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim arcObjs() As AcadEntity
 
     With ThisDrawing
        On Error Resume Next
        
        .SelectionSets("CurrentSelection").Delete
        Set UsersSelection = .SelectionSets.add("CurrentSelection")
        
        UsersSelection.SelectOnScreen
        
        ReDim arcObjs(0 To UsersSelection.count - 1) As AcadEntity
        For i = 0 To UsersSelection.count
            Set arcObjs(i) = UsersSelection.Item(i)
        Next i
        
        
    End With
    
    selectObjs = arcObjs

End Function


Function selectEntity() As AcadEntity
    Dim UsersSelection As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim arcObjs() As AcadEntity
    Dim returnObj1 As AcadObject
    Dim basePnt1 As Variant
 
    ThisDrawing.Utility.getEntity returnObj1, basePnt1, "Get Object: "
    
    Set selectEntity = returnObj1

End Function


Function selectSubEntity() As AcadEntity
    Dim UsersSelection As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim arcObjs() As AcadEntity
    Dim returnObj1 As AcadObject
    Dim basePnt1 As Variant
 
    ThisDrawing.Utility.GetSubEntity
    
    selectObj = returnObj1

End Function


Function selectlineObs() As AcadEntity()
    Dim UsersSelection As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim arcObjs() As AcadEntity
 
     With ThisDrawing
        On Error Resume Next
        
        .SelectionSets("CurrentSelection").Delete
        Set UsersSelection = .SelectionSets.add("CurrentSelection")
        
        FilterType(0) = 0
        FilterData(0) = "Line"
        
        UsersSelection.SelectOnScreen FilterType, FilterData
        
        
        count = getEntitySize(UsersSelection, "AcDbLine")
        
        ReDim arcObjs(0 To count - 1) As AcadEntity
        
        For i = 0 To count
            Set arcObjs(i) = UsersSelection.Item(i)
        Next i
        
        
    End With
    
    selectlineObs = arcObjs

End Function



Function selectPolylineObs() As AcadEntity()
    Dim UsersSelection As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim arcObjs() As AcadEntity
 
     With ThisDrawing
        On Error Resume Next
        
        .SelectionSets("CurrentSelection").Delete
        Set UsersSelection = .SelectionSets.add("CurrentSelection")
        
        FilterType(0) = 0
        FilterData(0) = "LWPolyline"
        
        UsersSelection.SelectOnScreen FilterType, FilterData
        
        
        count = getEntitySize(UsersSelection, "AcDbPolyline")
        
        ReDim arcObjs(0 To count - 1) As AcadEntity
        
        For i = 0 To count
            Set arcObjs(i) = UsersSelection.Item(i)
        Next i
        
        
    End With
    
    selectPolylineObs = arcObjs

End Function


Function selectObjs2() As AcadEntity()
    Dim UsersSelection As AcadSelectionSet
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim arcObjs() As AcadEntity
 
 
 
     With ThisDrawing
        On Error Resume Next
                
        
        .SelectionSets("CurrentSelection").Delete
        Set UsersSelection = .SelectionSets.add("CurrentSelection")
        
        UsersSelection.SelectOnScreen
        
        ReDim arcObjs(0 To UsersSelection.count - 1) As AcadEntity
        For i = 0 To UsersSelection.count
            Set arcObjs(i) = UsersSelection.Item(i)
        Next i
        
        
    End With
    
    
    selectObjs = arcObjs

End Function
