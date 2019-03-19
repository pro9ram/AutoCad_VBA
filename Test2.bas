Attribute VB_Name = "Test2"
Private Const PI As Double = 3.14159265358979
Private Const PI_2 As Double = 1.5707963267949


Sub Createcircle2()
    Dim i, ii, X, Y As Integer
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim UsersSelection As AcadSelectionSet
    Dim selObj() As AcadEntity
    Dim ret As Integer
    Dim arcObjs() As AcadEntity
    Dim ddd() As Double
    Dim x2, y2 As Variant
    Dim count As Integer
    Dim newv(0 To 1) As Double
    Dim centerPoint(0 To 2) As Double
    
    Dim returnObj1 As AcadObject
    Dim returnObj3 As AcadObject
    
    Dim retArc As AcadArc
    
    Dim p1(0 To 2) As Double
    Dim p2() As Double
    Dim p3(0 To 2) As Double
    
    Dim retline As AcadEntity
    
    With ThisDrawing
        On Error Resume Next
        .SelectionSets("CurrentSelection").Delete
        Set UsersSelection = .SelectionSets.Add("CurrentSelece e tion")

        
        ThisDrawing.SetVariable "osmode", 512
        ThisDrawing.Utility.GetEntity returnObj1, tmp, "Enter a point 1: "
            
        p1(0) = tmp(0)
        p1(1) = tmp(1)
            
        ThisDrawing.SetVariable "osmode", 0
        p2 = ThisDrawing.Utility.GetPoint(, "Enter a point2: ")
         
        ThisDrawing.SetVariable "osmode", 512
        ThisDrawing.Utility.GetEntity returnObj3, tmp, "Enter a point 3: "
        
        p3(0) = tmp(0)
        p3(1) = tmp(1)
        
        
        Set retArc = getArcBy3P(p1, p2, p3)
        
        
        Set retline = arc2line(retArc)
        retArc.Delete
        retArc.Update
        retline.Update
        
    
        Debug.Print ""
         
    End With
    
   
   findVertex returnObj1, retline
   'findVertex returnObj2, retline
   
   
    
    
End Sub





