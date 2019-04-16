Attribute VB_Name = "ModProp"
Sub show_properties()

    Dim ent As AcadEntity
    Dim tbls As ODTables
    Dim tbl As ODTable
    Dim recrds As odrecords
    
    Dim recrd As odrecord
    Dim fieldv As odfieldvalue
    
    Dim fieldds As ODFieldDefs
    Dim fieldd As ODFieldDef
        
    Dim amap As AcadMap
    Dim layername As String
    Dim xdata As Variant
    
    Dim adic As AcadDictionary
    Dim vvv As Variant
    
    
    
    Set ent = selectEntity()
    
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    Set tbls = amap.Projects(ThisDrawing).ODTables
    
    
    layername = ent.Layer
    Debug.Print ent.ObjectName & " " & ent.ObjectID & " " & ent.OwnerID & " " & ent.PlotStyleName
    Debug.Print " " & ent.HasExtensionDictionary
    
    
    Set adic = ent.GetExtensionDictionary()
    Set vvv = adic.Item(0)
    
        
    Dim xdataOut As Variant
    Dim xtypeOut As Variant
    ent.GetXData "", xtypeOut, xdataOut
    
    layername = "bg_rd_walk_l"
    
    For Each tbl In tbls
            
Try:
        On Error GoTo Catch
        
        Set odrecrds = tbl.GetODRecords
        odrecrds.init ent, True, False
        
        If odrecrds.IsDone = False Then
            Set odrecrd = odrecrds.record
            Set fieldds = tbl.ODFieldDefs
            
            count = fieldds.count
            
            
            For i = 0 To count - 1
                Set fieldd = fieldds.Item(i)
                Set fieldv = odrecrd.Item(i)
            
                Debug.Print i & ")" & fieldd.name & ": " & fieldv.Value
            
            Next
        End If
      
        GoTo Finally
Catch:

Finally:
        
    Next
    
End Sub

Sub show_properties_old()

    Dim ent As AcadEntity
    Dim tbls As ODTables
    Dim tbl As ODTable
    Dim recrds As odrecords
    
    Dim recrd As odrecord
    Dim fieldv As odfieldvalue
    
    Dim fieldds As ODFieldDefs
    Dim fieldd As ODFieldDef
        
    Dim amap As AcadMap
    Dim layername As String
    Dim xdata As Variant
    
    Dim adic As AcadDictionary
    Dim vvv As Variant
    
    
    
    Set ent = selectEntity()
    
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    Set tbls = amap.Projects(ThisDrawing).ODTables
    
    
    layername = ent.Layer
    Debug.Print ent.ObjectName & " " & ent.ObjectID & " " & ent.OwnerID & " " & ent.PlotStyleName
    Debug.Print " " & ent.HasExtensionDictionary
    
    
    Set adic = ent.GetExtensionDictionary()
    Set vvv = adic.Item(0)
    
        
    Dim xdataOut As Variant
    Dim xtypeOut As Variant
    ent.GetXData "", xtypeOut, xdataOut
    
    layername = "bg_rd_width_sudo_"
    
    For Each tbl In tbls
    
        Debug.Print ">>>>>>>>>> " & tbl.name & " " & tbl.Description
        
    
        If tbl.name = "Default_" & layername Then
            Set odrecrds = tbl.GetODRecords
            odrecrds.init ent, True, False
            
            Set odrecrd = odrecrds.record
            Set fieldds = tbl.ODFieldDefs
            
            count = fieldds.count
            
            
            For i = 0 To count - 1
                Set fieldd = fieldds.Item(i)
                Set fieldv = odrecrd.Item(i)
            
                Debug.Print i & ")" & fieldd.name & ": " & fieldv.Value
            
            Next
             
        End If
    Next
    
End Sub
