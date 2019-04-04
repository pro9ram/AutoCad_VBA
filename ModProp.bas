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
    
    
    Set ent = selectEntity()
    
    Set amap = ThisDrawing.Application.GetInterfaceObject("AutoCADMap.Application")
    Set tbls = amap.Projects(ThisDrawing).ODTables
    
    
    layername = ent.Layer
    
    For Each tbl In tbls
    
    
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
