Attribute VB_Name = "chamferModule"


Function findVertexAtLine(ent As AcadEntity, ent2 As AcadLine) As Double()

    Dim ddd(0 To 3) As Double
    Dim filletpoint(0 To 7) As Double   '8
        
    
    
    ddd(0) = ent2.startPoint(0)
    ddd(1) = ent2.startPoint(1)
    ddd(2) = ent2.endPoint(0)
    ddd(3) = ent2.endPoint(1)
            
    count = UBound(ddd)
    
    vx = ddd(count - 1)
    vy = ddd(count)
    
    newp = addVertex(ent, vx, vy)
    
    If newp(0) <> 0 And newp(1) <> 0 Then
    
        filletpoint(index) = newp(0)
        index = index + 1
        filletpoint(index) = newp(1)
        index = index + 1
         
    Else
        vx = ddd(0)
        vy = ddd(1)
        
        newp = addVertex(ent, vx, vy)
        
        If newp(0) <> 0 And newp(1) <> 0 Then
        
            filletpoint(index) = newp(0)
            index = index + 1
            filletpoint(index) = newp(1)
            index = index + 1
        End If
    End If
    
    findVertexAtLine = newp

End Function

Function findVertexAtLine2(ent As AcadEntity, ent2 As AcadLine, pt As clsPoint) As Double()

    Dim filletpoint(0 To 7) As Double   '8
    Dim nv() As Double
    
    nv = getNearestVertexAtLine(ent2, pt)
    vx = nv(0)
    vy = nv(1)
    
    newp = addVertex2(ent, vx, vy)
    
    If newp(0) <> 0 And newp(1) <> 0 Then
    
        filletpoint(index) = newp(0)
        index = index + 1
        filletpoint(index) = newp(1)
        index = index + 1
   
    End If
    
    findVertexAtLine2 = newp

End Function


Function getNearestVertexAtLine(ent2 As AcadLine, ref As clsPoint) As Double()

    Dim ent2d() As Double
    Dim x1, y1, x2, y2 As Double
    
    Dim dist1, dist2 As Double
    Dim ret(0 To 1) As Double
    
    
    x1 = ent2.startPoint(0)
    y1 = ent2.startPoint(1)
    x2 = ent2.endPoint(0)
    y2 = ent2.endPoint(1)
    
    'addDonutA x1, y1
    'addDonutA x2, y2
        
    dist1 = getDistance(x1, y1, ref.x, ref.y)
    dist2 = getDistance(x2, y2, ref.x, ref.y)
    
    If dist1 > dist2 Then
        ret(0) = x2
        ret(1) = y2
    Else
        ret(0) = x1
        ret(1) = y1
    End If
    
    getNearestVertexAtLine = ret


End Function
