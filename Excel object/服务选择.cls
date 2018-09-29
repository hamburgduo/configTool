


Private Sub Worksheet_Change(ByVal target As Range)
 Set operateRange = Intersect(target, Range("userServiceSelect"))
 If Not operateRange Is Nothing Then
   Application.EnableEvents = False

   serviceSelectChanged (operateRange)
   
   Application.EnableEvents = True
 End If
   
End Sub

'=====================================================================
'
'                  "服务选择"表改变，更新"组件部署"表
'
'=====================================================================
Sub serviceSelectChanged(ByVal target As Range)

    For Each cell In target
        serviceSelectChangedByRow (cell)
    Next

    Dim serviceSelected
    serviceSelected = getServicesSelected()

    For Each serviceName In Split(serviceSelected(1), " ")
    With Sheets("组件部署")                             '完全匹配LookAt:=xlWhole   以解决Ranger与Range_KMS的问题
        Set serviceRange = .Range("deployServicesName").Find(serviceName, lookat:=xlWhole)
        If Not serviceRange Is Nothing Then
            serviceRange.MergeArea.Columns.EntireColumn.Hidden = False
        End If
    End With
    Next

    For Each serviceName In Split(serviceSelected(2), " ")
    With Sheets("组件部署")
        Set serviceRange = .Range("deployServicesName").Find(serviceName, lookat:=xlWhole)
        If Not serviceRange Is Nothing Then
            serviceRange.MergeArea.Columns.EntireColumn.Hidden = True
        End If
    End With
    Next
    
End Sub

'=====================================================================
'        根据依赖关系，自动填充服务选择
'=====================================================================
Sub serviceSelectChangedByRow(ByVal cell As Range)

    Row = cell.Row
    col = cell.Column                  'cell的地址为相对地址
    serviceName = Range("servicesName").Cells(Row - 6, col).value

    If cell.value = "Y" Then
        dependList = getRequiredServices(serviceName) '获取当前选择服务依赖的服务
        serviceList = Split(dependList, " ")
        For Each service In serviceList
            Row = Range("servicesName").Find(service).Row
            If Cells(Row, col).value = "" Then
                Cells(Row, col).value = "Y" '依赖的服务自动填写Y
            End If
        Next
    Else
        dependingList = getDependencyServices(serviceName) '获取 依赖于 当前选择服务的 服务

        serviceList = Split(dependingList, " ")
        For Each service In serviceList
            Row = Range("servicesName").Find(service).Row
            If Cells(Row, col).value = "Y" Then '依赖的服务自动清空（没选择服务）
                Cells(Row, col).value = ""
            End If
        Next
    End If
End Sub


Sub test()
 ser = getDependencyServices("Ranger_KMS")
 ser2 = getRequiredServices("Ranger_KMS")
 ser3 = getServicesSelected
 Set serR = Sheets("组件部署").Range("deployServicesName").Find("Ranger", lookat:=xlWhole)
 Set serR1 = Sheets("组件部署").Range("deployServicesName").Find("Ranger_KMS", lookat:=xlWhole)
 
 For Each serviceName In Split(ser3(1), " ")
    With Sheets("组件部署")
        Set serviceRange = .Range("deployServicesName").Find(serviceName)
        If Not serviceRange Is Nothing Then
            serviceRange.MergeArea.Columns.EntireColumn.Hidden = False
        End If
    End With
    Next

    For Each serviceName In Split(ser3(2), " ")
    With Sheets("组件部署")
        Set serviceRange = .Range("deployServicesName").Find(serviceName)
        If Not serviceRange Is Nothing Then
            serviceRange.MergeArea.Columns.EntireColumn.Hidden = True
        End If
    End With
    Next
End Sub
