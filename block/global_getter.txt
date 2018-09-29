'==============================================================
'                   全局方法
'==============================================================

'==============================================================
'                 获取集群节点数量
'==============================================================
Function getClusterNodesNumber()
    getClusterNodesNumber = Sheets("基础配置").Range("nodesNum").value
End Function
'==============================================================
'                 获取方案的选择值
'==============================================================
Function getPlanSelected()
    getPlanSelected = Sheets("基础配置").Range("planOption").value
End Function
'==============================================================
'                 获取最小节点数量
'==============================================================
Function getMinNodesNum()
    getMinNodesNum = Sheets("constantValue").Range("MIN_CULSTER_NODE_NUM").value
End Function


'==============================================================
'                 获取总数据盘  数量
'==============================================================
Function getTotalDiskNum(rowIndex)

    With Sheets("磁盘配置")
        totalDiskNum = .Cells(rowIndex, .Range("totalDiskCount").Column).value
    End With
    getTotalDiskNum = totalDiskNum
    
End Function

'==============================================================
'                 统计元数据组件 安装数量
'==============================================================
Function getMetaDiskCount(rowIndex)

    With Sheets("组件部署")
      If .Cells(rowIndex, .Range("zooKeeperServer").Column).value = "Y" And .Range("zooKeeperServer").MergeArea.Columns.EntireColumn.Hidden = False Then
        zkDiskCount = 1
      End If
               
      If .Cells(rowIndex, .Range("nameNode").Column).value = "Y" And .Range("nameNode").MergeArea.Columns.EntireColumn.Hidden = False Then
        nnDiskCount = 1
      End If
    
      If .Cells(rowIndex, .Range("journalNode").Column).value = "Y" And .Range("nameNode").MergeArea.Columns.EntireColumn.Hidden = False Then
        jnDiskCount = 1
      End If
      
      If .Cells(rowIndex, .Range("KuduMaster").Column).value = "Y" And .Range("KuduMaster").MergeArea.Columns.EntireColumn.Hidden = False Then
        kmDiskCount = 1
      End If
       
      ' MN节点上默认安装MySQL，计数，like模糊匹配
      If .Cells(rowIndex, .Range("nodeType").Column).value Like "MN*" Then
        mySQLDiskCount = 1
      End If
      
    End With
    
    metaDiskCount = zkDiskCount + nnDiskCount + jnDiskCount + kmDiskCount + mySQLDiskCount
    getMetaDiskCount = metaDiskCount
    
End Function

'==============================================================
'                 获取最大元数据盘  数量
'==============================================================
Function getMaxMetaDiskNum(rowIndex)
    
    metaDiskCount = getMetaDiskCount(rowIndex)
    temp = Cells(rowIndex, Range("totalDiskCount").Column).value - Cells(rowIndex, Range("dataDiskCount").Column).value
    '取二者较小的值
    getMaxMetaDiskNum = Application.Min(metaDiskCount, temp)
    
End Function
'==============================================================
'                 获取最小元数据盘  数量
'==============================================================
Function getMinMetaDiskNum(rowIndex)
    
    
    If isMetaComponentInstalled(rowIndex) Then
      minMetaDiskNum = 1
    Else
      minMetaDiskNum = 0
    End If
    
    getMinMetaDiskNum = minMetaDiskNum
    
End Function

'==============================================================
'                 统计数据组件  数量
'==============================================================
Function getDataDiskCount(rowIndex)

    With Sheets("组件部署")
      If .Cells(rowIndex, .Range("dataNode").Column).value = "Y" And .Range("dataNode").MergeArea.Columns.EntireColumn.Hidden = False Then
        dnDiskCount = 1
      End If
      
      If .Cells(rowIndex, .Range("KafkaBroker").Column).value = "Y" And .Range("KafkaBroker").MergeArea.Columns.EntireColumn.Hidden = False Then
        kafkaDiskCount = 1
      End If
      
      If .Cells(rowIndex, .Range("EsServer").Column).value = "Y" And .Range("EsServer").MergeArea.Columns.EntireColumn.Hidden = False Then
        esDiskCount = 1
      End If
      
      If .Cells(rowIndex, .Range("NiFiMaster").Column).value = "Y" And .Range("NiFiMaster").MergeArea.Columns.EntireColumn.Hidden = False Then
        nifiDiskCount = 1
      End If
      
      If .Cells(rowIndex, .Range("KuduTServer").Column).value = "Y" And .Range("KuduTServer").MergeArea.Columns.EntireColumn.Hidden = False Then
        kuduDiskCount = 1
      End If
    
    End With
    
    dataDiskCount = dnDiskCount + kafkaDiskCount + esDiskCount + nifiDiskCount + kuduDiskCount

    getDataDiskCount = dataDiskCount
    
End Function

'==============================================================
'                 获取最大数据盘  数量
'==============================================================
Function getMaxDataDiskNum(rowIndex)

      maxDataDiskNum = Cells(rowIndex, Range("totalDiskCount").Column).value - Cells(rowIndex, Range("metaDiskCount").Column).value
      
      getMaxDataDiskNum = maxDataDiskNum
    
End Function
'==============================================================
'                 获取最小数据盘  数量
'==============================================================
Function getMinDataDiskNum(rowIndex)

    dataDiskCount = getDataDiskCount(rowIndex)
    temp = Cells(rowIndex, Range("totalDiskCount").Column).value - Cells(rowIndex, Range("metaDiskCount").Column).value
  
    If temp > dataDiskCount Then
       minDataDiskNum = dataDiskCount
    ElseIf isDataComponentInstalled(rowIndex) Then
      minDataDiskNum = 1
    Else
      minDataDiskNum = 0
    End If
    
    getMinDataDiskNum = minDataDiskNum
    
End Function

'==============================================================
'              获取controller 1所在主机名 master1
'==============================================================
Function getMaster1HostName()

    ' 获取master1所在的行
    optionValue = getPlanSelected
    Select Case optionValue
        Case plan1Value
            rowMN1 = nodesInfoStartRow
        Case plan2Value
            rowMN1 = nodesInfoStartRow + 2
        Case plan3Value
            rowMN1 = nodesInfoStartRow
    End Select
    
    With Sheets("组件部署")
    
      hostIP = .Cells(rowMN1, .Range("componentNodesIP").Column).value
      
    End With
    
    hostName = LCase(Range("clusterName")) & "-" & Replace(hostIP, ".", "-") & "." & Range("domainName")
     
    getMaster1HostName = hostName
    
End Function
'==============================================================
'                 获取ranger_admin所在主机的IP
'==============================================================
Function getRangerHostIP()
    
    nodesNum = getClusterNodesNumber
    startRow = nodesInfoStartRow
    lastRow = nodesInfoStartRow + nodesNum - 1
    For i = startRow To lastRow
      With Sheets("组件部署")
        If .Cells(i, .Range("rangerAdmin").Column) = "Y" Then
          hostIP = .Cells(i, .Range("componentNodesIP").Column).value
          Exit For
        End If
      End With
    Next
    getRangerHostIP = hostIP
    
End Function

'==============================================================
'                 获取kafka_broker所在主机的IP
'==============================================================
Function getKafkaHostIP()
    
    nodesNum = getClusterNodesNumber
    startRow = nodesInfoStartRow
    lastRow = nodesInfoStartRow + nodesNum - 1
    For i = startRow To lastRow
      With Sheets("组件部署")
        If .Cells(i, .Range("KafkaBroker").Column) = "Y" Then
          hostIP = .Cells(i, .Range("componentNodesIP").Column).value
          Exit For
        End If
      End With
    Next
    getKafkaHostIP = hostIP
    
End Function
'==============================================================
'                 获取nifi_master所在主机的IP
'==============================================================
Function getNifiHostIP()
    
    nodesNum = getClusterNodesNumber
    startRow = nodesInfoStartRow
    lastRow = nodesInfoStartRow + nodesNum - 1
    For i = startRow To lastRow
      With Sheets("组件部署")
        If .Cells(i, .Range("NiFiMaster").Column) = "Y" Then
          hostIP = .Cells(i, .Range("componentNodesIP").Column).value
          Exit For
        End If
      End With
    Next
    getNifiHostIP = hostIP
    
End Function

'==============================================================
'                 获取knox所在主机的IP
'==============================================================
Function getKnoxHostIP()
    
    nodesNum = getClusterNodesNumber
    startRow = nodesInfoStartRow
    lastRow = nodesInfoStartRow + nodesNum - 1
    For i = startRow To lastRow
      With Sheets("组件部署")
        If .Cells(i, .Range("knoxGateway").Column) = "Y" Then
          hostIP = .Cells(i, .Range("componentNodesIP").Column).value
          Exit For
        End If
      End With
    Next
    getKnoxHostIP = hostIP
    
End Function

'=======================================================================
'     获取nifimaster所在主机,拼接成nifi-ambari-ssl-config中的contentValue
'========================================================================
Function getNifiContent()
    
    nodesNum = getClusterNodesNumber
    startRow = nodesInfoStartRow
    lastRow = nodesInfoStartRow + nodesNum - 1
    cnt = 0
    nifiContent = ""
    For i = startRow To lastRow
      With Sheets("组件部署")
        If .Cells(i, .Range("NiFiMaster").Column) = "Y" Then
          hostIP = .Cells(i, .Range("componentNodesIP").Column).value
          cnt = cnt + 1
          '单引号 保证json数据合法
          val1 = "<property name=" & Chr(39) & "Node Identity " & cnt & Chr(39) & ">CN="
          nifiHostName = LCase(Range("clusterName")) & "-" & Replace(hostIP, ".", "-") & "." & Range("domainName")
          
          If i <> lastRow Then
            val2 = ", OU=NIFI</property>" & "\n"
          Else
            val2 = ", OU=NIFI</property>"
          End If
          contentValue = contentValue & val1 & nifiHostName & val2
        End If
      End With
    Next
    getNifiContent = contentValue
    
End Function
'=======================================================================
'     获取Knox所在主机名
'========================================================================
Function getKnoxHostName()
    
    nodesNum = getClusterNodesNumber
    startRow = nodesInfoStartRow
    lastRow = nodesInfoStartRow + nodesNum - 1
    For i = startRow To lastRow
      With Sheets("组件部署")
        If .Cells(i, .Range("knoxGateway").Column) = "Y" Then
          hostIP = .Cells(i, .Range("componentNodesIP").Column).value
          knoxHostName = LCase(Range("clusterName")) & "-" & Replace(hostIP, ".", "-") & "." & Range("domainName")
          Exit For
        End If
      End With
    Next
    getKnoxHostName = knoxHostName
    
End Function

'==============================================================
'                 获取三台zookeeper,拼接成zk.quorum
'==============================================================
Function getZKQuorum()
    
    nodesNum = getClusterNodesNumber
    startRow = nodesInfoStartRow
    lastRow = nodesInfoStartRow + nodesNum - 1
    cnt = 0
    zkQuorum = ""
    For i = startRow To lastRow
      With Sheets("组件部署")
        If .Cells(i, .Range("zooKeeperServer").Column) = "Y" Then
          hostIP = .Cells(i, .Range("componentNodesIP").Column).value
          cnt = cnt + 1
          zkServer = LCase(Range("clusterName")) & "-" & Replace(hostIP, ".", "-") & "." & Range("domainName") & ":2181"
          
        End If
        If cnt = 3 Then
          zkQuorum = zkQuorum & zkServer
          Exit For
        Else
          zkQuorum = zkServer & "," & zkQuorum
        End If
      End With
    Next
    getZKQuorum = zkQuorum
    
End Function

'=====================================================================
'        获取 选择 的服务名称，"服务选择"表中调用
'=====================================================================
Function getServicesSelected() As String()
    servicesSelected = ""
    servicesUnselected = ""
    Dim arr(2) As String
    
    For Each cell In Range("userServiceSelect")
        Row = cell.Row
        col = cell.Column
        If cell.value = "Y" Then                                                           'cell的地址为相对地址
            servicesSelected = servicesSelected & Sheets("服务选择").Range("servicesName").Cells(Row - 6, col).value & " "
        ElseIf cell.value = "" Then                                                           'cell的地址为相对地址
            servicesUnselected = servicesUnselected & Sheets("服务选择").Range("servicesName").Cells(Row - 6, col).value & " "
        End If
    Next
    arr(1) = Trim(servicesSelected)
    arr(2) = Trim(servicesUnselected)
    getServicesSelected = arr
End Function

'==========================================================================
'
'               获取服务依赖的服务(按行遍历)
'
'==========================================================================

Function getRequiredServices(serviceName)
    retText = ""
    
    RowCount = Range("requiredRelation").Rows.Count
    startRow = Range("requiredRelation").Row
    serviceNameColumn = Range("rowServices").Column
    
    For Each serviceNameCell In Range("columnServices")
        lServiceName = serviceNameCell.value
        
        If lServiceName = serviceName Then
            columnIndex = serviceNameCell.Column
            
            For Index = 0 To RowCount - 1
                cellText = Sheets("required_services").Cells(startRow + Index, columnIndex)
                
                If cellText = "Y" Then
                    retText = retText & Sheets("required_services").Cells(startRow + Index, serviceNameColumn) & " "
                End If
            Next
        End If
    Next
    
    getRequiredServices = Trim(retText)

End Function
'==========================================================================
'
'               获取服务依赖的服务(按列遍历)
'
'==========================================================================
Function getDependencyServices(serviceName)
    retText = ""
    
    ColumnCount = Range("requiredRelation").Columns.Count
    startColumn = Range("requiredRelation").Column
    serviceNameRow = Range("columnServices").Row
    
    For Each serviceNameCell In Range("rowServices")
        lServiceName = serviceNameCell.value
        
        If lServiceName = serviceName Then
            rowIndex = serviceNameCell.Row
            
            For Index = 0 To ColumnCount - 1
                cellText = Sheets("required_services").Cells(rowIndex, startColumn + Index)
                
                If cellText = "Y" Then
                    retText = retText & Sheets("required_services").Cells(serviceNameRow, startColumn + Index) & " "
                End If
            Next
        End If
        
    Next
    
    getDependencyServices = Trim(retText)

End Function


'=====================================================================
'        获取输出配置文件的路径
'=====================================================================
Function getConfFilePath()

  Dim str As String
  dirp = Range("confFilePath").value
  If Right(dirp, 1) = "\" Then    '除去最后一位的"\"
    str = Left(dirp, Len(dirp) - 1)
  Else
    str = dirp
  End If
  getConfFilePath = Replace(str & "\indata_config", "/", "\")
End Function
'=====================================================================
'        遍历子目录，创建路径
'=====================================================================
Sub mkdirp(dirp As String)
    Dim i As Integer
    Dim str As String
    Dim str_path As String

    If Right(dirp, 1) = "\" Then    '获取完整路径名称
        str = Left(dirp, Len(dirp) - 1)
    Else
        str = dirp
    End If
    
    On Error Resume Next
    str_path = Split(dirp, "\")(0) & "\" & Split(dirp, "\")(1)    '创建一级目录
    MkDir (str_path)
    For i = 2 To UBound(Split(dirp, "\"))  '依次创建子目录
        str_path = str_path & "\" & Split(dirp, "\")(i)
        MkDir (str_path)
    Next
End Sub
