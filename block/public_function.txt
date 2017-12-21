
'禁用事件 关闭自动计算
Sub stopListener()
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    '防止闪屏
    Application.ScreenUpdating = False
End Sub
'启用事件 开启自动计算
Sub startListener()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Calculate
    
    '防止闪屏
    Application.ScreenUpdating = True
End Sub

    '==================================================
    '
    '       是否安装 元数据组件
    '
    '==================================================
    
Function isMetaComponentInstalled(rowIndex)
    metaComponentInstalled = False
    
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
    
    If metaDiskCount > 0 Then
      metaComponentInstalled = True
    Else
      metaComponentInstalled = False
    End If
    
    isMetaComponentInstalled = metaComponentInstalled
    
End Function
    
    '=======================================================
    '
    '       是否安装 数据组件
    '
    '=======================================================
    
Function isDataComponentInstalled(rowIndex)
    
    dataComponentInstalled = False
    
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
    
    If dataDiskCount > 0 Then
      dataComponentInstalled = True
    Else
      dataComponentInstalled = False
    End If
   
   isDataComponentInstalled = dataComponentInstalled
   
End Function


Sub test()

   dataComponentInstall = isDataComponentInstalled(8)
   metaComponentInstall = isMetaComponentInstalled(8)
   
End Sub

