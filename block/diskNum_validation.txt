'=====================================================================
'
'                  更新"元数据盘数"，可选择范围(数据有效性验证)
'
'=====================================================================

Sub update_meta_DiskNumValidation(minMetaDiskNum, maxMetaDiskNum, rowIndex)
    
    minCount = minMetaDiskNum
    maxCount = maxMetaDiskNum
    limits = ""
    With Cells(rowIndex, Range("metaDiskCount").Column).Validation.Delete
    End With
    
    For i = minCount To maxCount
      If i <> maxCount Then
        limits = limits & i & ","
      Else
        limits = limits & i
      End If
    Next
    
    With Cells(rowIndex, Range("metaDiskCount").Column).Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=limits

    End With
End Sub
'=====================================================================
'
'                  更新"数据盘数"，可选择范围(数据有效性验证)
'
'=====================================================================

Sub update_data_DiskNumValidation(minDataDiskNum, maxDataDiskNum, rowIndex)

    minCount = minDataDiskNum
    maxCount = maxDataDiskNum
    limits = ""
    With Cells(rowIndex, Range("dataDiskCount").Column).Validation.Delete
    End With
    
    For i = minCount To maxCount
      If i <> maxCount Then
        limits = limits & i & ","
      Else
        limits = limits & i
      End If
    Next
    
    With Cells(rowIndex, Range("dataDiskCount").Column).Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=limits

    End With

'    'maxDataDiskCount < minDataDiskCount;
'    If minCount >= maxCount Then
'
'      For i = 1 To maxCount
'        If i <> maxCount Then
'          limits = limits & i & ","
'        Else
'          limits = limits & i
'        End If
'      Next
'      '===================组件盘数限制为1===========================
'      With Cells(rowIndex, Range("dnDiskCount").Column).Validation
'           .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:=1
'      End With
'      With Cells(rowIndex, Range("kafkaDiskCount").Column).Validation
'           .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:=1
'      End With
'      With Cells(rowIndex, Range("esDiskCount").Column).Validation
'           .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:=1
'      End With
'      With Cells(rowIndex, Range("nifiDiskCount").Column).Validation
'           .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:=1
'      End With
'      With Cells(rowIndex, Range("kuduDiskCount").Column).Validation
'           .Delete
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:=1
'      End With
'      '======================================================================
'    Else
'      For i = minCount To maxCount
'        If i <> maxCount Then
'          limits = limits & i & ","
'        Else
'          limits = limits & i
'        End If
'      Next
'
'      With Cells(rowIndex, Range("dataDiskCount").Column).Validation
'        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'        xlBetween, Formula1:=limits
'
'      End With
'
'    End If
    
End Sub
'=====================================================================
'
'                  更新"datanode数据盘数"，可选择范围(数据有效性验证)
'
'=====================================================================

Sub update_dn_DiskNumValidation(rowIndex)
    maxCount = Cells(rowIndex, Range("dataDiskCount").Column) _
             - Cells(rowIndex, Range("kafkaDiskCount").Column) _
             - Cells(rowIndex, Range("esDiskCount").Column) _
             - Cells(rowIndex, Range("nifiDiskCount").Column) _
             - Cells(rowIndex, Range("kuduDiskCount").Column)
    With Sheets("组件部署")
      If .Cells(rowIndex, .Range("dataNode").Column).value = "Y" And .Range("dataNode").MergeArea.Columns.EntireColumn.Hidden = False Then
        dnDiskCount = 1
      Else
        dnDiskCount = 0
      End If
    End With
    limits = ""
    With Cells(rowIndex, Range("dnDiskCount").Column).Validation.Delete
    End With
    For i = dnDiskCount To maxCount
      If i <> maxCount Then
         limits = limits & i & ","
      Else
         limits = limits & i
      End If
    Next
    
    With Cells(rowIndex, Range("dnDiskCount").Column).Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=limits

    End With
End Sub
'=====================================================================
'
'                  更新"kafka数据盘数"，可选择范围(数据有效性验证)
'
'=====================================================================
Sub update_kafka_DiskNumValidation(rowIndex)
    maxCount = Cells(rowIndex, Range("dataDiskCount").Column) _
             - Cells(rowIndex, Range("dnDiskCount").Column) _
             - Cells(rowIndex, Range("esDiskCount").Column) _
             - Cells(rowIndex, Range("nifiDiskCount").Column) _
             - Cells(rowIndex, Range("kuduDiskCount").Column)
    With Sheets("组件部署")
      If .Cells(rowIndex, .Range("KafkaBroker").Column).value = "Y" And .Range("KafkaBroker").MergeArea.Columns.EntireColumn.Hidden = False Then
        kafkaDiskCount = 1
      Else
        kafkaDiskCount = 0
      End If
    End With
    limits = ""
    With Cells(rowIndex, Range("kafkaDiskCount").Column).Validation.Delete
    End With
    For i = kafkaDiskCount To maxCount
      If i <> maxCount Then
         limits = limits & i & ","
      Else
         limits = limits & i
      End If
    Next
    
    With Cells(rowIndex, Range("kafkaDiskCount").Column).Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=limits

    End With
End Sub
'=====================================================================
'
'                  更新"elasticsearch数据盘数"，可选择范围(数据有效性验证)
'
'=====================================================================
Sub update_es_DiskNumValidation(rowIndex)
    maxCount = Cells(rowIndex, Range("dataDiskCount").Column) _
             - Cells(rowIndex, Range("kafkaDiskCount").Column) _
             - Cells(rowIndex, Range("dnDiskCount").Column) _
             - Cells(rowIndex, Range("nifiDiskCount").Column) _
             - Cells(rowIndex, Range("kuduDiskCount").Column)
    With Sheets("组件部署")
      If .Cells(rowIndex, .Range("EsServer").Column).value = "Y" And .Range("EsServer").MergeArea.Columns.EntireColumn.Hidden = False Then
        esDiskCount = 1
      Else
        esDiskCount = 0
      End If
    End With
    limits = ""
    With Cells(rowIndex, Range("esDiskCount").Column).Validation.Delete
    End With
    For i = esDiskCount To maxCount
      If i <> maxCount Then
         limits = limits & i & ","
      Else
         limits = limits & i
      End If
    Next
    
    With Cells(rowIndex, Range("esDiskCount").Column).Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=limits

    End With
End Sub
'=====================================================================
'
'                  更新"nifi数据盘数"，可选择范围(数据有效性验证)
'
'=====================================================================
Sub update_nifi_DiskNumValidation(rowIndex)
    maxCount = Cells(rowIndex, Range("dataDiskCount").Column) _
             - Cells(rowIndex, Range("kafkaDiskCount").Column) _
             - Cells(rowIndex, Range("esDiskCount").Column) _
             - Cells(rowIndex, Range("dnDiskCount").Column) _
             - Cells(rowIndex, Range("kuduDiskCount").Column)
    With Sheets("组件部署")
      If .Cells(rowIndex, .Range("NiFiMaster").Column).value = "Y" And .Range("NiFiMaster").MergeArea.Columns.EntireColumn.Hidden = False Then
        nifiDiskCount = 1
      Else
        nifiDiskCount = 0
      End If
    End With
    limits = ""
    With Cells(rowIndex, Range("nifiDiskCount").Column).Validation.Delete
    End With
    For i = nifiDiskCount To maxCount
      If i <> maxCount Then
         limits = limits & i & ","
      Else
         limits = limits & i
      End If
    Next
    
    With Cells(rowIndex, Range("nifiDiskCount").Column).Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=limits

    End With
End Sub
'=====================================================================
'
'                  更新"kudu数据盘数"，可选择范围(数据有效性验证)
'
'=====================================================================
Sub update_kudu_DiskNumValidation(rowIndex)
    maxCount = Cells(rowIndex, Range("dataDiskCount").Column) _
             - Cells(rowIndex, Range("kafkaDiskCount").Column) _
             - Cells(rowIndex, Range("esDiskCount").Column) _
             - Cells(rowIndex, Range("nifiDiskCount").Column) _
             - Cells(rowIndex, Range("dnDiskCount").Column)
    With Sheets("组件部署")
      If .Cells(rowIndex, .Range("KuduTServer").Column).value = "Y" And .Range("KuduTServer").MergeArea.Columns.EntireColumn.Hidden = False Then
        kuduDiskCount = 1
      Else
        kuduDiskCount = 0
      End If
    End With
    limits = ""
    With Cells(rowIndex, Range("kuduDiskCount").Column).Validation.Delete
    End With
    For i = kuduDiskCount To maxCount
      If i <> maxCount Then
         limits = limits & i & ","
      Else
         limits = limits & i
      End If
    Next
    
    With Cells(rowIndex, Range("kuduDiskCount").Column).Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=limits

    End With
End Sub


Sub test()
'
     update_data_DiskNumValidation (8)
     update_dn_DiskNumValidation (8)
     update_kafka_DiskNumValidation (8)
     update_es_DiskNumValidation (8)
     update_nifi_DiskNumValidation (8)
     update_kudu_DiskNumValidation (8)
End Sub
Sub validate(rowIndex)
    With Cells(rowIndex, Range("dataDiskCount").Column).FormatConditions.Add(xlCellValue, xlBetween, "6", "8")
    End With
End Sub
