Private Sub Worksheet_Change(ByVal target As Range)

    For Each cell In target
        ' 如果是套餐变更
        If cell.Column = Range("planOption").Column And cell.Row = Range("planOption").Row Then
           stopListener
           
           planSelectChanged
           
           startListener

        ' 如果是节点数变更
        ElseIf cell.Column = Range("nodesNum").Column And cell.Row = Range("nodesNum").Row Then
            ' New Code
            stopListener
            
            nodesNumChanged
            
            startListener
        
        ' 如果是ROOT密码是否一致变更
        ElseIf cell.Column = Range("samePassWordorNot").Column And cell.Row = Range("samePassWordorNot").Row Then
            stopListener
            
            samePassWordChanged
            
            startListener

        End If
    Next
    
    Application.CutCopyMode = False
End Sub

'=====================================================================
'
'            "基础配置"表,方案选择改变，更新"组件部署"表
'
'=====================================================================

Private Sub planSelectChanged()
     
    '最小节点数约束，更新
    Sheets("constantValue").updateMinNodesNum
    minNodesNum = getMinNodesNum
    nodesNum = getClusterNodesNumber
    If nodesNum < minNodesNum Then
      MsgBox ("选择该方案集群节点的数量不能少于：" & minNodesNum)
      Range("nodesNum").value = minNodesNum
    End If
    
    Sheets("节点信息").planSelectChanged
    
    Sheets("组件部署").planSelectChanged
    
    
End Sub

Private Sub nodesNumChanged()
    '最小节点数约束，更新
    Sheets("constantValue").updateMinNodesNum
    minNodesNum = getMinNodesNum
    nodesNum = getClusterNodesNumber
    If nodesNum < minNodesNum Then
        MsgBox ("选择该方案集群节点的数量不能少于：" & minNodesNum)
        Range("nodesNum").value = minNodesNum
    End If
      
     Sheets("节点信息").nodesNumberChanged
     
     Sheets("组件部署").nodesNumberChanged
     

End Sub

Sub samePassWordChanged()

    sameOrNot = Range("samePassWordorNot").value
    nodesNum = getClusterNodesNumber
    
    With Sheets("节点信息")
    passWordCol = .Range("infoRootPassWord").Column
      If sameOrNot = "Y" Then
        Sheets("constantValue").Range("rootPassWordTemplate").Copy
        For i = 0 To nodesNum - 1
          .Cells(nodesInfoStartRow + i, passWordCol).PasteSpecial
        Next
      ElseIf sameOrNot = "N" Then
        For i = 0 To nodesNum - 1
          .Cells(nodesInfoStartRow + i, passWordCol).value = ""
        Next
      End If
    
    End With
    
End Sub

Sub test()

 passWordCol = Sheets("节点信息").Range("infoRootPassWord").Column
 
End Sub

