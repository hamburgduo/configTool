Sub updateMinNodesNum()
    startListener
    
    ' 更新最小节点数
    optionValue = getPlanSelected
    Select Case optionValue
        Case plan1Value
            Sheets("constantValue").Range("MIN_CULSTER_NODE_NUM").value = 6
        Case plan2Value
            Sheets("constantValue").Range("MIN_CULSTER_NODE_NUM").value = 8
        Case plan3Value
            Sheets("constantValue").Range("MIN_CULSTER_NODE_NUM").value = 3
        Case Else
            ' 错误情况下, 使用套餐3
            Sheets("constantValue").Range("MIN_CULSTER_NODE_NUM").value = 3
    End Select
    
    stopListener
End Sub
