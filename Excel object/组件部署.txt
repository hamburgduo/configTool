




Sub planSelectChanged()
     
    nodesNum = getClusterNodesNumber
    optionValue = getPlanSelected
    
    ' 1 清除Topo界面所有部署数据, 不能清理IP数据
    With Sheets("组件部署")
    
        curTypeLastRow = getTypeLastRowIndex
        lastCol = .Range("lastCol").Column - 1
        startRow = nodesInfoStartRow
        .Range("B" & startRow, .Cells(curTypeLastRow, lastCol)).Clear
        
        Select Case optionValue
            Case plan1Value
                'IP列单独处理
                curIpLastRow = getIpLastRowIndex
                If curIpLastRow < startRow + nodesNum - 1 Then
                  Sheets("constantValue").Range("ipTemplate").Copy
                  .Range("A" & (curIpLastRow + 1) & ":A" & (startRow + nodesNum - 1)).PasteSpecial
                End If
                
                Sheets("option1").Range("option1NodeType").Copy
                .Range("B" & startRow).PasteSpecial
                
                ' 设置节点Topo, 从option1ComponentTopo拷贝
                Sheets("option1").Range("option1ComponentTopo").Copy
                .Range("C" & startRow).PasteSpecial
                '7 + 6
                If startRow + nodesNum > 13 Then
                  Sheets("constantValue").Range("dataNodeTemplate").Copy
                  .Range("A" & startRow + 6, .Cells(startRow + nodesNum - 1, lastCol)).PasteSpecial
                End If
                
            Case plan2Value
                'IP列单独处理
                curIpLastRow = getIpLastRowIndex
                If curIpLastRow < startRow + nodesNum - 1 Then
                  Sheets("constantValue").Range("ipTemplate").Copy
                  .Range("A" & (curIpLastRow + 1) & ":A" & (startRow + nodesNum - 1)).PasteSpecial
                End If
                  
                Sheets("option2").Range("option2NodeType").Copy
                .Range("B" & startRow).PasteSpecial

                ' 设置节点Topo, 从option2ComponentTopo拷贝
                Sheets("option2").Range("option2ComponentTopo").Copy
                .Range("C" & startRow).PasteSpecial
                '7 + 8
                If startRow + nodesNum > 15 Then
                  Sheets("constantValue").Range("dataNodeTemplate").Copy
                  .Range("A" & startRow + 8, .Cells(startRow + nodesNum - 1, lastCol)).PasteSpecial
                End If
                
            Case plan3Value

                Sheets("option3").Range("option3NodeType").Copy
                .Range("B" & startRow).PasteSpecial
                
                ' 设置节点Topo, 从option3ComponentTopo拷贝
                Sheets("option3").Range("option3ComponentTopo").Copy
                .Range("C" & startRow).PasteSpecial
                '7 + 3
                If startRow + nodesNum > 10 Then
                  Sheets("constantValue").Range("dataNodeTemplate").Copy
                  .Range("A" & startRow + 3, .Cells(startRow + nodesNum - 1, lastCol)).PasteSpecial
                End If
                
            Case Else
        End Select
        
    End With
    Application.CutCopyMode = False
    
End Sub


Public Sub nodesNumberChanged()
    startRow = nodesInfoStartRow
    curIpLastRow = getIpLastRowIndex
    nodesNum = getClusterNodesNumber 'getClusterNodesNumber
    
    If curIpLastRow > startRow + nodesNum - 1 Then
        Range("A" & startRow + nodesNum, "A" & curIpLastRow).Rows.EntireRow.Delete
    ElseIf curIpLastRow < startRow + nodesNum - 1 Then
      With Sheets("组件部署")
        lastCol = .Range("lastCol").Column - 1
        Sheets("constantValue").Range("dataNodeTemplate").Copy
        .Range("A" & curIpLastRow + 1, .Cells(startRow + nodesNum - 1, lastCol)).PasteSpecial
      End With
    End If
    Application.CutCopyMode = False
End Sub



Function getIpLastRowIndex()
    getIpLastRowIndex = Range("A" & nodesInfoStartRow & ":A65535").Find("", lookat:=xlWhole).Row - 1
End Function

Function getTypeLastRowIndex()
    getTypeLastRowIndex = Range("B" & nodesInfoStartRow & ":B65535").Find("", lookat:=xlWhole).Row - 1
End Function
