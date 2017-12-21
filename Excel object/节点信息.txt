

'==============================================================
'                 方案选择变更
'==============================================================

Sub planSelectChanged()
    
    '最小节点数约束，更新
    Sheets("constantValue").updateMinNodesNum
    minNodesNum = getMinNodesNum
    nodesNum = getClusterNodesNumber
    startRow = nodesInfoStartRow
    lastRow = getLastRowIndex
    
  With Sheets("节点信息")
    lastCol = .Range("infoLastCol").Column - 1
    ' 最小节点数 > 节点数量
    If minNodesNum > nodesNum Then
      Sheets("constantValue").Range("backgroundTemplate").Copy
      .Range("A" & lastRow + 1, .Cells(startRow + minNodesNum - 1, lastCol)).PasteSpecial
    ' 当前最后一行  < 节点数量 + 起始行 - 1
    ElseIf lastRow < startRow + nodesNum - 1 Then
      Sheets("constantValue").Range("backgroundTemplate").Copy
      .Range("A" & lastRow + 1, .Cells(startRow + nodesNum - 1, lastCol)).PasteSpecial
    End If
    
  End With
  Application.CutCopyMode = False
End Sub
'==============================================================
'                 集群节点数量变更
'==============================================================
Sub nodesNumberChanged()

  With Sheets("节点信息")
    nodesNum = getClusterNodesNumber
    startRow = nodesInfoStartRow
    lastRow = getLastRowIndex
    lastCol = .Range("infoLastCol").Column - 1
    
    If lastRow < startRow + nodesNum - 1 Then
    Sheets("constantValue").Range("backgroundTemplate").Copy
    ' 从当前的最后一行，扩充
    .Range("A" & lastRow + 1, .Cells(startRow + nodesNum - 1, lastCol)).PasteSpecial
    ' 多余清除，正好相等不处理
    ElseIf lastRow > startRow + nodesNum - 1 Then
    .Range("A" & startRow + nodesNum, .Cells(lastRow, lastCol)).Clear  '多余的填充行删除
    End If
  End With
  Application.CutCopyMode = False
  
End Sub

'==============================================================
'            获取表的最后一行  Interior.ColorIndex
'==============================================================
Function getLastRowIndex()
    Set nodesRows = Range("A" & nodesInfoStartRow & ":A65535")
    For Each cell In nodesRows
      If cell.Interior.ColorIndex > 0 Then '根据是否填充颜色计数
        cnt = cnt + 1  '有填充颜色的行数计数
      End If
    Next
    getLastRowIndex = cnt + nodesInfoStartRow - 1 '最后一行=起始行+ 有填充的行数
End Function
