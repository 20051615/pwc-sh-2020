Attribute VB_Name = "Execute"
Function getDelta(adjustments_sheet, row_idx)
    net_adjust = adjustments_sheet.Cells(row_idx, "I").Value
    getDelta = net_adjust - _
        adjustments_sheet.Cells(row_idx, "J").Value - _
        adjustments_sheet.Cells(row_idx, "N").Value - _
        adjustments_sheet.Cells(row_idx, "R").Value
End Function
Sub processWorkbook(Workbook, RowIDs)

    Dim reason_G, reason_H As String
    reason_G = "1-5月二手车反馈调整"
    reason_H = "1-5月二手车反馈调整"

    'Delete rows
    Set to_delete_sheet = Workbook.Sheets("2.2二手车业务")
    lastRowIndex = to_delete_sheet.Cells(to_delete_sheet.Rows.Count, "H").End(xlUp).Row
    Dim rng As Range
    Set rng = Nothing
    
    Period = Workbook.Sheets("0.0 问题清单").Range("B11").Value

    If Period < 43982 Then startRowIndex = 22 Else startRowIndex = 26
    
    
    For i = startRowIndex To lastRowIndex
        If RowIDs.Exists(i - 7) Then
            Set cel = to_delete_sheet.Cells(i, "AI")
            If rng Is Nothing Then
                Set rng = cel
            Else
                Set rng = Union(rng, cel)
            End If
        End If
    Next i

    If Not (rng Is Nothing) Then
    
        'Find 调整序号
        Set adjustment_indices_sheet = Workbook.Sheets("1.0 调整分录check")
        idx = 13
        While adjustment_indices_sheet.Cells(idx, 3).Value <> ""
            idx = idx + 1
        Wend
        adjustment_idx = adjustment_indices_sheet.Cells(idx, 2).Value

        
        'Set to value
        Set adjustments_sheet = Workbook.Sheets("1.1 报表调整")
        If Period > 43861 Then
                Set set_to_val_range = adjustments_sheet.Range("J16:U24")
                set_to_val_range.Value = set_to_val_range.Value
                Set set_to_val_range = adjustments_sheet.Range("J26:U34")
                set_to_val_range.Value = set_to_val_range.Value
                Set set_to_val_range = adjustments_sheet.Range("J36:U44")
                set_to_val_range.Value = set_to_val_range.Value
                Set set_to_val_range = adjustments_sheet.Range("H133:H146")
                set_to_val_range.Value = set_to_val_range.Value
        Else
                Set set_to_val_range = adjustments_sheet.Range("J16:U20")
                set_to_val_range.Value = set_to_val_range.Value
                Set set_to_val_range = adjustments_sheet.Range("J22:U26")
                set_to_val_range.Value = set_to_val_range.Value
                Set set_to_val_range = adjustments_sheet.Range("J28:U32")
                set_to_val_range.Value = set_to_val_range.Value
                Set set_to_val_range = adjustments_sheet.Range("H121:H134")
                set_to_val_range.Value = set_to_val_range.Value
        
        End If
        
        Set dr_adjustments_sheet = Workbook.Worksheets("2.4衍生业务")
        
        dr_row_num = dr_adjustments_sheet.Range("J65536").End(xlUp).Row
        
        If dr_adjustments_sheet.Range("J115") <> "" Then
        
          Set set_to_val_range = dr_adjustments_sheet.Range("J115:J" & dr_row_num)
                set_to_val_range.Value = set_to_val_range.Value
        
        End If
        
        rng.EntireRow.Delete
        
        'Compute deltas
        
        If Period > 43861 Then   '非1月的情况
            For row_idx = 16 To 44
                If row_idx <> 25 And row_idx <> 35 Then
                    adjust_val = getDelta(adjustments_sheet, row_idx)
                    If adjust_val <> 0 Then
                        adjustments_sheet.Cells(row_idx, "V").Value = adjust_val
                        adjustments_sheet.Cells(row_idx, "W").Value = "7.1 非真实业务引起的调整"
                        adjustments_sheet.Cells(row_idx, "X").Value = adjustment_idx
                        adjustments_sheet.Cells(row_idx, "Y").Value = 10
                    End If
                End If
            Next row_idx
            
            row_idx = 132
            While adjustments_sheet.Cells(row_idx, "G").Value <> ""
                row_idx = row_idx + 1
            Wend
            If adjustments_sheet.Range("V25").Value <> 0 Then
                adjustments_sheet.Cells(row_idx, "G").Value = "收入类"
                adjustments_sheet.Cells(row_idx, "H").Value = adjustments_sheet.Range("V25").Value
                adjustments_sheet.Cells(row_idx, "L").Value = adjustment_idx
                adjustments_sheet.Cells(row_idx, "M").Value = 10
            End If
            If adjustments_sheet.Range("V35").Value <> 0 Then
                adjustments_sheet.Cells(row_idx + 1, "G").Value = "成本类"
                adjustments_sheet.Cells(row_idx + 1, "H").Value = adjustments_sheet.Range("V35").Value
                adjustments_sheet.Cells(row_idx + 1, "L").Value = adjustment_idx
                adjustments_sheet.Cells(row_idx + 1, "M").Value = 10
            End If
        
        Else '1月的情况
        
            For row_idx = 16 To 32
                If row_idx <> 21 And row_idx <> 27 Then
                    adjust_val = getDelta(adjustments_sheet, row_idx)
                    If adjust_val <> 0 Then
                        adjustments_sheet.Cells(row_idx, "V").Value = adjust_val
                        adjustments_sheet.Cells(row_idx, "W").Value = "7.1 非真实业务引起的调整"
                        adjustments_sheet.Cells(row_idx, "X").Value = adjustment_idx
                        adjustments_sheet.Cells(row_idx, "Y").Value = 10
                    End If
                End If
            Next row_idx
            
            row_idx = 120
            While adjustments_sheet.Cells(row_idx, "G").Value <> ""
                row_idx = row_idx + 1
            Wend
            If adjustments_sheet.Range("V21").Value <> 0 Then
                adjustments_sheet.Cells(row_idx, "G").Value = "收入类"
                adjustments_sheet.Cells(row_idx, "H").Value = adjustments_sheet.Range("V21").Value
                adjustments_sheet.Cells(row_idx, "L").Value = adjustment_idx
                adjustments_sheet.Cells(row_idx, "M").Value = 10
            End If
            If adjustments_sheet.Range("V27").Value <> 0 Then
                adjustments_sheet.Cells(row_idx + 1, "G").Value = "成本类"
                adjustments_sheet.Cells(row_idx + 1, "H").Value = adjustments_sheet.Range("V27").Value
                adjustments_sheet.Cells(row_idx + 1, "L").Value = adjustment_idx
                adjustments_sheet.Cells(row_idx + 1, "M").Value = 10
            End If
        
        
        End If
        
        
        
        '衍生业务调整
 
        For dr_idx = 92 To 111
            If dr_idx <> 95 And dr_idx <> 96 And dr_idx <> 101 And dr_idx <> 102 And dr_idx <> 106 And dr_idx <> 110 Then
            With dr_adjustments_sheet
                If .Cells(dr_idx, "I") <> 0 Then
                    dr_row_num = dr_row_num + 1
                    .Range("B" & dr_row_num).Value = .Cells(dr_idx, 1).Value
                    .Range("D" & dr_row_num).Value = .Cells(87, "I").Value
                    .Range("G" & dr_row_num).Value = .Cells(89, "I").Value
                    .Range("J" & dr_row_num).Value = .Cells(dr_idx, "I").Value
                    .Range("K" & dr_row_num).Value = "6.6 分项调整错误"
                    .Range("M" & dr_row_num).Value = adjustment_idx
                    .Range("N" & dr_row_num).Value = 10
                    .Cells(dr_idx - 52, "J").Value = .Cells(dr_idx - 52, "J").Value - .Cells(dr_idx, "I").Value
                    dr_row_num = dr_row_num + 1
                    .Range("B" & dr_row_num).Value = .Cells(dr_idx, 1).Value
                    .Range("D" & dr_row_num).Value = .Cells(87, "J").Value
                    .Range("G" & dr_row_num).Value = .Cells(89, "J").Value
                    .Range("J" & dr_row_num).Value = .Cells(dr_idx, "J").Value
                    .Range("K" & dr_row_num).Value = "6.6 分项调整错误"
                    .Range("M" & dr_row_num).Value = adjustment_idx
                    .Range("N" & dr_row_num).Value = 10
                  End If
                  
                If .Cells(dr_idx, "M") <> 0 Then
                    dr_row_num = dr_row_num + 1
                    .Range("B" & dr_row_num).Value = .Cells(dr_idx, 1).Value
                    .Range("D" & dr_row_num).Value = .Cells(87, "M").Value
                    .Range("G" & dr_row_num).Value = .Cells(89, "M").Value
                    .Range("J" & dr_row_num).Value = .Cells(dr_idx, "M").Value
                    .Range("K" & dr_row_num).Value = "6.6 分项调整错误"
                    .Range("M" & dr_row_num).Value = adjustment_idx
                    .Range("N" & dr_row_num).Value = 10
                
                End If
            End With
            End If
       
        Next
       
      
      If adjustment_indices_sheet.Cells(idx, 3).Value <> "" Then
        adjustment_indices_sheet.Cells(idx, "G").Value = reason_G
        adjustment_indices_sheet.Cells(idx, "H").Value = reason_H
           
      End If
    
    End If
End Sub
Sub getStatusCodes(Workbook, Store)
    With Workbook.Sheets("0.0 问题清单")
        Store(0) = .Range("C8").Value
        Store(1) = .Range("B11").Value
        Store(2) = .Range("C4").Value
        Store(3) = .Range("D4").Value
        Store(4) = .Range("E4").Value
        Store(5) = .Range("F4").Value
        Store(6) = .Range("G4").Value
        Store(7) = .Range("H4").Value
        Store(8) = .Range("I4").Value
        Store(9) = .Range("J4").Value
        Store(10) = .Range("K4").Value
    End With
    With Workbook.Sheets("1.1 报表调整")
        If Period > 43861 Then
            Store(11) = .Range("H15").Value
            Store(12) = .Range("H25").Value
            Store(13) = .Range("H35").Value
        Else
            Store(11) = .Range("H15").Value
            Store(12) = .Range("H21").Value
            Store(13) = .Range("H27").Value
        End If
    End With
End Sub
Sub Execute()
    MsgBox ("请将该含宏的文件置于下载好的需要修改的底稿的同文件夹" + vbCrLf + _
    "（即，与1/, 2/等月份文件夹处于同一级）")
    
    Application.ScreenUpdating = False
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = ThisWorkbook.Path
    
    statusRow = 1
    Application.ScreenUpdating = True
    Set statusCodesWB = Workbooks.Add
    With statusCodesWB.Sheets(1)
        .Cells(statusRow, 1).Value = "经销商代码"
        .Cells(statusRow, 2).Value = "报表月份"
        .Cells(statusRow, 3).Value = "指标问题清单"
        .Cells(statusRow, 4).Value = "公允性统计"
        .Cells(statusRow, 5).Value = "调整分录记录"
        .Cells(statusRow, 6).Value = "调整原因选择"
        .Cells(statusRow, 7).Value = "报表调整"
        .Cells(statusRow, 8).Value = "整车业务"
        .Cells(statusRow, 9).Value = "二手车业务"
        .Cells(statusRow, 10).Value = "整车库存"
        .Cells(statusRow, 11).Value = ""
        .Cells(statusRow, 12).Value = "二手车台次"
        .Cells(statusRow, 13).Value = "二手车收入"
        .Cells(statusRow, 14).Value = "二手车成本"
    End With
    Application.ScreenUpdating = False
    Dim statusCodes(13)
    statusRow = statusRow + 1
    
    Set validMonthFolders = CreateObject("Scripting.Dictionary")
    For folder_name = 1 To 12
        validMonthFolders.Add CStr(folder_name), 0
    Next folder_name
    
    Dim xlsmFiless(1 To 12) As Variant
    For idx = 1 To 12
        Set xlsmFiless(idx) = CreateObject("Scripting.Dictionary")
    Next idx
    
    Set RowIDs_idx_lookup = CreateObject("Scripting.Dictionary")
    RowIDs_idx = 0
    For Each monthFolder In fs.GetFolder(folder).SubFolders
        If validMonthFolders.Exists(monthFolder.Name) Then
            Set xlsmFiles = xlsmFiless(CInt(monthFolder.Name))
            For Each xlsmFile In monthFolder.Files
                ID = Split(xlsmFile.Name, "_")(0)
                xlsmFiles.Add ID, xlsmFile
                RowIDs_idx_lookup.Add xlsmFile.Name, RowIDs_idx
                RowIDs_idx = RowIDs_idx + 1
            Next xlsmFile
        End If
    Next monthFolder
    
    Dim RowIDss() As Variant
    ReDim RowIDss(RowIDs_idx - 1)
    For RowIDs_idx = 0 To UBound(RowIDss)
        Set RowIDss(RowIDs_idx) = CreateObject("Scripting.Dictionary")
    Next RowIDs_idx

    With ThisWorkbook.Sheets(1)
        lastRowIndex = .Cells(.Rows.Count, "A").End(xlUp).Row
        For row_idx = 2 To lastRowIndex
            Set xlsmFile = xlsmFiless(CInt(Format(.Cells(row_idx, 2).Value, "m")))(.Cells(row_idx, 1).Value)
            Set RowIDs = RowIDss(RowIDs_idx_lookup(xlsmFile.Name))
            RowIDs.Add .Cells(row_idx, 3).Value, 0
        Next row_idx
    End With
    
    For Each monthFolder In fs.GetFolder(folder).SubFolders
        If validMonthFolders.Exists(monthFolder.Name) Then
            Set xlsmFiles = xlsmFiless(CInt(monthFolder.Name))
            For Each xlsmFile In xlsmFiles.Items
                Set RowIDs = RowIDss(RowIDs_idx_lookup(xlsmFile.Name))
                Set Workbook = Workbooks.Open(xlsmFile)
                processWorkbook Workbook, RowIDs
                getStatusCodes Workbook, statusCodes
                Workbook.Close True
                
                Application.ScreenUpdating = True
                With statusCodesWB.Sheets(1)
                    For status_cel_idx = 0 To 13
                        .Cells(statusRow, status_cel_idx + 1).Value = statusCodes(status_cel_idx)
                    Next status_cel_idx
                End With
                Application.ScreenUpdating = False
                statusRow = statusRow + 1
            Next xlsmFile
        End If
    Next monthFolder
    
    MsgBox ("Finished!")
    Application.ScreenUpdating = True
End Sub
