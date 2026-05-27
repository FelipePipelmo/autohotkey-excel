Insert::
    Clipboard := ""
    Send, ^c
    ClipWait, 1
    if ErrorLevel
        return

    try {
        xl := ComObjActive("Excel.Application")
    } catch {
        MsgBox, O Excel não está aberto.
        return
    }

    sheet := xl.ActiveWorkbook.ActiveSheet

    lastRow := sheet.Cells(sheet.Rows.Count, 3).End(-4162).Row
    nextRow := lastRow + 1

    sheet.Cells(nextRow, 3).Value := Clipboard
return
