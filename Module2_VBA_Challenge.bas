Attribute VB_Name = "Module2"
Sub loop_sheets_assign()

    'declare current as a worksheet object variable
    'Dim thissheet As Worksheet
    Dim WS_Count As Integer
    
    'loop through all worksheets in the active workbook
    WS_Count = ActiveWorkbook.Worksheets.Count
    For thissheet = 1 To WS_Count
    
    'insert code to occur n each sheet
    'create var for stock name and volume
    Dim ticker As String
    Dim volume As LongLong
    Dim numrows As Long
    'Dim y_open As Double
    'Dim y_close As Double
    Dim percentchange As Double
    
    
    
    
    'create summary table
    Dim Summary_Table_row As Integer
    Summary_Table_row = 2
    volume = 0
    numrows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    Range("L1").Value = "year open"
    Range("M1").Value = "year close"
    Range("N1").Value = "volume"
    Range("K1").Value = "ticker"
    
    
    'create loop through the sheets
    For I = 2 To numrows
    
    
    
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
            ticker = Cells(I, 1).Value
            'volume = volume + Cells(I, 7).Value
            Range("K" & Summary_Table_row).Value = ticker
            Range("N" & Summary_Table_row).Value = volume
            Range("L" & Summary_Table_row).Value = y_open
            Range("M" & Summary_Table_row).Value = y_close
            Summary_Table_row = Summary_Table_row + 1
            volume = 0
            y_close = 0
        Else
            volume = volume + Cells(I, 7).Value
            y_open = Cells(I, 3).Value
            y_close = Cells(I - 1, 6).Value
            'percentchange = (y_close - y_open) / y_open
            
        End If
        Range("L" & Summary_Table_row).Value = Format(y_open, "Currency")
        Range("M" & Summary_Table_row).Value = Format(y_close, "Currency")
    Next I
    
    'display the current worksheet name in message box
        MsgBox ActiveWorkbook.Worksheets(thissheet).Name
    Next thissheet
End Sub
