Sub Main_Monthly_Site_Addition()
'This module is the main module for monthly site addition will call all related subprocidue/function execution of target tasks.

Last_Month_RFS_ADDL_Col
Rearrange_column_BTS_Database
Map_from_BTS_Database
Modify_Last_Month_Before_Append
Append_To_BTS_RFS_Data
Addition_Col_BTS_RFS_Data
Pivot_Table_BTS_RFS_Data
Line_Chart
bar_graph

End Sub

Sub Last_Month_RFS_ADDL_Col()
'This Subprocedure add columns in last_month_RFS worksheets
Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\Last_Month_RFS_Data.xlsx"
Workbooks("Last_Month_RFS_Data.xlsx").Sheets("Last_Month_RFS").Activate

For i = 1 To 10
    Range("F:F").Insert
Next i

Range("F1") = "SITE_NAME"
Range(" G1") = "FACILITY_ID"
Range(" H1") = "SITE_ADDRESS"
Range("I1") = "PIN_CODE"
Range("J1") = "TOWN_NAME"
Range("K1") = "LATITUDE"
Range("L1") = "LONGITUDE"
Range("M1") = "SOLUTION_TYPE"
Range("N1") = "PLANNED_DATE"
Range("O1") = "RFS Done"
Range("Q1") = "RFS Month No"
Range("R1") = "RFS Quart No"
Range("S1") = "RFS Year No"
Range("T1") = "RFS Month Name"
Range("U1") = "RFS Quarter Name"
Range("V1") = "RFS FY"


Workbooks("Last_Month_RFS_Data.xlsx").Save
Workbooks("Last_Month_RFS_Data.xlsx").Close
End Sub
Sub Rearrange_column_BTS_Database()
'This module rearrange columns of BTS_Database and make ready this date for lookup
Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\Last_Month_RFS_Data.xlsx"
Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\BTS_Database.xlsx"
Sheets("BTS_Database").Copy After:=Workbooks("Last_Month_RFS_Data.xlsx").Sheets("Last_Month_RFS")
Workbooks("Last_Month_RFS_Data.xlsx").Sheets("BTS_Database").Activate


Dim i As Integer
i = 1
Do While Cells(1, i) <> ""
    If Cells(1, i).Value = "SL NO" Or Cells(1, i).Value = "TEL_CIRCLE" Or Cells(1, i).Value = "OP_CIRCLE_CD" Or Cells(1, i).Value = "OP_STATE_CD" Or Cells(1, i).Value = "SITE_ID" Or Cells(1, i).Value = "SITE_NAME" Or Cells(1, i).Value = "FACILITY_ID" Or Cells(1, i).Value = "SITE_ADDRESS" Or Cells(1, i).Value = "PIN_CODE" Or Cells(1, i).Value = "TOWN_NAME" Or Cells(1, i).Value = "LATITUDE" Or Cells(1, i).Value = "LONGITUDE" Or Cells(1, i).Value = "SOLUTION_TYPE" Or Cells(1, i).Value = "PLANNED_DATE" Or Cells(1, i).Value = "RFS Done" Then
    i = i + 1
    Else
        Cells(1, i).EntireColumn.Delete
        i = i - 1
    End If
Loop


Workbooks("Last_Month_RFS_Data.xlsx").Save
Workbooks("Last_Month_RFS_Data.xlsx").Close
Workbooks("BTS_Database.xlsx").Close
End Sub

Sub Map_from_BTS_Database()
'This Subprocedure map in last month RFS from BTS Database

Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\Last_Month_RFS_Data.xlsx"
Workbooks("Last_Month_RFS_Data.xlsx").Sheets("Last_Month_RFS").Activate

Dim i, j, k, x As Integer
i = 2
k = 1
j = 6

Do While Cells(1, j).Value <> "RFS Done"
    i = 2
    x = k + 1
    Do While Cells(i, 1).Value <> ""
        Cells(i, j).Value = Application.WorksheetFunction.VLookup(Cells(i, j).Offset(0, -k), Sheets("BTS_Database").Range("E1:N11528"), x, 0)
        i = i + 1
    Loop
    k = k + 1
    j = j + 1
Loop



Workbooks("Last_Month_RFS_Data.xlsx").Sheets("BTS_Database").Delete
Workbooks("Last_Month_RFS_Data.xlsx").Save
Workbooks("Last_Month_RFS_Data.xlsx").Close

End Sub
Sub Modify_Last_Month_Before_Append()
'This Subprocedure update pending fields in Last_Month_RFS Data
Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\Last_Month_RFS_Data.xlsx"
Workbooks("Last_Month_RFS_Data.xlsx").Sheets("Last_Month_RFS").Activate

Dim i As Integer
i = 2
Do While Cells(i, 1).Value <> ""
    Cells(i, 15).Value = "Yes"
    Cells(i, 17).FormulaR1C1 = "=MONTH(RC[-1])"
    Cells(i, 18).FormulaR1C1 = "=IF(RC[-1]<4,4,IF(RC[-1]<7,1,IF(RC[-1]<10,2,3)))"
    Cells(i, 19).FormulaR1C1 = "=Year(RC[-3])"
    Cells(i, 20).FormulaR1C1 = "=TEXT(RC[-4],""MMM"")"
    Cells(i, 21).FormulaR1C1 = "=IF(RC[-4]<4,""Q-4"",IF(RC[-4]<7,""Q-1"",IF(RC[-4]<10,""Q-2"",""Q-3"")))"
    Cells(i, 22).FormulaR1C1 = "=IF(RC[-4]=4,CONCATENATE(""FY "",(RC[-3]-1),""-"",RC[-3]),CONCATENATE(""FY "",RC[-3],""-"",(RC[-3]+1)))"
    i = i + 1
Loop

Range(Cells(2, 17), Cells(i - 1, 22)).Copy
Range(Cells(2, 17), Cells(i - 1, 22)).PasteSpecial xlPasteValues
Workbooks("Last_Month_RFS_Data.xlsx").Sheets("Last_Month_RFS").Application.CutCopyMode = False

Workbooks("Last_Month_RFS_Data.xlsx").Save
Workbooks("Last_Month_RFS_Data.xlsx").Close

End Sub
Sub Append_To_BTS_RFS_Data()
'This Subprocedure append last month's RFS data to Main RFS data

Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\Last_Month_RFS_Data.xlsx"
Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\BTS_RFS_Dashboard.xlsx"

Workbooks("Last_Month_RFS_Data.xlsx").Sheets("Last_Month_RFS").Activate
Dim i, j As Integer
i = 2
j = 1
Do While Cells(i, 1).Value <> ""
    i = i + 1
Loop
Do While Cells(1, j).Value <> ""
    j = j + 1
Loop
Range(Cells(2, 1), Cells(i - 1, j - 1)).Copy


Workbooks("BTS_RFS_Dashboard.xlsx").Sheets("BTS_RFS_Data").Activate
Dim k As Integer
k = 2
Do While Cells(k, 1).Value <> ""
    k = k + 1
Loop

Range(Cells(k, 1), Cells(k + i - 1, j - 1)).PasteSpecial

Workbooks("Last_Month_RFS_Data.xlsx").Sheets("Last_Month_RFS").Application.CutCopyMode = False
Workbooks("BTS_RFS_Dashboard.xlsx").Sheets("BTS_RFS_Data").Application.CutCopyMode = False


k = 2
Do While Cells(k, 1).Value <> ""
    Cells(k, 1).Value = k - 1
    k = k + 1
Loop

Workbooks("Last_Month_RFS_Data.xlsx").Save
Workbooks("Last_Month_RFS_Data.xlsx").Close
Workbooks("BTS_RFS_Dashboard.xlsx").Save
Workbooks("BTS_RFS_Dashboard.xlsx").Close

End Sub
Sub Addition_Col_BTS_RFS_Data()
Dim col_name As String
Dim col_num As Integer
Dim file_name As String
Dim Sheet_name As String

col_name = InputBox("Type the column name for insert")
col_num = InputBox("Type the column position for insert")
file_name = InputBox("Type the file name")
Sheet_name = InputBox("Type the worksheet name")

Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\" & file_name
Sheets(Sheet_name).Activate
Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\BTS_RFS_Dashboard.xlsx"
Workbooks("BTS_RFS_Dashboard.xlsx").Sheets("BTS_RFS_Data").Activate

Cells(1, col_num).EntireColumn.Insert
Cells(1, col_num).Value = col_name

Dim i As Integer
i = 2
Do While Cells(i, 1).Value <> ""
    Cells(i, col_num).Value = Application.WorksheetFunction.VLookup(Cells(i, col_num).Offset(0, 1), Workbooks(file_name).Sheets(Sheet_name).Range("B1:C7656"), 2, 0)
    i = i + 1
Loop

Workbooks("BTS_RFS_Dashboard.xlsx").Save
Workbooks("BTS_RFS_Dashboard.xlsx").Close

End Sub
Sub Pivot_Table_BTS_RFS_Data()

Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\BTS_RFS_Dashboard.xlsx"
Workbooks("BTS_RFS_Dashboard.xlsx").Sheets("BTS_RFS_Data").Activate


Dim PTable As PivotTable
Dim PCache As PivotCache
Dim PRange As Range
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim LR As Long
Dim LC As Long

On Error Resume Next
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Worksheets("Pivot Sheet").Delete 'This will delete the exisiting pivot table worksheet
Worksheets.Add After:=ActiveSheet ' This will add new worksheet
ActiveSheet.Name = "Pivot Sheet" ' This will rename the worksheet as "Pivot Sheet"
On Error GoTo 0

Set PSheet = Workbooks("BTS_RFS_Dashboard.xlsx").Worksheets("Pivot Sheet")
Set DSheet = Workbooks("BTS_RFS_Dashboard.xlsx").Worksheets("BTS_RFS_Data")

'Find Last used row and column in data sheet
LR = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LC = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column

'Set the pivot table data range
Set PRange = DSheet.Cells(1, 1).Resize(LR, LC)

'Set pivot cahe
Set PCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=PRange)

'Create blank pivot table
Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), TableName:="BTS_RFS_Report")

'Insert country to Row Filed
With PSheet.PivotTables("BTS_RFS_Report").PivotFields("RFS FY")
.Orientation = xlRowField
.Position = 1
End With

'Insert Product to Row Filed & position 2
'With PSheet.PivotTables("Sales_Report").PivotFields("Product")
'.Orientation = xlRowField
'.Position = 2
'End With

''Insert Segment to Column Filed & position 1
'With PSheet.PivotTables("Sales_Report").PivotFields("Segment")
'.Orientation = xlColumnField
'.Position = 1
'End With

'Insert Sales column to the data field
With PSheet.PivotTables("BTS_RFS_Report").PivotFields("SITE_ID")
.Orientation = xlDataField
.Position = 1
End With

'Format Pivot Table
PSheet.PivotTables("BTS_RFS_Report").ShowTableStyleRowStripes = True
PSheet.PivotTables("BTS_RFS_Report").TableStyle2 = "PivotStyleMedium14"

'Show in Tabular form
PSheet.PivotTables("BTS_RFS_Report").RowAxisLayout xlTabularRow

Application.DisplayAlerts = True
Application.ScreenUpdating = True

ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("BTS_RFS_Report"), _
        "State_Name").Slicers.Add ActiveSheet, , "State_Name", "State_Name", 100, _
        400, 150, 200
ActiveSheet.Shapes.Range(Array("State_Name")).Select
ActiveSheet.Shapes("State_Name").IncrementLeft 193.5
ActiveSheet.Shapes("State_Name").IncrementTop -185.75
    
Workbooks("BTS_RFS_Dashboard.xlsx").Save
Workbooks("BTS_RFS_Dashboard.xlsx").Close

End Sub

Sub Line_Chart()
Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\BTS_RFS_Dashboard.xlsx"
Workbooks("BTS_RFS_Dashboard.xlsx").Worksheets("Pivot Sheet").Activate


ActiveSheet.Shapes.AddChart.Select
ActiveSheet.Shapes(1).Top = 8
ActiveSheet.Shapes(1).Left = 2
ActiveChart.ChartType = xlLineMarkers

ActiveChart.PlotArea.Select
ActiveChart.HasTitle = True
ActiveChart.ChartTitle.Text = "BTS RFS Report Finacial year wise"

Workbooks("BTS_RFS_Dashboard.xlsx").Save
Workbooks("BTS_RFS_Dashboard.xlsx").Close

End Sub

Sub bar_graph()

Workbooks.Open Filename:="D:\Training\Self_Paced_Project\Excel VBA\Telecom_Wireline_Dashboard\Monthly Site addition tracker & Dashboard\Addition of Last Month's RFS Data\BTS_RFS_Dashboard.xlsx"
Workbooks("BTS_RFS_Dashboard.xlsx").Worksheets("Pivot Sheet").Activate

Worksheets("Pivot Sheet").Range(Cells(1, 1), Cells(8, 2)).Select
ActiveSheet.Shapes.AddChart2(216, xlBarClustered).Select
ActiveChart.HasTitle = True
ActiveChart.ChartTitle.Text = "BTS RFS Report Finacial year wise"



Workbooks("BTS_RFS_Dashboard.xlsx").Save
Workbooks("BTS_RFS_Dashboard.xlsx").Close
End Sub
