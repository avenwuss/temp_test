
Private Sub 提取数据_Click()
Dim RowNdx As Long
Dim ColNdx As Integer
Dim StartRow As Long
Dim EndRow As Long
Dim CompanyID As Long
Dim Path As String
Dim File As String
Dim WB As Workbook
Dim CheckValue As String
Dim CellValue As Double
Dim myApp As New Application
Dim ThisMon As String
Dim StartMon As Long
Dim EndMon As Long
Dim CMon1 As Long
Dim CMon2 As Long

StartMon = 6
StartRow = 6
StartCol = 5
EndRow = ThisWorkbook.Sheets("参数").Cells(Rows.Count, "L").End(xlUp).Row
EndMon = Sheets("参数").Cells(Rows.Count, "J").End(xlUp).Row

If (EndRow >= StartRow) Then

ThisWorkbook.Sheets("全口径全年预算").Range("E6:SZ97").ClearContents

ColNdx = StartCol

For RowNdx = StartRow To EndRow
    For MonNdx = StartMon To EndMon
        Company = Sheets("参数").Cells(RowNdx, 12).Value
        ThisMon = Sheets("参数").Cells(MonNdx, 10).Value
        Sheets("全口径全年预算").Cells(7, ColNdx).Value = Company
        Sheets("全口径全年预算").Cells(9, ColNdx).Value = ThisMon
        ColNdx = ColNdx + 1
    Next MonNdx
Next RowNdx


End If


EndRow = ThisWorkbook.Sheets("全口径全年预算").Cells(Rows.Count, "D").End(xlUp).Row
EndCol = ThisWorkbook.Sheets("全口径全年预算").Range("SZ7").End(xlToLeft).Column


'Application.ScreenUpdating = False
Path = ThisWorkbook.Sheets("参数").Cells(6, 2).Value & "年度预算\"
File = Dir(Path & "*.xlsm")
Do While File <> ""
    ThisWorkbook.Sheets("全口径全年预算").Cells(1, 8).Value = File
   If (InStr(File, "采集") = 0) Then
   myApp.Visible = False
   Set WB = myApp.Workbooks.Open(Path & File)
   Company = ""
   CheckValue = WB.Sheets("全年预算").Cells(6, 5).Value
   Company = WB.Sheets("全年预算").Cells(6, 2).Value
   If (Company <> "") Then
     StartCol = 5
     For ColNdx = StartCol To EndCol
                If (Company = ThisWorkbook.Sheets("全口径全年预算").Cells(7, ColNdx).Value) Then
                   If (InStr(CheckValue, "数据正确") <> 0) Then
                   StartRow = 10
                   For RowNdx = StartRow To EndRow
                       StartMon = 0
                       EndMon = 11
                       For MonNdx = StartMon To EndMon
                       CMon1 = 5 + MonNdx
                       CMon2 = ColNdx + MonNdx
                       If (WB.Sheets("全年预算").Cells(RowNdx, CMon1).Value <> "") Then
                        CellValue = WB.Sheets("全年预算").Cells(RowNdx, CMon1).Value
                        ThisWorkbook.Sheets("全口径全年预算").Cells(RowNdx, CMon2).Value = CellValue
                        End If
                        Next MonNdx
                      Next RowNdx
                      ThisWorkbook.Sheets("全口径全年预算").Cells(6, ColNdx).Value = "数据正确，已提取"
                    Else
                      ThisWorkbook.Sheets("全口径全年预算").Cells(6, ColNdx).Value = "数据有误，未提取"
                    End If
                   Exit For
                 End If
            Next ColNdx
           End If
 WB.Close
 End If
File = Dir
Loop
Application.ScreenUpdating = True
myApp.Quit
End Sub
