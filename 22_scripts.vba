Sub Table_info()
'----------------------------------------------------------------------------------------------------------------------------
Set CallerBook = ThisWorkbook
Dim TimeStamp As String
'----------------------------------------------------------------------------------------------------------------------------
Debug.Print "Открытие соединения DSN" & Now()
Dim db As ADODB.Connection
Set db = New ADODB.Connection
db.ConnectionString = "DSN=TD_RDV"
db.Open
db.CommandTimeout = 0
'----------------------------------------------------------------------------------------------------------------------------
'Указатель на номер строки
'----------------------------------------------------------------------------------------------------------------------------
rownumber = 2

'дата обработки
SqlCode01 = "(SELECT CAST(CURRENT_TIME AS TIMESTAMP(0)) DTMWhenProcessed)"
Set dt = db.Execute(SqlCode01)
dttime = Format(dt!DTMWhenProcessed, "YYYY-MM-DD HH:MM:SS")
'---------------------------------------------------------------------------------------------------------------------------
Do While CallerBook.Worksheets("Неописанные витрины").Cells(rownumber, 1).Value <> ""
'---------------------------------------------------------------------------------------------------------------------------
PTABLE_NAME = Worksheets("Неописанные витрины").Cells(rownumber, 1).Value              'Наименование таблицы
PTABLE_COMMENT = Worksheets("Неописанные витрины").Cells(rownumber, 3).Value           'Комментарий таблицы
PTABLE_ID = Worksheets("Неописанные витрины").Cells(rownumber, 4).Value                'ID таблицы
PBDOM_ID = -1
'---------------------------------------------------------------------------------------------------------------------------

If PTABLE_ID <> "" Then
    
    SqlCode01 = "insert into PRD_DB_DMT.PLDM_TABLE(TABLE_NAME , TABLE_COMMENT , TABLE_ID , CHANGE_DTM)" & Chr(10) & Chr(13) & _
        "values (" & "'" & PTABLE_NAME & "'" & "," & "'" & PTABLE_COMMENT & "'" & "," & PTABLE_ID & "," & "'" & dttime & "'" & ")"
    db.Execute (SqlCode01)
    
    'SqlCode02 = "insert into PRD_DB_DMT.PLDM_BDOM_TABLE_LNK (TABLE_ID,BDOM_ID)" & Chr(10) & Chr(13) & _
        '"values (" & PTABLE_ID & "," & PBDOM_ID & ")"
    'db.Execute (SqlCode02)
            
End If
'----------------------------------------------------------------------------------------------------------------------------
rownumber = rownumber + 1
'----------------------------------------------------------------------------------------------------------------------------
Loop
'----------------------------------------------------------------------------------------------------------------------------
' Обновление данных витрины Absent
'----------------------------------------------------------------------------------------------------------------------------
    Range("D2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
'----------------------------------------------------------------------------------------------------------------------------
    'Sheets("Неописанные витрины").Select
    'Selection.ListObject.QueryTable.Refresh
    'Sheets("Неописанные витрины").Select
db.Close
End Sub
