Sub delete_not_exist_table_column_lnk()
'----------------------------------------------------------------------------------------------------------------------------
'Удаление неактуальных связей полей и таблиц
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
'----------------------------------------------------------------------------------------------------------------------------
Do While CallerBook.Worksheets("Уд_неакт_св_полей_и_таблиц").Cells(rownumber, 1).Value <> ""
'----------------------------------------------------------------------------------------------------------------------------
'пока нет пустой строки
'----------------------------------------------------------------------------------------------------------------------------
PTABLE_ID = Worksheets("Уд_неакт_св_полей_и_таблиц").Cells(rownumber, 1).Value                 'ID таблицы
PCOLUMN_ID = Worksheets("Уд_неакт_св_полей_и_таблиц").Cells(rownumber, 3).Value               'ID поля
'----------------------------------------------------------------------------------------------------------------------------
If PTABLE_ID <> "" Then
            
            SqlCode03 = "delete from PRD_DB_DMT.PLDM_TABLE_COLUMN_LNK WHERE TABLE_ID = " & PTABLE_ID & " AND COLUMN_ID = " & PCOLUMN_ID & ";"
            db.Execute (SqlCode03)
        
End If
'----------------------------------------------------------------------------------------------------------------------------
    rownumber = rownumber + 1
'----------------------------------------------------------------------------------------------------------------------------
Loop
'----------------------------------------------------------------------------------------------------------------------------
    Range("A2:E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
'----------------------------------------------------------------------------------------------------------------------------
    Sheets("Уд_неакт_св_полей_и_таблиц").Select
    Selection.ListObject.QueryTable.Refresh
'----------------------------------------------------------------------------------------------------------------------------
db.Close
End Sub
