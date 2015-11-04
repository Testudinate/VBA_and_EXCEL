Sub CheckColumnId()
'----------------------------------------------------------------------------------------------------------------------------
'1.Добавление новых полей.
'2.Добавление связи поля с таблицой.
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
SqlCode05 = "(SELECT CAST(CURRENT_TIME AS TIMESTAMP(0)) DTMWhenProcessed)"
Set dt = db.Execute(SqlCode05)
dttime = Format(dt!DTMWhenProcessed, "YYYY-MM-DD HH:MM:SS")
'----------------------------------------------------------------------------------------------------------------------------
Do While CallerBook.Worksheets("Рабочий лист").Cells(rownumber, 1).Value <> ""
'----------------------------------------------------------------------------------------------------------------------------
'пока нет пустой строки
'----------------------------------------------------------------------------------------------------------------------------
PTABLE_ID = Worksheets("Рабочий лист").Cells(rownumber, 1).Value                'ID таблицы
PCOLUMN_ID = Worksheets("Рабочий лист").Cells(rownumber, 3).Value               'ID поля
PCOLUMN_NAME = Worksheets("Рабочий лист").Cells(rownumber, 4).Value             'Наименование поля
PCOLUMN_COMMENT = Worksheets("Рабочий лист").Cells(rownumber, 5).Value          'Комментарий поля
PRELATED_TASKS = Worksheets("Рабочий лист").Cells(rownumber, 7).Value           'Связанные задачи
'----------------------------------------------------------------------------------------------------------------------------
If PCOLUMN_ID <> "" Then
'----------------------------------------------------------------------------------------------------------------------------
SqlCode00 = " SELECT COUNT(*) as COLUMN_ID from (SELECT c.COLUMN_ID FROM PRD_DB_DMT.PLDM_COLUMN c WHERE c.COLUMN_ID= " & PCOLUMN_ID & " ) a"
Set b = db.Execute(SqlCode00)
'----------------------------------------------------------------------------------------------------------------------------
        If b!COLUMN_ID = 0 Then
            SqlCode01 = "insert into PRD_DB_DMT.PLDM_COLUMN (COLUMN_ID,COLUMN_NAME,COLUMN_COMMENT)" & Chr(10) & Chr(13) & _
                "values (" & PCOLUMN_ID & "," & "'" & PCOLUMN_NAME & "'" & "," & "'" & PCOLUMN_COMMENT & "'" & ")"
            db.Execute (SqlCode01)
        End If
'----------------------------------------------------------------------------------------------------------------------------
SqlCode02 = " SELECT COUNT(*) as COLUMN_ID from (SELECT COLUMN_ID FROM PRD_DB_DMT.PLDM_TABLE_COLUMN_LNK WHERE COLUMN_ID = " & PCOLUMN_ID & " AND TABLE_ID = " & PTABLE_ID & ") a"
Set c = db.Execute(SqlCode02)
'----------------------------------------------------------------------------------------------------------------------------
        If c!COLUMN_ID = 0 Then
            SqlCode03 = "insert into PRD_DB_DMT.PLDM_TABLE_COLUMN_LNK (COLUMN_ID,TABLE_ID,CHANGE_DTM)" & Chr(10) & Chr(13) & _
                "values (" & PCOLUMN_ID & "," & PTABLE_ID & "," & "'" & dttime & "'" & ")"
            db.Execute (SqlCode03)
        End If
'----------------------------------------------------------------------------------------------------------------------------
        '2014-02-17 Добавлено поле PRELATED_TASKS
        'If PRELATED_TASKS <> "" Then
            'SqlCode05 = "insert into PRD_DB_DMT.PLDM_TABLE (RELATED_TASKS)" & Chr(10) & Chr(13) & _
                '"SELECT " & "'" & PRELATED_TASKS & "'" & " || " & "','" & " || COALESCE(RELATED_TASKS,'') FROM DB_ADMIN.PLDM_TABLE WHERE TABLE_ID = " & PTABLE_ID & " "
            'db.Execute (SqlCode05)
        'End If
'----------------------------------------------------------------------------------------------------------------------------
End If
'----------------------------------------------------------------------------------------------------------------------------
    rownumber = rownumber + 1
'----------------------------------------------------------------------------------------------------------------------------
Loop
'----------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------
 SqlCode06 = "Update PRD_DB_DMT.PLDM_COLUMN SET CHANGE_DTM = " & "'" & dttime & "'" & " WHERE CHANGE_DTM IS NULL;"
            db.Execute (SqlCode06)
'----------------------------------------------------------------------------------------------------------------------------
Range("J5").Value = rownumber - 2
'----------------------------------------------------------------------------------------------------------------------------
db.Close
End Sub
