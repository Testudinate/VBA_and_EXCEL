Sub update_column_comment_for_PRD()
'----------------------------------------------------------------------------------------------------------------------------
'Добавление комментариев полей в модели данных, которых нет на PRD.
'Добавленные в "модели данных" можно будет подтянуть к пустым комментариям для схеме PRD_VD_DM
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
Do While CallerBook.Worksheets("Нет комментариев полей на PRD").Cells(rownumber, 1).Value <> ""
'----------------------------------------------------------------------------------------------------------------------------
'пока нет пустой строки
'----------------------------------------------------------------------------------------------------------------------------
PCOLUMN_ID = Worksheets("Нет комментариев полей на PRD").Cells(rownumber, 3).Value                  'ID поля
PCOLUMN_NAME = Worksheets("Нет комментариев полей на PRD").Cells(rownumber, 4).Value                'Наименование поля
PCOLUMN_COMMENT = Worksheets("Нет комментариев полей на PRD").Cells(rownumber, 5).Value             'Комментарий поля (добавляемый)
'----------------------------------------------------------------------------------------------------------------------------
If PCOLUMN_COMMENT <> "" Then
            
            'SqlCode03 = "update PRD_DB_DMT.PLDM_TABLE set TABLE_COMMENT= " & "'" & PTABLE_COMMENT & "'" & " where TABLE_ID=" & PTABLE_ID & "  "
        
            
            SqlCode03 = "Update PRD_DB_DMT.PLDM_COLUMN SET COLUMN_COMMENT = " & "'" & PCOLUMN_COMMENT & "'" & "WHERE COLUMN_NAME = " & "'" & PCOLUMN_NAME & "'" & " AND COLUMN_ID = " & "'" & PCOLUMN_ID & "'" & ";"
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
    Sheets("Нет комментариев полей на PRD").Select
    Selection.ListObject.QueryTable.Refresh
'----------------------------------------------------------------------------------------------------------------------------
db.Close
End Sub

