Sub search_table_column_lnk()
'----------------------------------------------------------------------------------------------------------------------------
'Поиск связей поля и таблицы
'----------------------------------------------------------------------------------------------------------------------------
Set CallerBook = ThisWorkbook
Dim TimeStamp As String
'----------------------------------------------------------------------------------------------------------------------------
Debug.Print "Открытие соединения DSN" & Now()
'Создание подключения к таблице Excel
    Dim db As New ADODB.Connection
    db.ConnectionString = "DSN=TD_RDV"
    db.CommandTimeout = 0
    'Set db = New ADODB.Connection
    db.Open
'Открываем набор данных (результат выполнения запроса)
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
'----------------------------------------------------------------------------------------------------------------------------
'Указатель на номер строки
'----------------------------------------------------------------------------------------------------------------------------
rownumber = 2
'----------------------------------------------------------------------------------------------------------------------------
Do While CallerBook.Worksheets("Удаление связей полей и таблиц").Cells(rownumber, 1).Value <> ""
'----------------------------------------------------------------------------------------------------------------------------
'пока нет пустой строки
'----------------------------------------------------------------------------------------------------------------------------
PTABLE_NAME = Worksheets("Удаление связей полей и таблиц").Cells(rownumber, 1).Value               'Наименование таблицы
'----------------------------------------------------------------------------------------------------------------------------
PTABLE_NAME = UCase(PTABLE_NAME)

If PTABLE_NAME <> "" Then
   
    SqlCode00 = " SELECT t.TABLE_ID, t.TABLE_NAME, c.COLUMN_ID, c.COLUMN_NAME, c.COLUMN_COMMENT FROM PRD_VD_DMT.V_PLDM_TABLE t JOIN PRD_VD_DMT.V_PLDM_TABLE_COLUMN_LNK l ON l.TABLE_ID = t.TABLE_ID JOIN PRD_VD_DMT.V_PLDM_COLUMN c ON c.COLUMN_ID = l.COLUMN_ID WHERE t.TABLE_NAME = " & "'" & PTABLE_NAME & "'" & "  ;"
    rst.Open SqlCode00, db
     
    rst.MoveFirst
    
    i = 0
    P = 1
    Do While Not rst.EOF
                Cells(P + i, 5).Value = rst![TABLE_ID]
                Cells(P + i, 6).Value = rst![TABLE_NAME]
                Cells(P + i, 7).Value = rst![COLUMN_ID]
                Cells(P + i, 8).Value = rst![COLUMN_NAME]
                Cells(P + i, 9).Value = rst![COLUMN_COMMENT]
                i = i + 1
                rst.MoveNext
            Loop
            
End If
'----------------------------------------------------------------------------------------------------------------------------
    rownumber = rownumber + 1
'----------------------------------------------------------------------------------------------------------------------------
db.Close
Loop
'----------------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------
    Worksheets("Удаление связей полей и таблиц").Select
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
'----------------------------------------------------------------------------------------------------------------------------
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit

End Sub


