Sub search_column()
'----------------------------------------------------------------------------------------------------------------------------
'Поиск поля
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
Do While CallerBook.Worksheets("Поиск_поля").Cells(rownumber, 1).Value <> ""
'----------------------------------------------------------------------------------------------------------------------------
'пока нет пустой строки
'----------------------------------------------------------------------------------------------------------------------------
PCOLUMN_COMMENT = Worksheets("Поиск_поля").Cells(rownumber, 1).Value               'Комментарий поля
'----------------------------------------------------------------------------------------------------------------------------
PCOLUMN_COMMENT = UCase(PCOLUMN_COMMENT)

If PCOLUMN_COMMENT <> "" Then
   
    'SqlCode00 = " SELECT COMMENT_DDL as COMMENT_DDL  FROM db_admin.v_gen_comment c WHERE LOWER(c.databasename) = LOWER( " & "'" & default_schema & "'" & ") AND LOWER(c.tablename) = LOWER(" & "'" & PTABLE_NAME & "'" & ") ORDER BY COMMENT_TYPE DESC ;"
    SqlCode00 = " SELECT TABLE_NAME,COLUMN_NAME,COLUMN_COMMENT FROM PRD_VD_DMT.V_PLDM_SEARCH_COLUMN WHERE UPPER(COLUMN_COMMENT) LIKE ('%" & PCOLUMN_COMMENT & "%');"
    rst.Open SqlCode00, db
     
    rst.MoveFirst
    
    i = 0
    P = 1
    Do While Not rst.EOF
                Cells(P + i, 7).Value = rst![TABLE_NAME]
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
    Worksheets("Поиск_поля").Select
    Range("A2:A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
'----------------------------------------------------------------------------------------------------------------------------
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
End Sub

