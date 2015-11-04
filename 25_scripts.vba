Sub gen_wiki()
'----------------------------------------------------------------------------------------------------------------------------
'Генерация комментариев
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
default_schema = "PRD_VD_DM"
'----------------------------------------------------------------------------------------------------------------------------
Do While CallerBook.Worksheets("Генерация вики").Cells(rownumber, 2).Value <> ""
'----------------------------------------------------------------------------------------------------------------------------
'пока нет пустой строки
'----------------------------------------------------------------------------------------------------------------------------
PTABLE_NAME = Worksheets("Генерация вики").Cells(rownumber, 2).Value               'Наименование таблицы
PSCHEMA = Worksheets("Генерация вики").Cells(rownumber, 1).Value                   'Наименование схемы
'----------------------------------------------------------------------------------------------------------------------------
If PTABLE_NAME <> "" Then
   
   If PSCHEMA = "" Then
    PSCHEMA = default_schema
   Else: default_schema = PSCHEMA
   End If
   
    'SqlCode00 = " SELECT COMMENT_DDL as COMMENT_DDL  FROM PRD_DB_DMT.v_gen_comment c WHERE LOWER(c.databasename) = LOWER( " & "'" & default_schema & "'" & ") AND LOWER(c.tablename) = LOWER(" & "'" & PTABLE_NAME & "'" & ") ORDER BY COMMENT_TYPE DESC ;"
    SqlCode00 = " SELECT WIKI_COMMENTS as WIKI_COMMENTS FROM  PRD_VD_DMT.V_GEN_WIKI_COMMENTS  C WHERE LOWER(C.DATABASENAME) = LOWER( " & "'" & default_schema & "'" & ") AND LOWER(C.TABLENAME) = LOWER(" & "'" & PTABLE_NAME & "'" & ")   ORDER BY COLUMN_ID;"
    rst.Open SqlCode00, db
     
    rst.MoveFirst
    
    i = 0
    P = 1
    Do While Not rst.EOF
                Cells(P + i, 7).Value = rst![WIKI_COMMENTS]
                i = i + 1
                rst.MoveNext
            Loop
            
End If
'----------------------------------------------------------------------------------------------------------------------------
    rownumber = rownumber + 1
'----------------------------------------------------------------------------------------------------------------------------
Loop
'----------------------------------------------------------------------------------------------------------------------------
    Range("B2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
'----------------------------------------------------------------------------------------------------------------------------
    Columns("G:G").EntireColumn.AutoFit
'----------------------------------------------------------------------------------------------------------------------------
db.Close
End Sub






