Sub delete_not_exist_pldm()
'----------------------------------------------------------------------------------------------------------------------------
'Удаление неактульных БС
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
Do While CallerBook.Worksheets("Удаление неактуальных БС").Cells(rownumber, 1).Value <> ""
'----------------------------------------------------------------------------------------------------------------------------
'пока нет пустой строки
'----------------------------------------------------------------------------------------------------------------------------
PBSN_ENT_ID = Worksheets("Удаление неактуальных БС").Cells(rownumber, 1).Value                 'ID БС
'----------------------------------------------------------------------------------------------------------------------------
If PBSN_ENT_ID <> "" Then
            
            SqlCode02 = "delete from PRD_DB_DMT.PLDM WHERE BSN_ENT_ID = " & PBSN_ENT_ID & ";"
            db.Execute (SqlCode02)
              
End If
'----------------------------------------------------------------------------------------------------------------------------
    rownumber = rownumber + 1
'----------------------------------------------------------------------------------------------------------------------------
Loop
'----------------------------------------------------------------------------------------------------------------------------
    Range("A2:C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
'----------------------------------------------------------------------------------------------------------------------------
    Sheets("Удаление неактуальных БС").Select
    Selection.ListObject.QueryTable.Refresh
'----------------------------------------------------------------------------------------------------------------------------
db.Close
End Sub



