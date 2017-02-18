Attribute VB_Name = "main"
Option Explicit

Sub WashData()
    Dim ws As Worksheet
    Dim rngRow, rngCol As Range
    Dim i As Integer
    Dim index As Integer
    
    
' Set the Output WorkSheet and Clear its Contents
    Set ws = ActiveWorkbook.Worksheets("BLPINPUT")
    
    For i = ws.UsedRange.Rows.Count To 1 Step -1
        
        Set rngRow = ws.UsedRange.Rows(i)
        
        If i <> 1 And rngRow.Range("A1") <> "" Then
            ' "B1": "Ticket Num" field
            rngRow.Range("B1").Value = Trim(str(rngRow.Range("B1")))
            rngRow.Range("B1").NumberFormat = "@"
            ' "D1": "Cash Impact" field
            rngRow.Range("D1").Value = Val(Replace(rngRow.Range("D1").Value, ",", ""))
            rngRow.Range("D1").NumberFormat = "0.00"
            
            ' "E1": "Amount Type" field
            Dim nSpace As Integer
            nSpace = InStr(rngRow.Range("E1").Value, " ")
            
            If nSpace <> 0 Then
                rngRow.Range("E1").Value = Left(rngRow.Range("E1").Value, nSpace - 1)
                rngRow.Range("E1").NumberFormat = "@"
            End If
            
            ' "H1": "Settle Date" field
            rngRow.Range("H1").Value = DateValue(rngRow.Range("H1").Value)
        ElseIf i <> 1 Then
            Call rngRow.EntireRow.Delete(xlShiftUp)
        End If
    Next i
    
    For i = ws.UsedRange.Columns.Count To 1 Step -1
        
        Set rngCol = ws.UsedRange.Columns(i)
        
        If WorksheetFunction.CountA(rngCol) = 0 Then
            Call rngCol.EntireColumn.Delete(xlShiftToLeft)
        End If
        
    Next i
    
    Set ws = ActiveWorkbook.Worksheets("BOCINPUT")
    
    ws.Range("S1").Value = "Cash"
    ws.Range("T1").Value = "TKTNUM"
    
    For i = ws.UsedRange.Rows.Count To 1 Step -1
    
        Set rngRow = ws.UsedRange.Rows(i)
        
        If i <> 1 And rngRow.Range("A1") <> "" Then
            ' "M1": "Debit/ Credit" field
            ' "N1": "Amount" field
            If rngRow.Range("M1").Value = "D" Then
                rngRow.Range("S1").Value = -1 * Val(rngRow.Range("N1").Value)
            Else
                rngRow.Range("S1").Value = Val(rngRow.Range("N1").Value)
            End If
            
            rngRow.Range("S1").NumberFormat = "0.00"
            
            ' "Q1": "Particulars" field
            Dim nDelimit As Integer
            nDelimit = InStr(rngRow.Range("Q1").Value, "//")
            
            If nDelimit <> 0 And IsNumeric(Mid(rngRow.Range("Q1").Value, nDelimit + 2, 5)) Then
                rngRow.Range("T1").Value = Trim(str(Mid(rngRow.Range("Q1").Value, nDelimit + 2, 5)))
            End If
            
            If InStr(rngRow.Range("Q1").Value, "517817") <> 0 Then
                rngRow.Range("T1").Value = "SETOFF"
            End If
            
            If InStr(rngRow.Range("Q1").Value, "REMIT") <> 0 Then
                rngRow.Range("T1").Value = Left(rngRow.Range("Q1").Value, InStr(rngRow.Range("Q1").Value, "/") - 1)
            End If
            
            If rngRow.Range("T1").Value = "" Or IsEmpty(rngRow.Range("T1").Value) Then
                rngRow.Range("T1").Value = rngRow.Range("Q1").Value
            End If
            
            rngRow.Range("T1").NumberFormat = "@"
        ElseIf i <> 1 Then
            Call rngRow.EntireRow.Delete(xlShiftUp)
        End If
    Next i
    
    For i = ws.UsedRange.Columns.Count To 1 Step -1
        
        Set rngCol = ws.UsedRange.Columns(i)
        
        If WorksheetFunction.CountA(rngCol) = 0 Then
            Call rngCol.EntireColumn.Delete(xlShiftToLeft)
        End If
        
    Next i
    
    ActiveWorkbook.Save
    
End Sub

Sub BLPGrpByTkt()
    'Run after WashData()

    Dim strSQL As String
    
    strSQL = "SELECT `Ticket Num`, `Settle Date` " & _
             "FROM [BLP$] WHERE CDbl(`Cash Impact`) <> 0 " & _
             "GROUP BY `Ticket Num`, `Settle Date`, `Amount Type`, Counterparty, Currency " & _
             "HAVING COUNT(*) = 1 "
             
    Call SQLOnWS(strSQL, "BLPSINGLETKT")
    
    strSQL = "SELECT `Ticket Num`, SUM(`Cash Impact`) as `Cash Impact`, `Settle Date`, `Amount Type`, Counterparty, Currency " & _
             "FROM [BLP$] WHERE CDbl(`Cash Impact`) <> 0 " & _
             "GROUP BY `Ticket Num`, `Settle Date`, `Amount Type`, Counterparty, Currency " & _
             "HAVING COUNT(*) > 1 " & _
             "UNION " & _
             "SELECT A.`Ticket Num`, A.`Cash Impact`, A.`Settle Date`, A.`Amount Type`, A.Counterparty, A.Currency " & _
             "FROM [BLP$] A INNER JOIN [BLPSINGLETKT$] B ON A.`Ticket Num` = B.`Ticket Num` AND A.`Settle Date` = B.`Settle Date`"

    Call SQLOnWS(strSQL, "BLPGRPBYTKT")
    
End Sub

Sub BLPTKTGrpFX()
    'Run after BLPGrpByTkt()
    
    Dim strSQL As String
    
    strSQL = "SELECT * FROM [BLPGRPBYTKT$] WHERE `Amount Type` <> ""FX"" " & _
             "UNION " & _
             "SELECT ""SETOFF"" AS `Ticket Num`, SUM(`Cash Impact`) AS `Cash Impact`, `Settle Date`, ""FX"" AS `Amount Type`, Counterparty, Currency FROM [BLPGRPBYTKT$] WHERE `Amount Type` = ""FX"" GROUP BY `Settle Date`, Counterparty, Currency " & _
             "ORDER BY `Settle Date`"
    
    Call SQLOnWS(strSQL, "BLPTKTGRPFX")
    
End Sub

Sub BOCClean()
    'Run after WashData()

    Dim strSQL As String
    
    strSQL = "SELECT Cash, TKTNUM, Amount, `Value Date`, Currency FROM [BOC$] WHERE TKTNUM <> ""SETOFF"" AND TKTNUM NOT LIKE ""%INT. DATE%"" " & _
             "UNION ALL " & _
             "SELECT SUM(Cash) AS Cash, TKTNUM, ABS(SUM(Cash)) AS Amount, `Value Date`, Currency FROM [BOC$] WHERE TKTNUM = ""SETOFF"" GROUP BY `Value Date`, TKTNUM, Currency"
    
    Call SQLOnWS(strSQL, "BOCCLEAN1")

             
    strSQL = "SELECT * FROM (SELECT SUM(Cash) AS Cash, TKTNUM, AMOUNT, MAX(CDate(`Value Date`)) AS `Value Date`, Currency FROM (" & _
             "SELECT Cash, TKTNUM, Amount, `Value Date`, Currency " & _
             "FROM [BOCCLEAN1$] " & _
             ") GROUP BY TKTNUM, AMOUNT, Currency) AS S2 WHERE S2.CASH <> 0 ORDER BY `Value Date`"
             
    Call SQLOnWS(strSQL, "BOCCLEAN")
End Sub

Sub BLPTKTFXWithBOCCLEAN()
    'Run after BLPTKTGrpFX() and BOCCLEAN()
    
    Dim strSQL As String
    
    strSQL = "SELECT [BLPTKTGRPFX$].*, [BOCCLEAN$].* " & _
             "FROM [BLPTKTGRPFX$] INNER JOIN [BOCCLEAN$] " & _
             "ON CStr([BLPTKTGRPFX$].`Ticket Num`) = CStr([BOCCLEAN$].TKTNUM) " & _
             "AND CStr([BLPTKTGRPFX$].Currency) = CStr([BOCCLEAN$].Currency) " & _
             "AND Abs(CDbl([BLPTKTGRPFX$].`Cash Impact`) - CDbl([BOCCLEAN$].Cash)) < 0.05"
             
    Call SQLOnWS(strSQL, "BLPTKTFXWITHBOCCLEAN")
    
    strSQL = "SELECT [BLPTKTGRPFX$].* " & _
             "FROM [BLPTKTGRPFX$] LEFT JOIN [BOCCLEAN$] " & _
             "ON Cstr([BLPTKTGRPFX$].`Ticket Num`) = CStr([BOCCLEAN$].TKTNUM) " & _
             "AND CStr([BLPTKTGRPFX$].Currency) = CStr([BOCCLEAN$].Currency) " & _
             "AND Abs(CDbl([BLPTKTGRPFX$].`Cash Impact`) - CDbl([BOCCLEAN$].Cash)) < 0.05 " & _
             "WHERE [BOCCLEAN$].Cash IS NULL"
             
    Call SQLOnWS(strSQL, "BLPTKTNUMNOMATCH")
    
    strSQL = "SELECT [BOCCLEAN$].* " & _
             "FROM [BLPTKTGRPFX$] RIGHT JOIN [BOCCLEAN$] " & _
             "ON Cstr([BLPTKTGRPFX$].`Ticket Num`) = CStr([BOCCLEAN$].TKTNUM) " & _
             "AND CStr([BLPTKTGRPFX$].Currency) = CStr([BOCCLEAN$].Currency) " & _
             "AND Abs(CDbl([BLPTKTGRPFX$].`Cash Impact`) - CDbl([BOCCLEAN$].Cash)) < 0.05 " & _
             "WHERE [BLPTKTGRPFX$].`Cash Impact` IS NULL"
             
    Call SQLOnWS(strSQL, "BOCTKTNUMNOMATCH")
    
End Sub

Sub BLPGrpREPOAndWithBOC()
    'Run after BLPTKTFXWithBOCCLEAN()
    
    Dim strSQL As String
    
    strSQL = "SELECT * FROM [BLPTKTNUMNOMATCH$] WHERE `Amount Type` <> ""Repo"" " & _
             "UNION " & _
             "SELECT NULL AS `Ticket Num`, SUM(`Cash Impact`), `Settle Date`, `Amount Type`, Counterparty, Currency FROM [BLPTKTNUMNOMATCH$] WHERE `Amount Type` = ""Repo"" GROUP BY `Amount Type`, `Settle Date`, Counterparty, Currency"
    
    Call SQLOnWS(strSQL, "BLPGRPREPO")
    
    strSQL = "SELECT [BLPGRPREPO$].*, [BOCTKTNUMNOMATCH$].* " & _
             "FROM [BLPGRPREPO$] INNER JOIN [BOCTKTNUMNOMATCH$] " & _
             "ON Abs(CDbl([BLPGRPREPO$].`Cash Impact`) - CDbl([BOCTKTNUMNOMATCH$].Cash)) < 0.05 " & _
             "AND CStr([BLPGRPREPO$].Currency) = CStr([BOCTKTNUMNOMATCH$].Currency)"
             
    Call SQLOnWS(strSQL, "BLPBOCFINALMATCH")
    
    strSQL = "SELECT [BLPGRPREPO$].*, [BOCTKTNUMNOMATCH$].* " & _
             "FROM [BLPGRPREPO$] LEFT JOIN [BOCTKTNUMNOMATCH$] " & _
             "ON Abs(CDbl([BLPGRPREPO$].`Cash Impact`) - CDbl([BOCTKTNUMNOMATCH$].Cash)) < 0.05 " & _
             "AND CStr([BLPGRPREPO$].Currency) = CStr([BOCTKTNUMNOMATCH$].Currency) " & _
             "WHERE [BOCTKTNUMNOMATCH$].Cash IS NULL AND [BLPGRPREPO$].`Cash Impact` <> 0 " & _
             "UNION ALL " & _
             "SELECT [BLPGRPREPO$].*, [BOCTKTNUMNOMATCH$].* " & _
             "FROM [BLPGRPREPO$] RIGHT JOIN [BOCTKTNUMNOMATCH$] " & _
             "ON Abs(CDbl([BLPGRPREPO$].`Cash Impact`) - CDbl([BOCTKTNUMNOMATCH$].Cash)) < 0.05 " & _
             "AND CStr([BLPGRPREPO$].Currency) = CStr([BOCTKTNUMNOMATCH$].Currency) " & _
             "WHERE [BLPGRPREPO$].`Cash Impact` IS NULL AND [BOCTKTNUMNOMATCH$].Cash <> 0"
    
    Call SQLOnWS(strSQL, "BLPBOCFINALNONMATCH")

End Sub

Sub SQLOnWS(strSQL As String, strOutputWS As String)
    
    Dim oCONN As ADODB.Connection
    Dim oRS As ADODB.Recordset
    Dim strConn As String
    Dim wsOutput As Worksheet
    
    Dim i As Long
    Dim hdrName As Variant
    
' Set the Output WorkSheet and Clear its Contents
    If Evaluate("ISREF('" & strOutputWS & "'!A1)") Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Worksheets(strOutputWS).Delete
        Application.DisplayAlerts = True
    End If
    
    Set wsOutput = ActiveWorkbook.Worksheets.Add
    wsOutput.Name = strOutputWS
    wsOutput.Cells.ClearContents
    
' Establish the ADO Connection, Recordset & Connection String variables for the ActiveWorkbook
' "Microsoft.ACE.OLEDB.12.0" and "Excel 12.0" imply Excel 2007 or higher
    Set oCONN = New ADODB.Connection
    Set oRS = New ADODB.Recordset
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & ActiveWorkbook.FullName & ";" & _
                "Extended Properties=""Excel 12.0;HDR=Yes;"";"
    
' Open the Connection and run the SQL Statement
    oCONN.Open strConn
    oRS.Open strSQL, oCONN

' Write the Headers to the Output worksheet
    i = 1
    For Each hdrName In oRS.Fields
        wsOutput.Cells(1, i).Value = hdrName.Name
        i = i + 1
    Next
    
' Write the Result Set of the SQL Query to the Output Worksheet
    wsOutput.Cells(2, 1).CopyFromRecordset oRS
    
' Convert the Output data list to an Excel Data
    With wsOutput
        .Activate
        .ListObjects.Add(xlSrcRange, ActiveSheet.UsedRange, , xlYes).Name = "Tbl_Output"
        .Cells.EntireColumn.AutoFit
    End With
    
' Cleanup and close the connection
    oCONN.Close
    Set oCONN = Nothing
    Set oRS = Nothing
    
    ActiveWorkbook.Save
    
End Sub

Sub SQLInPlace(strSQL As String)
    
    Dim oCONN As ADODB.Connection
    Dim oRS As ADODB.Recordset
    Dim strConn As String
    Dim wsOutput As Worksheet
    
    Dim i As Long
    Dim hdrName As Variant
    
' Establish the ADO Connection, Recordset & Connection String variables for the ActiveWorkbook
' "Microsoft.ACE.OLEDB.12.0" and "Excel 12.0" imply Excel 2007 or higher
    Set oCONN = New ADODB.Connection
    Set oRS = New ADODB.Recordset
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & ActiveWorkbook.FullName & ";" & _
                "Extended Properties=""Excel 12.0;HDR=Yes;"";"
    
' Open the Connection and run the SQL Statement
    oCONN.Open strConn
    oRS.Open strSQL, oCONN
    
' Cleanup and close the connection
    oCONN.Close
    Set oCONN = Nothing
    Set oRS = Nothing
    
    ActiveWorkbook.Save
    
End Sub

Function SQLInArray(strSQL As String) As Variant

    Dim result()
    
    Dim oCONN As ADODB.Connection
    Dim oRS As ADODB.Recordset
    Dim strConn As String
    Dim wsOutput As Worksheet
    
    Dim i As Long
    Dim row As Long
    Dim cols As Long

    Dim hdrName As Variant
    
' Establish the ADO Connection, Recordset & Connection String variables for the ActiveWorkbook
' "Microsoft.ACE.OLEDB.12.0" and "Excel 12.0" imply Excel 2007 or higher
    Set oCONN = New ADODB.Connection
    Set oRS = New ADODB.Recordset
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & ActiveWorkbook.FullName & ";" & _
                "Extended Properties=""Excel 12.0;HDR=Yes;"";"
    
' Open the Connection and run the SQL Statement
    oCONN.Open strConn
    oRS.Open strSQL, oCONN
    
    cols = oRS.Fields.Count
    
    row = 0
    
    Do Until oRS.EOF
    
        ReDim Preserve result(cols - 1, row)
    
        For i = 0 To cols - 1
            result(i, row) = oRS.Fields.Item(i)
        Next i
        
        oRS.MoveNext
        
        row = row + 1
        
    Loop
    
' Cleanup and close the connection
    oCONN.Close
    Set oCONN = Nothing
    Set oRS = Nothing
    
    SQLInArray = result

End Function


Sub GetDateRange()
    
    Dim strBLPSQL As String, strBOCSQL As String
    Dim ws As Worksheet
    
    strBLPSQL = "SELECT * FROM [BLPINPUT$]"
    strBOCSQL = "SELECT * FROM [BOCINPUT$]"
    
    
    Set ws = ActiveWorkbook.Worksheets("Command")
    
    If Not IsEmpty(ws.Range("enddate").Value) Then
        
        strBLPSQL = strBLPSQL & " WHERE `Settle Date` <= CDate(""" & ws.Range("enddate").Value & """)"
        strBOCSQL = strBOCSQL & " WHERE `Value Date` <= CDate(""" & ws.Range("enddate").Value & """)"
        
    End If
    
    Call SQLOnWS(strBLPSQL, "BLP")
    Call SQLOnWS(strBOCSQL, "BOC")
    
    strBLPSQL = "UPDATE [BLP$] SET Account = ""HRAM1"" WHERE Account = ""HRAMLEV"""
    
    Call SQLInPlace(strBLPSQL)
    
End Sub

Sub Check()
    
    Call WashData
    Call GetDateRange
    Call BLPGrpByTkt
    Call BLPTKTGrpFX
    Call BOCClean
    Call BLPTKTFXWithBOCCLEAN
    Call BLPGrpREPOAndWithBOC
    
End Sub

'Experiments

Sub trial()

    Dim strSQL As String
    Dim strBLPSQL As String
    Dim prev, curr As Double
    
    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Worksheets("Command")
    
    If Not IsEmpty(ws.Range("prev").Value) Then
                
        prev = ws.Range("prev").Value
        
    Else
        MsgBox "Please specify the cash of previous month."
        Exit Sub
    End If
             
    strSQL = "SELECT SUM(`Cash Impact`) FROM [BLPCATE$]"
    
    Dim result()
    result = SQLInArray(strSQL)
    
    curr = result(0, 0) + prev
    
    ws.Range("curr").Value = curr
    
    
    
    'Call SQLOnWS(strSQL, "SINGLETKT")
    
    'strSQL = "SELECT `Ticket Num`, SUM(`Cash Impact`) as `Cash Impact`, `Settle Date`, `Amount Type`, Counterparty, Currency " & _
             "FROM [BLP$] WHERE CDbl(`Cash Impact`) <> 0 " & _
             "GROUP BY `Ticket Num`, `Settle Date`, `Amount Type`, Counterparty, Currency " & _
             "HAVING COUNT(*) > 1 " & _
             "UNION " & _
             "SELECT A.`Ticket Num`, A.`Cash Impact`, A.`Settle Date`, A.`Amount Type`, A.Counterparty, A.Currency " & _
             "FROM [BLP$] A INNER JOIN [SINGLETKT$] B ON A.`Ticket Num` = B.`Ticket Num` AND A.`Settle Date` = B.`Settle Date`"
    
    'Call SQLOnWS(strSQL, "trial")
    
    
End Sub

Sub SETTSTAT()
    
    Dim strBLPSQL As String
                
    strBLPSQL = "SELECT * FROM " & _
                "(SELECT SUM(`Cash Impact`) AS BLPRECEIVABLE FROM [BLPBOCFINALNONMATCH$] WHERE Cash IS NULL AND `Cash Impact` > 0) A, " & _
                "(SELECT SUM(`Cash Impact`) AS BLPPAYABLE FROM [BLPBOCFINALNONMATCH$] WHERE Cash IS NULL AND `Cash Impact` < 0) B, " & _
                "(SELECT SUM(Cash) AS BOCUNCONFIRMEDIN FROM [BLPBOCFINALNONMATCH$] WHERE `Cash Impact` IS NULL AND Cash > 0) C, " & _
                "(SELECT SUM(Cash) AS BOCUNCONFIRMEDOUT FROM [BLPBOCFINALNONMATCH$] WHERE `Cash Impact` IS NULL AND Cash < 0) D"
    
    
    Call SQLOnWS(strBLPSQL, "UNSETTSTAT")
    
    strBLPSQL = "SELECT `Ticket Num`, `Cash Impact`, `Settle Date`, `Amount Type`, Counterparty, `BLPGRPREPO$#Currency` AS BLPCurrency, Cash, TKTNUM, AMOUNT, `Value Date`, `BOCTKTNUMNOMATCH$#Currency` AS BOCCurrency FROM [BLPBOCFINALMATCH$] " & _
                "UNION " & _
                "SELECT `Ticket Num`, `Cash Impact`, `Settle Date`, `Amount Type`, Counterparty, `BLPTKTGRPFX$#Currency` AS BLPCurrency, Cash, TKTNUM, AMOUNT, `Value Date`, `BOCCLEAN$#Currency` AS BOCCurrency FROM [BLPTKTFXWITHBOCCLEAN$] "
    
    Call SQLOnWS(strBLPSQL, "SETTSTAT")
    
End Sub

Sub GetBLPSett()
    
    Dim strBLPSQL As String
    Dim strStartDate As String
    
    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Worksheets("Command")
    
    If Not IsEmpty(ws.Range("startdate").Value) Then
                
        strStartDate = ws.Range("startdate").Value
        
    Else
        MsgBox "Please specify the Start Date."
        Exit Sub
    End If
    
    strBLPSQL = "SELECT * FROM " & _
                "(SELECT SUM(`Cash Impact`) AS BLPRECEIVABLE FROM [BLPBOCFINALNONMATCH$] WHERE Cash IS NULL AND `Cash Impact` > 0) A, " & _
                "(SELECT SUM(`Cash Impact`) AS BLPPAYABLE FROM [BLPBOCFINALNONMATCH$] WHERE Cash IS NULL AND `Cash Impact` < 0) B, " & _
                "(SELECT SUM(Cash) AS BOCUNCONFIRMEDIN FROM [BLPBOCFINALNONMATCH$] WHERE `Cash Impact` IS NULL AND Cash > 0) C, " & _
                "(SELECT SUM(Cash) AS BOCUNCONFIRMEDOUT FROM [BLPBOCFINALNONMATCH$] WHERE `Cash Impact` IS NULL AND Cash < 0) D"
    
    Call SQLOnWS(strBLPSQL, "UNSETTSTAT")
    
    strBLPSQL = "SELECT `Ticket Num`, `Cash Impact`, `Settle Date`, `Amount Type`, Counterparty, `BLPGRPREPO$#Currency` AS BLPCurrency, Cash, TKTNUM, AMOUNT, `Value Date`, `BOCTKTNUMNOMATCH$#Currency` AS BOCCurrency FROM [BLPBOCFINALMATCH$] " & _
                "UNION " & _
                "SELECT `Ticket Num`, `Cash Impact`, `Settle Date`, `Amount Type`, Counterparty, `BLPTKTGRPFX$#Currency` AS BLPCurrency, Cash, TKTNUM, AMOUNT, `Value Date`, `BOCCLEAN$#Currency` AS BOCCurrency FROM [BLPTKTFXWITHBOCCLEAN$] "
    
    Call SQLOnWS(strBLPSQL, "SETTSTAT")
    
    strBLPSQL = "SELECT A.* FROM [BLP$] A LEFT JOIN [BLPBOCFINALNONMATCH$] B ON A.`Ticket Num` = B.`Ticket Num` WHERE B.`Ticket Num` IS NULL AND A.`Settle Date` >= CDate(""" & strStartDate & """)"
    
    Call SQLOnWS(strBLPSQL, "BLPSETT")
    
    strBLPSQL = "SELECT `Amount Type`, SUM(`Cash Impact`) AS `Cash Impact` FROM [BLPSETT$] WHERE (Security NOT LIKE "".%"" OR Security IS NULL) GROUP BY `Amount Type` " & _
                "UNION " & _
                "SELECT ""Fee"" AS `Amount Type`, SUM(`Cash Impact`) AS `Cash Impact` FROM [BLPSETT$] WHERE Security LIKE "".%"" " & _
                "UNION SELECT ""BOC UNCONFIRMED IN"" AS `Amount Type`, BOCUNCONFIRMEDIN AS `Cash Impact` FROM [UNSETTSTAT$] " & _
                "UNION SELECT ""BOC UNCONFIRMED OUT"" AS `Amount Type`, BOCUNCONFIRMEDOUT AS `Cash Impact` FROM [UNSETTSTAT$] " & _
                "UNION SELECT ""BLP RECEIVABLE LAST MONTH"" AS `Amount Type`, SUM(`Cash Impact`) AS `Cash Impact` FROM [SETTSTAT$] WHERE `Settle Date` < CDate(""" & strStartDate & """) AND `Cash Impact` > 0 " & _
                "UNION SELECT ""BLP PAYABLE LAST MONTH"" AS `Amount Type`, SUM(`Cash Impact`) AS `Cash Impact` FROM [SETTSTAT$] WHERE `Settle Date` < CDate(""" & strStartDate & """) AND `Cash Impact` < 0"
                
    
    Call SQLOnWS(strBLPSQL, "BLPCATE")
    
End Sub

Sub GetSingleFX()
    
    Dim strBLPSQL As String
    
    strBLPSQL = "SELECT A.Date, A.`Ticket Num`, A.Security, A.`Cash Impact` AS LEG1, ABS(A.`Cash Impact`) AS QTYLEG1, B.`Cash Impact` AS LEG2, ABS(B.`Cash Impact`) AS QTYLEG2, A.`Accounting Date` AS TradeDate FROM (SELECT * FROM [FX$] WHERE `Amount Type` = ""FX Leg 1"") A INNER JOIN (SELECT * FROM [FX$] WHERE `Amount Type` = ""FX Leg 2"") B ON A.`Ticket Num` = B.`Ticket Num`"
    Call SQLOnWS(strBLPSQL, "DISTINCT")
    
    
    strBLPSQL = "SELECT QTYLEG1, TradeDate FROM [DISTINCT$] GROUP BY QTYLEG1, TradeDate HAVING COUNT(*) MOD 2 <> 0 OR SUM(LEG1) <> 0"
    Call SQLOnWS(strBLPSQL, "TMP1")
    
    strBLPSQL = "SELECT QTYLEG2, TradeDate FROM [DISTINCT$] GROUP BY QTYLEG2, TradeDate HAVING COUNT(*) MOD 2 <> 0 OR SUM(LEG2) <> 0"
    Call SQLOnWS(strBLPSQL, "TMP2")
    
    strBLPSQL = "SELECT A.Date, A.`Ticket Num`, A.Security, A.LEG1, A.LEG2, A.TradeDate FROM ([DISTINCT$] A LEFT JOIN [TMP1$] B ON A.QTYLEG1 = B.QTYLEG1 AND A.TradeDate = B.TradeDate) LEFT JOIN [TMP2$] C ON A.QTYLEG2 = C.QTYLEG2 AND A.TradeDate = C.TradeDate WHERE B.QTYLEG1 IS NOT NULL AND C.QTYLEG2 IS NOT NULL"
    Call SQLOnWS(strBLPSQL, "SINGLE")
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets("TMP1").Delete
    ActiveWorkbook.Worksheets("TMP2").Delete
    Application.DisplayAlerts = True
    
End Sub

Sub GetSwapFX()

    Dim strBLPSQL As String
    
    strBLPSQL = "SELECT A.Date, A.`Ticket Num`, A.Security, A.LEG1, A.LEG2, A.TradeDate FROM [DISTINCT$] A LEFT JOIN [SINGLE$] B ON A.`Ticket Num` = B.`Ticket Num` AND A.TradeDate = B.TradeDate WHERE B.`Ticket Num` IS NULL"
    
    Call SQLOnWS(strBLPSQL, "SWAP")

End Sub

Sub GetFX()
    
    Dim strBLPSQL As String
    
    strBLPSQL = "SELECT * FROM [ALL$] WHERE `Amount Type` LIKE ""FX%"""
    
    Call SQLOnWS(strBLPSQL, "FX")
    
    strBLPSQL = "UPDATE [FX$] SET [FX$].`Accounting Date` = [FX$].`Orig Trade Date` WHERE [FX$].`Accounting Date` IS NULL"
    Call SQLInPlace(strBLPSQL)
    
End Sub

Sub MAVCheck()

    Dim strBLPSQL As String

    strBLPSQL = "SELECT `Account Code`, Name, IIF(SEC_INCOME_YTD IS NULL, 0, SEC_INCOME_YTD) AS SEC_INCOME_YTD, IIF(ts_ytd IS NULL, 0, ts_ytd) AS ts_ytd, IIF(`UnRealised P&L Y` IS NULL, 0, `UnRealised P&L Y`) AS `UnRealised P&L Y`, IIF(f_realpnl_ytd IS NULL, 0, f_realpnl_ytd) AS f_realpnl_ytd, IIF(f_totalpnl_ytd IS NULL, 0, f_totalpnl_ytd) AS f_totalpnl_ytd, (IIF(SEC_INCOME_YTD IS NULL, 0, SEC_INCOME_YTD) + IIF(ts_ytd IS NULL, 0, ts_ytd) + IIF(`UnRealised P&L Y` IS NULL, 0, `UnRealised P&L Y`) + IIF(f_realpnl_ytd IS NULL, 0, f_realpnl_ytd) - IIF(f_totalpnl_ytd IS NULL, 0, f_totalpnl_ytd)) AS `P&L_YTD_CHECK` FROM [USD$] WHERE `Account Code` LIKE ""HRAM%"" AND Name IN (""Corporate Bond"", ""Convertible Bond"", ""Equity"", ""Treasury"", ""Money Market"", ""Repo Liability"")"
    
    Call SQLOnWS(strBLPSQL, "USDSTAT")
    
    strBLPSQL = "SELECT `Account Code`, Name, IIF(SEC_INCOME_YTD IS NULL, 0, SEC_INCOME_YTD) AS SEC_INCOME_YTD, IIF(ts_ytd IS NULL, 0, ts_ytd) AS ts_ytd, IIF(`UnRealised P&L Y` IS NULL, 0, `UnRealised P&L Y`) AS `UnRealised P&L Y`, IIF(f_realpnl_ytd IS NULL, 0, f_realpnl_ytd) AS f_realpnl_ytd, IIF(f_totalpnl_ytd IS NULL, 0, f_totalpnl_ytd) AS f_totalpnl_ytd, (IIF(SEC_INCOME_YTD IS NULL, 0, SEC_INCOME_YTD) + IIF(ts_ytd IS NULL, 0, ts_ytd) + IIF(`UnRealised P&L Y` IS NULL, 0, `UnRealised P&L Y`) + IIF(f_realpnl_ytd IS NULL, 0, f_realpnl_ytd) - IIF(f_totalpnl_ytd IS NULL, 0, f_totalpnl_ytd)) AS `P&L_YTD_CHECK` FROM [EUR$] WHERE `Account Code` LIKE ""HRAM%"" AND Name IN (""Corporate Bond"", ""Convertible Bond"", ""Equity"", ""Treasury"", ""Money Market"", ""Repo Liability"")"
    
    Call SQLOnWS(strBLPSQL, "EURSTAT")
    
    strBLPSQL = "SELECT `Account Code`, Name, IIF(SEC_INCOME_YTD IS NULL, 0, SEC_INCOME_YTD) AS SEC_INCOME_YTD, IIF(ts_ytd IS NULL, 0, ts_ytd) AS ts_ytd, IIF(`UnRealised P&L Y` IS NULL, 0, `UnRealised P&L Y`) AS `UnRealised P&L Y`, IIF(f_realpnl_ytd IS NULL, 0, f_realpnl_ytd) AS f_realpnl_ytd, IIF(f_totalpnl_ytd IS NULL, 0, f_totalpnl_ytd) AS f_totalpnl_ytd, (IIF(SEC_INCOME_YTD IS NULL, 0, SEC_INCOME_YTD) + IIF(ts_ytd IS NULL, 0, ts_ytd) + IIF(`UnRealised P&L Y` IS NULL, 0, `UnRealised P&L Y`) + IIF(f_realpnl_ytd IS NULL, 0, f_realpnl_ytd) - IIF(f_totalpnl_ytd IS NULL, 0, f_totalpnl_ytd)) AS `P&L_YTD_CHECK` FROM [HKD$] WHERE `Account Code` LIKE ""HRAM%"" AND Name IN (""Corporate Bond"", ""Convertible Bond"", ""Equity"", ""Treasury"", ""Money Market"", ""Repo Liability"")"
     
    Call SQLOnWS(strBLPSQL, "HKDSTAT")
    
    strBLPSQL = "SELECT `Account Code`, Name, IIF(SEC_INCOME_YTD IS NULL, 0, SEC_INCOME_YTD) AS SEC_INCOME_YTD, IIF(ts_ytd IS NULL, 0, ts_ytd) AS ts_ytd, IIF(`UnRealised P&L Y` IS NULL, 0, `UnRealised P&L Y`) AS `UnRealised P&L Y`, IIF(f_realpnl_ytd IS NULL, 0, f_realpnl_ytd) AS f_realpnl_ytd, IIF(f_totalpnl_ytd IS NULL, 0, f_totalpnl_ytd) AS f_totalpnl_ytd, (IIF(SEC_INCOME_YTD IS NULL, 0, SEC_INCOME_YTD) + IIF(ts_ytd IS NULL, 0, ts_ytd) + IIF(`UnRealised P&L Y` IS NULL, 0, `UnRealised P&L Y`) + IIF(f_realpnl_ytd IS NULL, 0, f_realpnl_ytd) - IIF(f_totalpnl_ytd IS NULL, 0, f_totalpnl_ytd)) AS `P&L_YTD_CHECK` FROM [CNY$] WHERE `Account Code` LIKE ""HRAM%"" AND Name IN (""Corporate Bond"", ""Convertible Bond"", ""Equity"", ""Treasury"", ""Money Market"", ""Repo Liability"")"
    
    Call SQLOnWS(strBLPSQL, "CNYSTAT")
    
End Sub

Sub PnL(strTBL As String, strCurrency As String)
    
    Dim strBLPSQL As String
    
    strBLPSQL = "SELECT """ & strTBL & """ AS `Category`, """ & strCurrency & """ AS `Currency`, " & _
                "SUM(Cash) As PnL FROM (" & _
                "SELECT SUM(LEG1) AS Cash FROM [" & strTBL & "$] WHERE Security LIKE """ & strCurrency & "/%"" " & _
                "UNION " & _
                "SELECT SUM(LEG2) AS Cash FROM [" & strTBL & "$] WHERE Security LIKE ""%/" & strCurrency & "%""" & _
                ")"
                
    Call SQLOnWS(strBLPSQL, strTBL & " " & strCurrency & " PnL")
    
End Sub

Sub ALLPnL()

    Call PnL("SINGLE", "USD")
    Call PnL("SINGLE", "EUR")
    Call PnL("SINGLE", "HKD")
    Call PnL("SINGLE", "CNH")
    
    Call PnL("SWAP", "USD")
    Call PnL("SWAP", "EUR")
    Call PnL("SWAP", "HKD")
    Call PnL("SWAP", "CNH")
    
End Sub


Sub CombineCCMSTH()
    Dim strSQL As String
    Dim strBLPSQL As String
    
    strBLPSQL = "SELECT [BLPINPUT$].*, [BLPTHINPUT$].`ISIN Number`, [BLPTHINPUT$].`Total Commission`, " & _
                "[BLPTHINPUT$].`Settlement Total in Settlement Currency`, " & _
                "[BLPTHINPUT$].`Settlement Accrued in Settlement Currency`, " & _
                "[BLPTHINPUT$].`Settlement Principal in Settlement Currency` " & _
                "FROM [BLPINPUT$] LEFT JOIN [BLPTHINPUT$] " & _
                "ON CStr([BLPINPUT$].`Ticket Num`) = CStr([BLPTHINPUT$].Ticket) " & _
                "AND CStr([BLPINPUT$].Account) =  CStr([BLPTHINPUT$].Account)"
    
    Call SQLOnWS(strBLPSQL, "BLPTH")
    
    strBLPSQL = "UPDATE [BLPTH$] SET `Total Commission` = -1 * ABS(`Total Commission`)"
    
    Call SQLInPlace(strBLPSQL)
    
    strBLPSQL = "UPDATE [BLPTH$] SET " & _
                "`Settlement Total in Settlement Currency` = IIF(`Cash Impact` <  0, -1, 1) * ABS(`Settlement Total in Settlement Currency`), " & _
                "`Settlement Accrued in Settlement Currency` = IIF(`Cash Impact` <  0, -1, 1) * ABS(`Settlement Accrued in Settlement Currency`), " & _
                "`Settlement Principal in Settlement Currency` = IIF(`Cash Impact` <  0, -1, 1) * ABS(`Settlement Principal in Settlement Currency`)"
    
    Call SQLInPlace(strBLPSQL)
    
End Sub
