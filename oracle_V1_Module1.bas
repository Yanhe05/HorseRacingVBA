Attribute VB_Name = "Module1"
Sub linkOracle()
    Dim strConn As String
    Dim dbConn As Object
    Dim resSet As Object
    Dim j, wk As Worksheet
    Dim db_sid, db_user, db_pass As String, qsql$
    Dim i As Integer, Rank As Range, PostThis As Range, finalRow&
    
    Set wk = ThisWorkbook.Worksheets("data")
    
    db_sid = "VBAcase55"
    db_user = "RACEDATA"
    db_pass = "RACEDATA"
      
    Set dbConn = CreateObject("ADODB.Connection")
    Set resSet = CreateObject("ADODB.Recordset")
    strConn = "Provider=OraOLEDB.Oracle.1; user id=" & db_user & "; password=" & db_pass & "; data source = " & db_sid & "; Persist Security Info=True"
'    strConn = "Provider=MSDAORA.1; user id=" & db_user & "; password=" & db_pass & "; data source = " & db_sid & "; Persist Security Info=True"
        
    dbConn.Open strConn
    
    qsql = "select V.RUNNING_RACENO," & _
              "V.DISTANCE," & _
              "V.CLASS_CD," & _
              "null as Variajt_Class," & _
              "V.CENTER_COURSE_CD," & _
              "V.HORSE," & _
              "V.RANK," & _
              "V.N_LENGTHS_BEHIND_WINNER," & _
              "V.URATING," & _
              "V.LR_VARIANT," & _
              "V.LB1R_VARIANT," & _
              "V.LB2R_VARIANT," & _
              "V.LB3R_VARIANT," & _
              "V.LB4R_VARIANT," & _
              "V.LB5R_VARIANT," & _
              "V.LB6R_VARIANT," & _
              "V.LB7R_VARIANT," & _
              "V.LB8R_VARIANT," & _
              "V.LB9R_VARIANT," & _
              "V.LB10R_VARIANT," & _
              "V.LB11R_VARIANT" & _
         " from VW_VARIANT V" & _
        " where to_number(to_char(race_date, 'yyyymmdd')) = 20160316"
    Set resSet = dbConn.Execute(qsql)
    With wk
        .Rows("13:1000000").ClearContents
        For j = 0 To resSet.Fields.Count - 1
          .Cells(1, j + 1) = resSet.Fields(j).Name
        Next
        .Range("E13").CopyFromRecordset resSet
        dbConn.Close
        
         'Clear format for LR_VARIANT to LB11R VARIANT'
        .Range("N13:Z1048576").ClearFormats
        .Range("1:1").ClearContents
        
        'Hightlight when rank equal to 1 and match distance number(not done yet)'
        finalRow = .Cells(.Rows.Count, 5).End(xlUp).Row
        For i = 13 To finalRow
            Set Rank = Range("K" & i)
            Set PostThis = Range("D" & i)
            If Rank = 1 Then
                PostThis.Interior.Color = vbRed
                PostThis = .Range("F" & i)
            Else
                PostThis.Interior.Color = xlNone
            End If
        Next i
    End With
    MsgBox "complete"
End Sub

