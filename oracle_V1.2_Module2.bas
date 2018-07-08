Attribute VB_Name = "Module2"
Sub Button2_Click()
    Dim wk As Worksheet
    Dim endRow&, i&, j&, k&, m&, n&, total, h&, key, ky, arr(), temp
    Set wk = ThisWorkbook.Worksheets("data")
    With wk
        endRow = .Cells(12, 5).End(xlDown).Row
        .Range("K3:M11").ClearContents
        For k = 11 To 13
            For j = 3 To 11
                key = .Range("J" & j) & "-" & .Cells(2, k)
                total = 0
                h = 0
                For i = 13 To endRow
                    If key = .Range("F" & i) & "-" & .Range("I" & i) Then
                        For m = 15 To 24
                            If .Cells(i, m).Interior.Color = 65535 Then
                                total = total + .Cells(i, m)
                                h = h + 1
                            End If
                        Next
                    End If
                Next
                If h > 0 Then
                    .Cells(j, k) = total / h
                End If
            Next
        Next
        
        endRow = .Cells(12, 5).End(xlDown).Row
        .Range("P3:R11").ClearContents
        For k = 16 To 18
            For j = 3 To 11
                key = .Range("J" & j) & "-" & .Cells(2, k)
                total = 0
                h = 1
                ReDim arr(1 To h)
                For i = 13 To endRow
                    If key = .Range("F" & i) & "-" & .Range("I" & i) Then
                        For m = 15 To 24
                            If .Cells(i, m).Interior.Color = 65535 Then
                                ReDim Preserve arr(1 To h)
                                arr(UBound(arr)) = .Cells(i, m)
                                h = h + 1
                            End If
                        Next
                    End If
                Next
                Debug.Print UBound(arr)
                If UBound(arr) > 0 Then
                    For m = 1 To UBound(arr) - 1
                        For n = m + 1 To UBound(arr)
                            If arr(m) > arr(n) Then
                                temp = arr(n)
                                arr(n) = arr(m)
                                arr(m) = temp
                            End If
                        Next
                    Next
                    If UBound(arr) Mod 2 = 1 Then
                        .Cells(j, k) = arr(UBound(arr) \ 2 + 1)
                    Else
                        .Cells(j, k) = (arr(UBound(arr) \ 2) + arr(UBound(arr) \ 2 + 1)) / 2
                    End If
                End If
            Next
        Next
    End With
    MsgBox "complete!"
End Sub

Sub insertData()
    Dim strConn As String
    Dim dbConn As Object
    Dim resSet As Object
    Dim wk As Worksheet
    Dim endRow&, i&, j&, k&, m&, n&, total, h&, key, ky, arr(), temp
    Dim dic As Object, st
    Dim db_sid, db_user, db_pass As String, qsql$
    Set wk = ThisWorkbook.Worksheets("data")
    Set dic = CreateObject("Scripting.Dictionary")
    
    db_sid = "VBAcase55"
    db_user = "RACEDATA"
    db_pass = "RACEDATA"
      
    Set dbConn = CreateObject("ADODB.Connection")
    Set resSet = CreateObject("ADODB.Recordset")
    strConn = "Provider=OraOLEDB.Oracle.1; user id=" & db_user & "; password=" & db_pass & "; data source = " & db_sid & "; Persist Security Info=True"
'    strConn = "Provider=MSDAORA.1; user id=" & db_user & "; password=" & db_pass & "; data source = " & db_sid & "; Persist Security Info=True"
        
    dbConn.Open strConn
    
    With wk
        endRow = .Cells(12, 5).End(xlDown).Row
        For i = 13 To endRow
            key = .Range("E" & i).Value
            If Not dic.exists(key) Then
                dic.Add key, ""
            End If
        Next
        For Each key In dic.keys
            h = 1
            ReDim arr(1 To h)
            st = 0
            For i = 13 To endRow
                If .Range("E" & i) = key Then
                    For m = 15 To 24
                        If .Cells(i, m).Interior.Color = 65535 Then
                            ReDim Preserve arr(1 To h)
                            arr(UBound(arr)) = .Cells(i, m)
                            h = h + 1
                        End If
                    Next
                End If
            Next
            If UBound(arr) > 1 Then
                st = Application.WorksheetFunction.StDev(arr)
                dbConn.BeginTrans
                qsql = "UPDATE post_variant SET post_this_variant =" & st & "WHERE running_raceno = " & key - 169
                dbConn.Execute qsql
                dbConn.CommitTrans
            End If
        Next
    End With
    dbConn.Close
    dic.RemoveAll
    Set dic = Nothing
    MsgBox "complete!"
End Sub
