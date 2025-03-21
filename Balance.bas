Attribute VB_Name = "Module1"
Sub occurence()
    Dim RowI As Integer, EachI As Integer, I As Integer, Counts As Integer, Countv As Integer, Countg As Integer
    Dim Status As String, FillCell As String, Account As String
    Dim Fillables As Boolean, Fillablev As Boolean, Fillableg As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    I = 1
    While I <= lastRow
        If ws.Rows(I).Interior.Color = RGB(255, 0, 0) Then
            ws.Rows(I).Interior.Color = RGB(255, 255, 255)
        End If
        If ws.Cells(I, 5).Value <> "" Then
            ws.Cells(I, 5).Value = ""
        End If
        If ws.Cells(I, 6).Value <> "" Then
            ws.Cells(I, 6).Value = ""
        End If
        I = I + 1
    Wend

    RowI = 2

    While RowI <= lastRow
        If ws.Cells(RowI, 5).Value = "" Then
            EachI = RowI
            Counts = 0
            Countv = 0
            Countg = 0
            Fillables = False
            Fillablev = False
            Fillableg = False
            FillCell = ""
            Account = ws.Cells(RowI, 1).Value
            While EachI <= lastRow And ws.Cells(EachI, 5).Value = ""
                If ws.Cells(EachI, 1).Value = Account Then
                    ws.Cells(RowI, 5).Value = "Checked"
                    If ws.Cells(EachI, 2).Value = "Standard" And ws.Cells(EachI, 3).Value < 0 Then
                        Counts = Counts + 1
                        ws.Cells(EachI, 6).Value = "Fills"
                        If Counts > 3 Then
                            Fillables = True
                        End If
                    End If
                    If ws.Cells(EachI, 2).Value = "VIP" And ws.Cells(EachI, 3).Value < -100 Then
                        Countv = Countv + 1
                        ws.Cells(EachI, 6).Value = "Fillv"
                        If Countv > 5 Then
                            Fillablev = True
                        End If
                    End If
                    If ws.Cells(EachI, 2).Value = "Golden" And ws.Cells(EachI, 3).Value < -500 Then
                        Countg = Countg + 1
                        ws.Cells(EachI, 6).Value = "Fillg"
                        If Countg > 10 Then
                            Fillableg = True
                        End If
                    End If
                End If
                EachI = EachI + 1
            Wend
            I = 2
            While I <= lastRow
                If ws.Cells(I, 1).Value = Account Then
                    If Fillables = True And ws.Cells(I, 6).Value = "Fills" Then
                        ws.Cells(I, 1).CurrentRegion.Rows(I).Interior.Color = RGB(255, 0, 0)
                    End If
                    If Fillablev = True And ws.Cells(I, 6).Value = "Fillv" Then
                        ws.Cells(I, 1).CurrentRegion.Rows(I).Interior.Color = RGB(255, 0, 0)
                    End If
                    If Fillableg = True And ws.Cells(I, 6).Value = "Fillg" Then
                        ws.Cells(I, 1).CurrentRegion.Rows(I).Interior.Color = RGB(255, 0, 0)
                    End If
                End If
                I = I + 1
            Wend
        End If
        RowI = RowI + 1
    Wend

    I = 1
    While I <= lastRow
        ws.Cells(I, 5).Interior.Color = xlNone
        ws.Cells(I, 5).ClearContents
        ws.Cells(I, 6).Interior.Color = xlNone
        ws.Cells(I, 6).ClearContents
        I = I + 1
    Wend

End Sub
