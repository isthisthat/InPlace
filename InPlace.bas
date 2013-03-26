Attribute VB_Name = "InPlace"
Option Explicit
' Macros updated March 2013
' (c) 2007 by Stathis Kanterakis

Sub AlignInPlace()
Attribute AlignInPlace.VB_Description = "This is the description"
Attribute AlignInPlace.VB_ProcData.VB_Invoke_Func = " \n14"

Dim compA As String, compB As String, compA_cols As String, compB_cols As String, strA As String, strB As String
Dim start_str As String, msg_txt As String, msg_ret As String
Dim curr As Integer, start As Integer

compA = InputBox("Enter the first comparison column" & vbCrLf & "This should be the first column of the first table you wish to align and should contain sorted ids", "AlignInPlace", "A")
If Len(compA) = 0 Then: Exit Sub
compA_cols = InputBox("Enter the first range column" & vbCrLf & "This should be the last column of your first table", "AlignInPlace", compA)
If Len(compA_cols) = 0 Then: Exit Sub
compB = InputBox("Enter second comparison column" & vbCrLf & "This should be the first column of the second table you wish to align and should contain sorted ids", "AlignInPlace", rowToColumn(columnToRow(compA_cols) + 2))
If Len(compB) = 0 Then: Exit Sub
compB_cols = InputBox("Enter second range column" & vbCrLf & "This should be the last column of your second table", "AlignInPlace", compB)
If Len(compB_cols) = 0 Then: Exit Sub
start_str = InputBox("Starting row" & vbCrLf & "Enter the row where your data starts. For example, if you have headers in the first row, enter 2", "AlignInPlace", "2")
If Len(start_str) = 0 Then: Exit Sub

start = val(start_str)
curr = start

Application.StatusBar = "AlignInPlace is running..."

msg_txt = "I will align the ids of the first table (column " & UCase(compA) & " to column " & UCase(compA_cols) & ")" & vbCrLf & _
"against the ids of the second table (column " & UCase(compB) & " to column " & UCase(compB_cols) & ")," & vbCrLf & _
"starting at row " & curr & "." & vbCrLf & _
"This operation cannot be undone." & vbCrLf & "Continue?"
msg_ret = MsgBox(msg_txt, vbYesNoCancel, "AlignInPlace")
If msg_ret <> vbYes Then: Exit Sub

strA = Trim(Range(compA & curr))
strB = Trim(Range(compB & curr))

Range(compA & start).Select
    
Do While Len(strA) > 0 And StrComp(strA, "", vbTextCompare) <> 0
    If Len(strB) > 0 And StrComp(strB, "", vbTextCompare) <> 0 Then
        If IsNumeric(strA) And IsNumeric(strB) Then
            If CDbl(strA) < CDbl(strB) Then
                Range(compB & curr & ":" & compB_cols & curr).Select
                Selection.Insert Shift:=xlDown
            ElseIf CDbl(strA) > CDbl(strB) Then
                Range(compA_cols & curr & ":" & compA & curr).Select
                Selection.Insert Shift:=xlDown
            End If
        Else
            If StrComp(strA, strB, vbTextCompare) < 0 Then
                Range(compB & curr & ":" & compB_cols & curr).Select
                Selection.Insert Shift:=xlDown
            ElseIf StrComp(strA, strB, vbTextCompare) > 0 Then
                Range(compA_cols & curr & ":" & compA & curr).Select
                Selection.Insert Shift:=xlDown
            End If
        End If
    End If
    curr = curr + 1
    strA = Trim(Range(compA & curr))
    strB = Trim(Range(compB & curr))
Loop
Range(compA & start).Select

Application.StatusBar = False

End Sub

Sub MatchInPlace()
Dim compA As String, compB As String, compA_cols As String, compB_cols As String, strA As String, strB As String, maxRow As String
Dim start_str As String, msg_txt As String, msg_ret As String
Dim curr As Integer, start As Integer, max As Integer

compA = InputBox("Enter template column" & vbCrLf & "This is a column that contains blank rows which you wish to match in another table", "MatchInPlace", "A")
If Len(compA) = 0 Then: Exit Sub
compB = InputBox("Enter target column" & vbCrLf & "This is the first column of the table you wish to match to the template column", "MatchInPlace", rowToColumn(columnToRow(compA) + 2))
If Len(compB) = 0 Then: Exit Sub
compB_cols = InputBox("Enter target range column" & vbCrLf & "This is the last column of the target table", "MatchInPlace", compB)
If Len(compB_cols) = 0 Then: Exit Sub
start_str = InputBox("Starting row" & vbCrLf & "Enter the row where your data starts. For example, if you have headers in the first row, enter 2", "AlignInPlace", "2")
If Len(start_str) = 0 Then: Exit Sub
start = val(start_str)
curr = start

max = Range(compA & ":" & compA).SpecialCells(xlCellTypeLastCell).row
maxRow = InputBox("End row" & vbCrLf & "Enter the last row that contains data. An estimate has been given by default.", "AlignInPlace", max)
If Len(maxRow) = 0 Then: Exit Sub
max = val(maxRow)

Application.StatusBar = "MatchInPlace is running..."

msg_txt = "I will insert blank rows to the target table (column " & UCase(compB) & " to column " & UCase(compB_cols) & ")," & vbCrLf & _
"to match the template (column " & UCase(compA) & "), from row " & curr & " to row " & max & "." & vbCrLf & _
"This operation cannot be undone." & vbCrLf & "Continue?"
msg_ret = MsgBox(msg_txt, vbYesNoCancel, "MatchInPlace")
If msg_ret <> vbYes Then: Exit Sub

strA = Trim(Range(compA & curr))
strB = Trim(Range(compB & curr))

Range(compA & start).Select

Do While Len(strB) > 0 And curr < max
    If Len(strA) < 1 Then
        Range(compB & curr & ":" & compB_cols & curr).Select
        Selection.Insert Shift:=xlDown
    End If
    curr = curr + 1
    strA = Trim(Range(compA & curr))
    strB = Trim(Range(compB & curr))
Loop
Range(compA & start).Select
Application.StatusBar = False

End Sub

Function rowToColumn(row As Integer) As String
    Dim a, b, r As Integer
    r = row - 1
    a = r \ 26
    b = r Mod 26
    If r > 25 Then
        rowToColumn = Chr(a + 64) & Chr(b + 65)
    Else
        rowToColumn = Chr(b + 65)
    End If
End Function

Function columnToRow(column As String) As Integer
    column = StrConv(column, vbUpperCase)
    If Len(column) = 1 Then
        columnToRow = Asc(column) - 64
    ElseIf Len(column) = 2 Then
        columnToRow = (Asc(Mid(column, 1, 1)) - 64) * 26 + Asc(Mid(column, 2, 1)) - 64
    End If
End Function

