Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "20"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "30"
    Range("D5").Select
    ActiveCell.FormulaR1C1 = "40"
    Range("E7").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E6").Select
    ActiveCell.FormulaR1C1 = "50"
    Range("E7").Select
End Sub

Sub Concept1()
    'these are comments, notes that you leave in your code
    MsgBox ("Hello Everyone")
End Sub

Sub Concept2()
    'x=2
    '2 + x = ?
    X = 2
    MsgBox (2 + X)
End Sub

Sub Concept3()
    'In programming, values have diff data types
    'The main data types are words and nums
    'Words--->Strings, Numbers--->Long
    Dim X As Long
    'I am creating a variable called X, and X must be a Long(Number)
    X = "Hello"
    'X = 2
    MsfBox (2 + X)
End Sub

Sub Concept4a()
    'Cells and Ranges
    'Read and write data into specific cells
    Range("A1").Select
    Range("A2:B10").Select
    
    'Write data into cells
    'write the number 30 into cell A1
    Range("A1") = 30
    Range("A2") = "hello"
    Range("A3:B5") = 20
    
    'read date from cells
    X = Range("A1")
    MsgBox (X)
    
    
    'read the data from cell A3 using the Cells Object
    Y = Cells(3, 1) 'row 3,column 1
    MsgBox (Y)
    
End Sub

Sub Selfprac()
    Range("D1").Select
    Range("D1") = 30
    Range("A1:C30").Select
    Range("A1:C30") = 30
    X = Range("B1") 'X = Cells(1,2)
    MsgBox (X)
    
End Sub

Sub Concept4b()
    'Write the value 50 into Cell F1
    Range("F1") = 50
    'Write the value 40 into Cell G1
    Range("G1") = 40
    'Read the values in F1 and G1, and sum them up WRITE IN h1
    Range("H1") = Range("F1") + Range("G1")
    'Write formulas in H2, formulas are written using strings
    Range("H2") = "=F1+G1"
    
End Sub

Sub Concept4c()
    'join strings - concatenation - use &
    Range("F3") = "Hello"
    Range("G3") = "World"
    'Read the values in F3 and G3, join the text, and write it in H3
    Range("H3") = Range("F3") + Range("G3")
    'Write formulas
    Range("H4") = "=F3&G3"
End Sub

Sub Concept5()
    'Range("A1") = 100 'code run on the workbook u last clicked
    'writing 100 into cell A1 of the active workbook
    'active workbook -> last clicked workbook
    
    'write data into specific workbook,sheet and cell
    Workbooks("Workbook1.xlsx").Sheets("Sheet3").Range("A1") = 100
    'assign objects to a variable, using Set
    Set wb = Workbooks("Workbook1.xlsx")
    wb.Sheets("Sheet2").Range("B1") = 50
    Set ws = Workbooks("Workbook2.xlsx").Sheets("Sheet2")
    ws.Range("C3") = 100
    
End Sub

Sub Concept5a()
    'What happens if your workbook is closed?
    'Workbooks("Workbook2.xlsx").Sheets("Sheet1").Range("A1") = 50 cannot work
    Workbooks.Open("C:\Users\ganye\OneDrive - Nanyang Technological University\Workbook2.xlsx").Sheets("Sheet1").Range("A1") = 50
End Sub

Sub Concept6()
    'if control flow
    'if the value in cell A1 is bigger than 50, MsgBox("High") otherwise, MsgBox("Low")
    If Range("A1") > 50 Then
        MsgBox ("High")
    Else
        MsgBox ("Low")
    End If
    
End Sub

Sub Concept7()
    MsgBox (Rows.Count) 'this is the maximum number of rows for Excel
    LastRow = Workbooks("VBA Codes.xlsx").Sheets("Sheet2").Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox (LastRow)
    
End Sub

Sub Concept8()
    Set ws = Workbooks("VBA Codes.xlsx").Sheets("Sheet2")
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow
        ws.Cells(i, 3) = ws.Cells(i, 2) * ws.Cells(i, 1)
    Next i
End Sub
