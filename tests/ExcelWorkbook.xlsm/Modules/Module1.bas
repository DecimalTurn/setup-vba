Attribute VB_Name = "Module1"

'@Lang VBA
Option Explicit

Sub Demo()
    ' FormulaLocal syntax (e.g. list separator) depends on the Excel UI language
    ThisWorkbook.Sheets("Sheet1").Range("A1").FormulaLocal = "=SUM(1,2,3)"

    MsgBox "Hello, World!"
End Sub

'This code will be called via COM to test if the VBA import was successful
Sub WriteToFile()
    Dim filePath As String
    Dim fileNum As Integer
    
    ' Specify the path to the text file
    filePath = ThisWorkbook.Path & "\ExcelWorkbook.txt"
    
    ' Get a free file number
    fileNum = FreeFile
    
    ' Open the file for output
    Open filePath For Output As #fileNum
    
    ' Write some text to the file
    Print #fileNum, "Hello, World!"
    
    ' Close the file
    Close #fileNum
End Sub

'This code is called via COM by the custom test workflow to check for locale/language issues
Sub TestLocaleFormula()
    Dim filePath As String
    Dim fileNum As Integer

    ' Enter the locale-dependent formula before reading its result
    Demo

    ' Specify the path to the text file
    filePath = ThisWorkbook.Path & "\ExcelWorkbook_LocaleTest.txt"

    ' Get a free file number
    fileNum = FreeFile

    ' Open the file for output
    Open filePath For Output As #fileNum

    ' Write the value computed from the cell where the formula was edited
    Print #fileNum, ThisWorkbook.Sheets("Sheet1").Range("A1").Value

    ' Close the file
    Close #fileNum
End Sub