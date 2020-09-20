Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub CommandButton1_Click()
    n = CSng(samplingFrequency.Value) * CSng(testPeriod.Value)
    m = CInt(numofData.Value)
    Unload InputData
End Sub


Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit

Sub BrowseFile()

    Dim fd As FileDialog
    Dim vrtSelectedItem As Variant
    Dim textline As String
    'vrtSelectedItem must be variant because of For Each loop
    Dim i As Integer

    'Allow the user to select multiple files
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .AllowMultiSelect = True
        .Filters.Add "Text", "*.txt", 1
        If .Show = -1 Then
            Sheets(1).name = "Data File"
            For Each vrtSelectedItem In .SelectedItems
                i = i + 1
                Sheets(1).Cells(i + 1, 1) = vrtSelectedItem
                
                ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
                WriteFile (vrtSelectedItem)
                ActiveSheet.name = SheetName(vrtSelectedItem)
                
            Next vrtSelectedItem
            Sheets(1).Select
            Range("A1").Value = "Data Links as below:"
        Else
        End If
    End With
    Set fd = Nothing
    
End Sub

Sub WriteFile(vrtSelectedItem As Variant)
    Dim textline As String
    Dim j As Single
    Open vrtSelectedItem For Input As #1
                Do Until EOF(1)
                    Line Input #1, textline
                    Cells(j + 1, 1).Value = textline
                    j = j + 1
                Loop
                Close #1
End Sub

Function SheetName(name As Variant) As String
    Dim flashPos As Integer
    Dim dotPos As Integer
    Dim stringName As String
    
    stringName = CStr(name)
    dotPos = InStr(stringName, ".txt")
    flashPos = InStrRev(stringName, "\")
    
    SheetName = Mid(stringName, flashPos + 1, dotPos - flashPos - 1)
    
End Function

Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit
Public m As Integer
Public n As Single
'n is the number of samples
'm is the number of accelerometer

Sub ExportData()
    
    InputData.Show
    If Range("B11").Value = "" Then
        Call DivideData
    End If
    
    Call ArrangeData(n, m)
    Call WriteToTxt(n, m)
    
End Sub

Sub ArrangeData(n As Single, m As Integer)
    
    Dim i As Integer
    Dim rowIndex As Single
    
    Range(Cells(1, 1), Cells(n, 1)).Value = _
                Range(Cells(11, 2), Cells(n + 10, 2)).Value
    For i = 1 To m
        Range(Cells(1, i + 1), Cells(n, i + 1)).Value = _
                Range(Cells(n * (i - 1) + 11 * i, 3), Cells((n + 11) * i - 1, 3)).Value
    Next
    rowIndex = (n + 11) * m + 9
    Range(Cells(n + 1, 1), Cells(rowIndex, 5)).ClearContents
    Range("A1").Select
    
End Sub

Sub WriteToTxt(n As Single, m As Integer)

    Dim myFile As String, rng As Range
    Dim cellValue As String
    Dim i As Single
    Dim j As Single
    Dim length As Integer
    Dim SheetName As String
        
    SheetName = ActiveSheet.name
    myFile = Application.ActiveWorkbook.Path & "\MAT " & SheetName & ".txt"
    Set rng = Range(Cells(1, 1), Cells(1, 1).End(xlDown).Offset(0, m))
    
    Open myFile For Output As #1
    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            cellValue = CStr(rng.Cells(i, j).Value)
            length = Len(cellValue)
            Select Case j
                Case 1
                    Print #1, cellValue; Spc(10 - length);
                Case rng.Columns.Count
                    Print #1, cellValue
                Case Else
                    Print #1, cellValue; Spc(20 - length);
            End Select
        Next j
    Next i
    Close #1
    
End Sub

Sub DivideData()
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1)), TrailingMinusNumbers:=True
    Rows("1:1").EntireRow.Insert
End Sub
