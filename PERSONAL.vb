Sub ConditionalFormat()
    ' Format cells based on contents
    Dim SrchRng As Range
    Dim cel As Range

    Set SrchRng = Range("A1:B10") ' Change this
    For Each cel in SrchRng
        If InStr(1, cel.Value, "Cell Contents") > 0 _
        Or InStr(1, cel.Value, "Cell Contents Other") > 0 _
        Then
            cel.Interior.Color = RGB(255, 0, 0)
        End If
    Next cel
    End Sub


Sub GenerateRandomNumbers()
    ' Generate random numbers using a seed value
    ' Select the cells you wish to populate with random numbers, then run macro
    Low = Application.InputBox("Enter first valid value", Type:=1)
    High = Application.InputBox("Enter last valid value", Type:=1)
    Selection.Clear
    Rnd (-17) ' This is the initial seed value

    For Each cell In Selection.Cells
        If WorksheetFunction.CountA(Selection) = (High - Low + 1) Then Exit For
        Do
            rndNumber = Int((High - Low + 1) * Rnd() + Low)
            Loop Until Selection.Cells.Find(rndNumber, LookIn:=xlValues, lookat:=xlWhole) Is Nothing
        cell.Value = rndNumber
        Next
    End Sub


Sub SelectRows()
    ' Select the entire row for each active cell
    ' Select the cells whose rows you wish to select, then run macro
    Selection.EntireRow.Select
    End Sub


Sub MergeXLS()
    ' Combine multiple Excel (xlsx) files in a given directory
    ' Change path and range within this subroutine
    Dim bookList As Workbook
    Dim mergeObj As Object, dirObj As Object, filesObj As Object, everyObj As Object
    Application.ScreenUpdating = False
    Set mergeObj = CreateObject("Scripting.FileSystemObject")
 
    ' Change folder path of excel files here
    Set dirObj = mergeObj.Getfolder("C:\insert\path\here")
    Set filesObj = dirObj.Files
    For Each everyObj In filesObj
    Set bookList = Workbooks.Open(everyObj)
 
    ' Make sure "A" column on "A65536" is the same column as start point
    Range("A2:IV" & Range("A65536").End(xlUp).Row).Copy
    ThisWorkbook.Worksheets(1).Activate
 
    ' Do not change the following column. It's not the same column as above
    Range("A65536").End(xlUp).Offset(1, 0).PasteSpecial
    Application.CutCopyMode = False
    bookList.Close
    Next
    End Sub


Sub OpenHyperlink()
    ' Follow hyperlinks with ctrl+d
    ' Reload Excel or run Auto_Open and use CTRL-D to follow hyperlink of the active cell
    Application.OnKey "^d", "GoToHyperlink"
    End Sub
 
Private Sub GoToHyperlink()
    Dim i As Long, s As String
    With ActiveCell
        s = .Formula
        If s Like "=HYPERLINK(*" Then
            s = Mid(Split(s, ",")(0), 12)
            If Right(s, 1) = ")" Then s = Left(s, Len(s) - 1)
                s = .Worksheet.Evaluate(s)
                i = InStrRev(s, ".xl", , vbTextCompare)
            If i >= Len(s) - 4 Then
                ' HYPERLYNK formula is referenced to workbook - open in without warning
                Workbooks.Open s
            Else
                ' HYPERLYNK formula is referenced to something else
                ActiveWorkbook.FollowHyperlink Address:=s
            End If
        ElseIf .Hyperlinks.Count Then
            ' Hyperlink is inserted into the active cell
            .Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
            End If
        End With
    End Sub


Sub MakeHyperlink()
    ' Converts each text hyperlink selected into a working hyperlink
    For Each xCell In Selection
        ActiveSheet.Hyperlinks.Add Anchor:=xCell, Address:=xCell.Formula
    Next xCell
    End Sub


Function GetCellURL(rng As Range) As String
    ' Fetch as text the URL reference of a cell value
    On Error Resume Next
    GetCellURL = rng.Hyperlinks(1).Address
    End Function

