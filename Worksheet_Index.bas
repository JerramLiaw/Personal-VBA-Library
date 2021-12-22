Attribute VB_Name = "Worksheet_Index"
Option Explicit

Sub Worksheet_Index()
'This macro works by creating hyperlinks on the main page and hyperlinks to return back to the main page

'   Declare macro level variables
    Dim i, j As Integer
    Dim Sheet_Name As String
    Dim Sheet As Worksheet
    Dim Switch As Boolean
        Sheet_Name = "WorksheetIndex"
        j = 2

'   Check to see if the Worksheet Index already exists
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = Sheet_Name Then
            Switch = True
        Else
            Switch = False
        End If
    Next i
    
'   If it exists, initialize it. Else, create and format it
    If Switch = True Then
        Sheets("WorksheetIndex").Range("A2:B500").ClearContents
    Else
        Worksheets.Add.Name = Sheet_Name
        With Worksheets(Sheet_Name)
            .Range("A1") = "Worksheet Link"
            .Range("B1") = "Worksheet Description"
            .Range("C1") = "Additional Information"
            .Rows(1).Interior.ColorIndex = 23
            .Rows(1).Font.Color = vbWhite
            .Rows(1).Font.Bold = True
            .Columns("A").ColumnWidth = 25
            .Columns("B").ColumnWidth = 50
            .Columns("C").ColumnWidth = 50
        End With
    End If

'   Creates a hyperlink on the Index page to each worksheet available
    For Each Sheet In Worksheets
    If Sheet.Name <> "WorksheetIndex" And Sheet.Visible <> xlSheetHidden Then
        Sheets("WorksheetIndex").Range("A" & j).Value = Sheet.Name
        Sheets("WorksheetIndex").Hyperlinks.Add _
            Anchor:=Range("A" & j), _
            Address:="", _
            SubAddress:=Sheet.Name & "!A1", _
            ScreenTip:="Go to " & Sheet.Name
        j = j + 1
    End If
    Next Sheet
    
'   Insert a new row and insert a hyperlink to return back to the Index page for quick navigation
    For Each Sheet In Worksheets
    If Sheet.Name <> "WorksheetIndex" And Sheet.Visible <> xlSheetHidden Then
        Sheet.Rows(1).Insert
        Sheet.Range("A1").Value = "Return to Index"
        Sheet.Hyperlinks.Add _
            Anchor:=Sheet.Range("A1"), _
            Address:="", _
            SubAddress:="WorksheetIndex!A1", _
            ScreenTip:="Return to Index Sheet"
    End If
    Next Sheet
  
End Sub
