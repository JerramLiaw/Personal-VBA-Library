Attribute VB_Name = "Data_Clean"
Option Explicit

Sub Show_Everything()

'   Declare necessary variables
    Dim Sheet As Worksheet
    Dim Cell As Range
    Dim Merged_Cells As Range
    Dim Merged_Range As Range
 
'   Turn off events and updating to allow the macro to run faster
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

'   Loops through each sheet inside the workbook
    For Each Sheet In ThisWorkbook.Worksheets
        
'       Unhides all sheets
        If Sheet.Visible <> xlSheetVisible Then
            Sheet.Visible = xlSheetVisible
        End If
        
'       Unhides all rows/columns, autofits the data then unfilters it
        With Sheet
            .Cells.EntireColumn.Hidden = False
            .Cells.EntireRow.Hidden = False
            .Cells.EntireColumn.AutoFit
            .Cells.EntireRow.AutoFit
            
            On Error Resume Next
            .ShowAllData
        End With

'       Find all merged cells. If horizontal, change to center across selection. If vertical, fill values
        For Each Cell In Sheet.UsedRange
            If Cell.MergeCells = True And Cell.MergeArea.Rows.Count = 1 Then
                Merged_Range = Cell.MergeArea
                Merged_Range.Unmerge
                Merged_Range.HorizontalAlignment = xlCenterAcrossSelection
            Else
                Merged_Range = Cell.MergeArea
                Merged_Range.Unmerge
                Merged_Range.Formula = Cell.Formula
            End If
        Next Cell
    Next Sheet
    
'   Reset excel back to original state
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub
Sub Find_Missing_Data()

'   Declare necessary variables
    Dim Sheet As Worksheet
    Dim Cell As Range

'   Conditionally formats used range to highlight blank cells
    With ActiveSheet.UsedRange
        .FormatConditions.Delete
        .FormatConditions.Add _
            Type:=xlBlanksCondition
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.399945066682943
        End With
    End With

End Sub
Sub Remove_Error()

'   Declare necessary variables
    Dim Cell As Range
    Dim Current_Formula As String
    Dim Replacement As String
   
 '  Set what to replace error with
    Replacement = InputBox("What value to show?")
    
'   Replace all formulas with IFERROR formula
    For Each Cell In Selection.Cells
        If Cell.HasFormula = True And Cell.HasArray = False Then
            Current_Formula = Right(Cell.Formula, Len(Cell.Formula) - 1) '-1 to remove the equal sign
            Cell.Formula = "=IFERROR(" & Current_Formula & "," & Replacement & ")"
        End If
    Next Cell
End Sub

