Attribute VB_Name = "Version_Control"
Option Explicit

'   Declare Module level variables
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim VBA_Repo As String
    Dim File_Path As String
    Dim VB_Comp As VBIDE.VBComponent

Sub Export_Modules()
'   The objective is to export all VBA modules into a local repo which will be managed by Git in PowerShell.

'   Declaring all macro level variables
    Dim Export_Switch As Boolean
    Dim Export_Name As String

'   Setting variable values
    Set objFSO = New Scripting.FileSystemObject
    VBA_Repo = "C:\Users\Jerram\OneDrive - Singapore Management University\Desktop\Business Intelligence\VBA\VBA Github Repo"
    File_Path = VBA_Repo & "\"

'   Initialize the Repo folder by deleting all module files inside it
    For Each objFile In objFSO.GetFolder(File_Path).Files
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            Kill objFile.Path
        End If
    Next objFile

'   Iterate through each VBA Component inside our Workbook
    For Each VB_Comp In ThisWorkbook.VBProject.VBComponents

'       Switch to signal whether or not we want to export the component
        Export_Switch = True
'       Set the name of the export to be the component name
        Export_Name = VB_Comp.Name

'       Set the file extension based on their type; worksheet modules are not exported for simplicity reasons
        Select Case VB_Comp.Type
            Case vbext_ct_ClassModule
                Export_Name = Export_Name & ".cls"
            Case vbext_ct_MSForm
                Export_Name = Export_Name & ".frm"
            Case vbext_ct_StdModule
                Export_Name = Export_Name & ".bas"
            Case vbext_ct_Document
                Export_Switch = False
        End Select

'       Exports all non-worksheet modules into the Repo folder
        If Export_Switch = True Then
            VB_Comp.Export File_Path & Export_Name
            If Export_Name <> "Version_Control.bas" Then
                ThisWorkbook.VBProject.VBComponents.Remove VB_Comp
            End If
        End If
    Next VB_Comp

'   Pop-up indicator when the macro is complete
    MsgBox ("Export completed.")

End Sub

Sub Import_Modules()
'   The objective is to import all VBA modules from the local repo into the current workbook for use

'   Setting variable values
    Set objFSO = New Scripting.FileSystemObject
    VBA_Repo = "C:\Users\Jerram\OneDrive - Singapore Management University\Desktop\Business Intelligence\VBA\VBA Github Repo"
    File_Path = VBA_Repo & "\"

'   Imports all module files into the workbook
    For Each objFile In objFSO.GetFolder(File_Path).Files
        If objFile.Name <> "Version_Control.bas" Then
            If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
                (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
                (objFSO.GetExtensionName(objFile.Name) = "bas") Then
                    ThisWorkbook.VBProject.VBComponents.Import objFile.Path
            End If
        End If
    Next objFile

'   Pop-up indicator when the macro is complete
    MsgBox ("Import completed.")
End Sub




 
