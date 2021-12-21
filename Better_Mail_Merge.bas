Attribute VB_Name = "Better_Mail_Merge"
Option Explicit

'   Declare Module Level Variables
    Dim Num_Attach As Integer
    Dim Num_Merge As Integer
    Dim File_Path As String
    Dim Sheet_Name As String

Sub Create_Template()
'   The purpose of this Macro is to create a template that will be used for the Mailmerge macro

'   Declare Macro level Variables
    Dim Excel_Obj As Excel.Application
    Dim Excel_Workbook As Excel.Workbook
    Dim i, j, k As Integer
    
'   Defining static parameters
    Num_Attach = InputBox("How many attachments are there?")
    Num_Merge = InputBox("How many merging fields are there?")
    File_Path = "C:\Users\Jerram\OneDrive - Singapore Management University\Desktop\Business Intelligence\VBA\MailMerge.xlsx"
    Sheet_Name = "MailMerge"

'   Defining Excel Object Variables
    Set Excel_Obj = CreateObject("Excel.Application")
    Set Excel_Workbook = Excel_Obj.Workbooks.Add(xlWBATWorksheet)
    Excel_Workbook.Sheets(1).Name = Sheet_Name
    
'   Create email parameter cells
    With Excel_Workbook.Sheets(Sheet_Name)
        .Range("A1").Value = "Mail Merge Parameters"
        .Range("A2").Value = "Carbon Copy (CC)"
        .Range("A3").Value = "Blind Carbon Copy (BCC)"
        .Range("A4").Value = "Subject"
        .Range("A5").Value = "Number of attachments"
            .Range("B5").Value = Num_Attach
        .Range("A6").Value = "Number of Merges"
            .Range("B6").Value = Num_Merge
        .Range("A7").Value = "Send Immediately?"
    End With

'   Format email parameter cells
    With Excel_Workbook.Sheets(Sheet_Name)
        .Range("A1:A7").Font.Bold = True
        .Range("A1").Font.Color = vbWhite
        .Range("A1:B1").HorizontalAlignment = xlHAlignCenterAcrossSelection
        .Range("A2:A7").HorizontalAlignment = xlHAlignCenter
        .Range("B2:B7").HorizontalAlignment = xlHAlignCenter
        .Range("A1:B1").Interior.ColorIndex = 23
        .Range("A1:B7").Borders.LineStyle = xlContinuous
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 50
    End With

'   Create Mail Merge parameter cells. More should be added by the user as needed
    With Excel_Workbook.Sheets(Sheet_Name)
        .Range("A9").Value = "Recipient E-mail Address"
        .Range("B9").Value = "Recipient Name"
    End With
    
'   Set the number of attachment columns
    For i = 1 To Num_Attach
        With Excel_Workbook.Sheets(Sheet_Name)
            .Cells(9, i + 2) = "Attachment Path " & i
        End With
    Next i

'   Set the number of merged columns. Put after attachments as merge > attachments
    For j = 1 To Num_Merge
        With Excel_Workbook.Sheets(Sheet_Name)
            .Cells(9, j + 2 + Num_Attach) = "Merge Field " & j
        End With
    Next j
                
'   Format Mail Merge parameter cells
    With Excel_Workbook.Sheets(Sheet_Name)
        .Rows(9).Interior.ColorIndex = 23
        .Rows(9).Font.Color = vbWhite
        .Rows(9).Font.Bold = True
        .Rows(9).HorizontalAlignment = xlHAlignCenter
    End With

'   Format attachment and merged columns
    For k = 3 To 2 + Num_Attach + Num_Merge
        With Excel_Workbook.Sheets(Sheet_Name)
            .Columns(k).ColumnWidth = 20
        End With
    Next k

'   Visual confirmation that process was completed
    Excel_Obj.Visible = True
    
'   Delete any previous files and then save the new template there
    If Len(Dir$(File_Path)) > 0 Then
        Kill File_Path
    End If
    Excel_Workbook.SaveAs _
        FileName:=File_Path

'   Close all instances of Excel and the workbook
    Excel_Workbook.Close False
    Set Excel_Workbook = Nothing
    Excel_Obj.Quit
    Set Excel_Obj = Nothing
    
    
End Sub

Sub Better_Mail_Merge()

'   Declare macro level variables
    Dim Excel_Obj As Excel.Application
    Dim Excel_Workbook As Excel.Workbook
    Dim Last_Row As Integer
    Dim Outlook_Obj As Object
    Dim Outlook_Mail As Object
    Dim Signature As String
    Dim Attachment_Path As String
    Dim Switch As Integer
    Dim x, y, z As Integer

'   Defining the static parameters
    Sheet_Name = "MailMerge"
    File_Path = "C:\Users\Jerram\OneDrive - Singapore Management University\Desktop\Business Intelligence\VBA\MailMerge.xlsx"

'   Defining Excel and Outlook Object Variables
    Set Excel_Obj = CreateObject("Excel.Application")
    Set Excel_Workbook = Excel_Obj.Workbooks.Open(File_Path)
    Set Outlook_Obj = CreateObject("Outlook.Application")
    
'   Extract key information from excel workbook
    Last_Row = Excel_Workbook.Sheets(Sheet_Name).Range("A" & Rows.Count).End(xlUp).Row
    Num_Attach = Excel_Workbook.Sheets(Sheet_Name).Range("B5").Value
    Num_Merge = Excel_Workbook.Sheets(Sheet_Name).Range("B6").Value
    Switch = Excel_Workbook.Sheets(Sheet_Name).Range("B7").Value

'   Starts the mail merge process based on cells with an email address filled in
    On Error Resume Next
    For x = 10 To Last_Row
        If Cells(x, 1).Value Like "?*@?*" Then
            Set Outlook_Mail = Outlook_Obj.CreateItem(0)
           
'           Opens an instance of an email template and extracts the signature
            Outlook_Mail.Display
            Signature = Outlook_Mail.Body
            
'           Fills in the fixed areas of the email
            With Outlook_Mail
                .To = Cells(x, 1)
                .CC = Cells(2, 2)
                .BCC = Cells(3, 2)
                .Subject = Cells(4, 2)
                .Body = ThisDocument.Content & vbNewLine & Signature
                .Body = Replace(.Body, "<Name>", Excel_Workbook.Sheets(Sheet_Name).Cells(x, 2))

'               Attaches files to the email
                For y = 1 To Num_Attach
                    Attachment_Path = Cells(x, y + 2)
                    .Attachments.Add Attachment_Path
                Next y
                    
'               Changes the merged fields according to the Excel
                For z = 1 To Num_Merge
                    .Body = Replace(.Body, "<" & z & ">", Excel_Workbook.Sheets(Sheet_Name).Cells(x, z + 2 + Num_Attach))
                Next z

'               Immediately sends the email only if the switch is set to 1
                If Switch = 1 Then
                    .Send
                End If
 
            End With
        End If
    Next x

'   Close all instances of Excel and Outlook
    Excel_Workbook.Close False
    Set Excel_Workbook = Nothing
    Excel_Obj.Quit
    Set Excel_Obj = Nothing
    Set Outlook_Mail = Nothing
    Set Outlook_Obj = Nothing
    
End Sub
