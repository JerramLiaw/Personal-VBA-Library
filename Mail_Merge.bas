Attribute VB_Name = "Mail_Merge"
Option Explicit


Sub Mail_Merge()
'   The objective of this macro is to make an variation of the Microsoft Mail Merge function

'   Turn off other excel events to allow the macro to run faster
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

'   Defining Excel related objects
    Dim Sheet As Worksheet
    Dim Cell As Range
    Dim Body As String
    Dim Attachment1 As String
    Dim Attachment2 As String
        Set Sheet = Sheets("MailMerge")

'   Defining Outlook related Objects
    Dim Outlook_app As Object
    Dim Outlook_mail As Object
    Dim Signature As String
        Set Outlook_app = CreateObject("Outlook.Application")
        
    
'   Selecting only cells that contain email addresses to use. Email is used as it can be uniquely identified while the others are just strings
    For Each Cell In Sheet.Columns("B").Cells.SpecialCells(xlCellTypeConstants)
        If Cell.Value Like "?*@?*" Then
        
'           Create an Email template for each correct cell
            Set Outlook_mail = Outlook_app.CreateItem(0)
            
'           Create an instance of your default mail template to save your signature in the HTML.body property
            Outlook_mail.Display

'           Creating the body text using HTML syntax
            Body = "<BODY style=" & Chr(34) & "font-family: Calibri" & Chr(34) & ">" & _
                    "Dear " & Cells(Cell.Row, "A").Value & "," & _
                    "<p>" & Cells(Cell.Row, "C") & "</p>" & _
                    "<p>" & Cells(Cell.Row, "D") & "</p>" & _
                    "<p>" & Cells(Cell.Row, "E") & "</p>" & "</BODY>"
            
'           Create the attachment path and attach it. More attachments can be added by repeating the same process. Only two is included here
'           If Not statement is included to reduce pop-ups stating that invalid file directory for rows with fewwer attachments
'           Alternatively, you can just add the full directory to column F/H, but it takes time to copy and paste the full directory. This way is faster.
            Attachment1 = Cells(4, "C") & Cells(Cell.Row, "F") & Cells(Cell.Row, "G")
            Attachment2 = Cells(4, "C") & Cells(Cell.Row, "H") & Cells(Cell.Row, "I")
            If Not IsEmpty(Cells(Cell.Row, "F")) Then Outlook_mail.Attachments.Add (Attachment1)
            If Not IsEmpty(Cells(Cell.Row, "H")) Then Outlook_mail.Attachments.Add (Attachment2)
            
'           Fill the relevant fields of the Email Template
            With Outlook_mail
                .To = Cell.Value
                .CC = Cells(2, "B")
                .BCC = Cells(3, "B")
                .Subject = Cells(4, "B")
                .HTMLBody = Body & Outlook_mail.HTMLBody
'                .send 'uncomment this line to immediately send the email
            End With

        End If
    Next Cell

'   Re-enable excel events once we are done with the macro
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    
End Sub
