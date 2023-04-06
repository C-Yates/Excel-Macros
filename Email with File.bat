Sub SendEmailWithAttachments()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim Filepath As String
    Dim Filename As Variant
    Dim RecipientEmail As Variant
    Dim RecipientName As String
    
    Set OutApp = CreateObject("Outlook.Application")
    
    'Change this list to however many emails you want
    For Each RecipientEmail In Array("person.1@gmail.com", "person.2@gmail.com", "person.3@gmail.com", "person.4@gmail.com")
    
        'Extract the recipient's first name from their email address
        RecipientName = Split(RecipientEmail, ".")(0)
    
        Set OutMail = OutApp.CreateItem(0)
        
        'Setup the email, only modify subject and body
        OutMail.To = RecipientEmail

        OutMail.Subject = "Subject Line of Your Email"
        
        'Important: To add new lines to the email you need "& vbNewLine & _" at the end of a line
        '           Then the next line you're adding has to start with "vbNewLine &" before you add your
        '           Own new line to the email and put your email text in quote marks like I hvae done
        OutMail.Body = "Dear " & RecipientName & "," & vbNewLine & _
        vbNewLine & "Please find attached the file you requested." & vbNewLine & _
        vbNewLine & "Best regards," & vbNewLine & "yourname"
        
        'Change the filenames here to whatever it's called, each person in the email list
        'Needs a filename associated as I have done below
        'If you want two people to have the same file sent to them do this line below (just adding a comma before the next person
        'Case "person.1@gmail.com", "person.2@gmail.com"
        'Note: the names of anyone you're sending a file to MUST be in the array above
        Select Case RecipientEmail
            Case "person.1@gmail.com"
                Filename = "Employee Sample Data.xlsx"
            Case "person.2@gmail.com"
                Filename = "Financials Sample Data.xlsx"
            Case "person.3@gmail.com"
                Filename = "Some Data.xlsx"
            Case "person.4@gmail.com"
                Filename = "More Data.xlsx"
        End Select
        
        'Set the directory for the files then add a slash at the end '\'
        Filepath = "C:\Users\ChrisYates\Documents\Batch & Excel\" & Filename
        
        'Attach the file to the email
        If Dir(Filepath) <> "" Then 'check if file exists first
            OutMail.Attachments.Add Filepath
        End If
        
        OutMail.Send
        Set OutMail = Nothing
        
    Next RecipientEmail
    Set OutApp = Nothing
    MsgBox "Completed"
    
End Sub
