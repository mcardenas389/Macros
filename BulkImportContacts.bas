Attribute VB_Name = "Module1"
Sub CompanyChange()
    Dim ContactsFolder As Folder
    Set ContactsFolder = Session.GetDefaultFolder(olFolderContacts)
    Dim OldCompanyName As String
    Dim NewCompanyName As String
    Dim OldEmailDomain As String
    Dim NewEmailDomain As String
    Dim ContactsChangedCount As Integer
    
    ' Ask user for inputs
    MsgBox ("This script will change all of your contacts from one company to another.")
    OldCompanyName = InputBox("Under what name are the contacts listed in Outlook now?")
    NewCompanyName = InputBox("What is the new company name to set them to?")
    OldEmailDomain = InputBox("What is the e-mail domain name currently listed after the @ sign? e.g. mycompany.com")
    NewEmailDomain = InputBox("What should the e-mail domain be set to? Leave blank and click OK if no change")
    ContactsChangedCount = 0
    
    Dim Contact As ContactItem
 
    ' loop through Contacts and set those who need it
    For Each Contact In ContactsFolder.Items
        If Contact.CompanyName = OldCompanyName Then
            Contact.CompanyName = NewCompanyName
            If NewEmailDomain <> "" Then
                Contact.Email1Address = Replace(Contact.Email1Address, OldEmailDomain, NewEmailDomain)
            End If
            Contact.Save
            ContactsChangedCount = ContactsChangedCount + 1
            Debug.Print "Changed: " & Contact.FullName
        End If
    Next
    ' confirm and clean up
    MsgBox (ContactsChangedCount & " contacts were changed from '" & OldCompanyName & "' to '" & NewCompanyName)
    Set Contact = Nothing
    Set ContactsFolder = Nothing
End Sub

' imports contacts from emails from a folder
' searches for the folder with FindInFolders()
' calls ImportToContacts() when folder is found
Private Sub BulkImportContacts()
    Dim name As String
    Dim FoundFolder As Folder

    ' if nothing is entered, exit out of macro
    name = InputBox("Enter folder name:", "Search Folder")
    If Len(Trim$(name)) = 0 Then Exit Sub
    
    ' find out if folder exists
    Set FoundFolder = FindInFolders(Application.Session.Folders, name)
    
    ' if folder is not found or is empty, do nothing
    ' if the folder is found and has items, call ImportToContacts()
    If FoundFolder Is Nothing Then
        MsgBox "Not Found", vbInformation
        Exit Sub
    ElseIf FoundFolder.Items.Count = 0 Then
        MsgBox ("Folder is empty.")
        Exit Sub
    'If MsgBox("Activate Folder: " & vbCrLf & FoundFolder.FolderPath, vbQuestion Or vbYesNo) = vbYes Then
      'Set Application.ActiveExplorer.CurrentFolder = FoundFolder
    'End If
    Else
        Call ImportToContacts(FoundFolder)
    End If
    
    MsgBox ("Process was successful!")
    
    ' clean up
    Set FoundFolder = Nothing
End Sub

' searches for a given folder name in the inbox
Private Function FindInFolders(TheFolders As Outlook.Folders, name As String)
    Dim SubFolder As Outlook.Folder
    
    On Error Resume Next
    
    Set FindInFolders = Nothing
    
    For Each SubFolder In TheFolders
        If LCase(SubFolder.name) Like LCase(name) Then
            ' return value
            Set FindInFolders = SubFolder
            Exit For
        Else
            ' return value
            Set FindInFolders = FindInFolders(SubFolder.Folders, name)
            If Not FindInFolders Is Nothing Then Exit For
        End If
    Next
End Function

' searches through a given folder and gets contact information from email body
' calls CreateOrUpdateContact() to create new contacts
Private Sub ImportToContacts(FoundFolder As Folder)
    Dim MyItem As Outlook.MailItem
    Dim ContactsFolder As Folder
    Set ContactsFolder = Session.GetDefaultFolder(olFolderContacts)
    Dim sender As String
    sender = "donotreply_eventspot@constantcontact.com"
    Dim body As String
    Dim counter As Integer
    counter = 1
        
    For Each MyItem In FoundFolder.Items
        ' mark as read if it is unread
        If MyItem.UnRead Then
            MyItem.UnRead = False
        End If
        
        ' If MyItem.sender Is sender Then
        Call CreateOrUpdateContact(MyItem.body)
        
        ' for debugging
        MyItem.SaveAs "C:\Users\Hunter\Documents\out" & counter & ".txt", olTXT
        counter = counter + 1
    Next
    
    ' clean up
    Set ContactsFolder = Nothing
End Sub

' gets contact information from email body
Private Sub CreateOrUpdateContact(body As String)
    Dim messageArray() As String
    Dim splitArray() As String
    Dim delimitedMessage As String
    Dim Contact As Outlook.ContactItem
    
    ' replace specific text with ### in order to split it up into an array
    ' field names may change if email body changes in the future
    delimitedMessage = Replace(body, "First Name:", "###")
    delimitedMessage = Replace(delimitedMessage, "Last Name:", "###")
    delimitedMessage = Replace(delimitedMessage, "Email Address:", "###")
    delimitedMessage = Replace(delimitedMessage, "Phone:", "###")
    delimitedMessage = Replace(delimitedMessage, "Business Information", "###")
    delimitedMessage = Replace(delimitedMessage, "Company:", "###")
    delimitedMessage = Replace(delimitedMessage, "Job Title:", "###")
    delimitedMessage = Replace(delimitedMessage, "Address", "###")
    delimitedMessage = Replace(delimitedMessage, "City:", "###")
    delimitedMessage = Replace(delimitedMessage, "State:", "###")
    delimitedMessage = Replace(delimitedMessage, "ZIP Code:", "###")
    delimitedMessage = Replace(delimitedMessage, "Country:", "###")
    delimitedMessage = Replace(delimitedMessage, "What is your position?", "###")
    delimitedMessage = Replace(delimitedMessage, "Payment Summary", "###")
    delimitedMessage = Replace(delimitedMessage, "Total", "###")
    messageArray = Split(delimitedMessage, "###")
        
    ' search for contact after collecting the relevant data
    ' clean up email address string before passing it to FindContact()
    splitArray = Split(messageArray(3), Chr(34))
    Set Contact = FindContact(messageArray(1), messageArray(2), splitArray(UBound(splitArray)))
    
    ' build prompt
    ' add phone, job title, company name, address
    Dim prompt As String
    prompt = "Contact exists!" & vbNewLine & "Name: " & Replace(messageArray(1), vbNewLine, "") & _
            " " & Replace(messageArray(2), vbNewLine, "") & vbNewLine & _
            "Email: " & Replace(splitArray(UBound(splitArray)), vbNewLine, "") & vbNewLine & _
            vbNewLine & "Update with new information?"
    
    ' if the contact is not found, then create a new contact without prompting the user
    ' if the contact is found, then prompt the user before updating it
    If Contact Is Nothing Then
        Set Contact = Application.CreateItem(olContactItem)
    Else
        If MsgBox(prompt, vbQuestion Or vbYesNo) = vbNo Then
            Set Contact = Nothing
        End If
    End If
    
    ' create or update contact if contact object has been set
    If Not Contact Is Nothing Then
        Contact.firstName = messageArray(1)
        Contact.lastName = messageArray(2)
        
        ' split array at " marks from hyperlink
        splitArray = Split(messageArray(3), Chr(34))
        ' remove the newline characters from email address with empty string
        Contact.Email1Address = Replace(splitArray(UBound(splitArray)), vbNewLine, "")
        
        splitArray = Split(messageArray(4), Chr(34))
        Contact.BusinessTelephoneNumber = splitArray(UBound(splitArray))
        
        Contact.CompanyName = Replace(messageArray(6), vbNewLine, "")
        Contact.JobTitle = Replace(messageArray(7), vbNewLine, "")
        
        splitArray = Split(messageArray(8), Chr(34))
        Contact.BusinessAddressStreet = splitArray(UBound(splitArray))
            
        Contact.BusinessAddressCity = messageArray(9)
        Contact.BusinessAddressState = messageArray(10) ' change to state abbreviations
        Contact.BusinessAddressPostalCode = messageArray(11)
        Contact.BusinessAddressCountry = messageArray(12)
        
        splitArray = Split(messageArray(13), vbNewLine)
        ' add to notes, [current year] Regional Conference
        Contact.body = "Position: " & splitArray(2) & vbNewLine & _
            "Total payment: " & Replace(messageArray(UBound(messageArray)), vbNewLine, "")
        
        ' save contact data
        Contact.Save
    End If
    
    ' clean up
    Set Contact = Nothing
End Sub

' searches for a contact
' returns contact object if it is found and Nothing if it is not found
Function FindContact(firstName As String, lastName As String, email As String)
    Dim filter As String
    filter = "[FullName] = " & Chr(34) & firstName & " " & lastName & Chr(34) & _
        " And [E-mail] = " & Chr(34) & email & Chr(34)
    
    ' clean up string and remove newline characters with empty strings
    filter = Replace(filter, vbNewLine, "")
    
    Dim ContactsFolder As Folder
    Set ContactsFolder = Session.GetDefaultFolder(olFolderContacts)
    Dim Contact As Outlook.ContactItem
    Set Contact = ContactsFolder.Items.Find(filter)
    
    ' return value
    Set FindContact = Contact
    
    ' clean up
    Set ConctacsFolder = Nothing
    Set Conact = Nothing
End Function
