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

Sub BulkImportContacts()
    Dim Name As String
    Dim FoundFolder As Folder

    ' if nothing is entered, exit out of macro
    Name = InputBox("Enter folder name:", "Search Folder")
    If Len(Trim$(Name)) = 0 Then Exit Sub
    
    ' find out if folder exists
    Set FoundFolder = FindInFolders(Application.Session.Folders, Name)
    
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
End Sub

Function FindInFolders(TheFolders As Outlook.Folders, Name As String)
    Dim SubFolder As Outlook.Folder
    
    On Error Resume Next
    
    Set FindInFolders = Nothing
    
    For Each SubFolder In TheFolders
        If LCase(SubFolder.Name) Like LCase(Name) Then
            Set FindInFolders = SubFolder
            Exit For
        Else
            Set FindInFolders = FindInFolders(SubFolder.Folders, Name)
            If Not FindInFolders Is Nothing Then Exit For
        End If
    Next
End Function

Sub ImportToContacts(FoundFolder As Folder)
    Dim MyItem As Outlook.MailItem
    Dim ContactsFolder As Folder
    Set ContactsFolder = Session.GetDefaultFolder(olFolderContacts)
    Dim sender As String
    sender = "donotreply_eventspot@constantcontact.com"
    Dim body As String
    Dim counter As Integer
    counter = 1
        
    For Each MyItem In FoundFolder.Items
        ' If MyItem.sender Is sender Then
        Call CreateContact(MyItem.body)
        
        ' mark as read if it is unread
        If MyItem.UnRead Then
            MyItem.UnRead = False
        End If
        
        ' for debugging
        MyItem.SaveAs "C:\Users\Hunter\Documents\out" & counter & ".txt", olTXT
        counter = counter + 1
    Next
End Sub

Sub CreateContact(body As String)
    Dim messageArray() As String
    Dim splitArray() As String
    Dim delimitedMessage As String
    Dim Contact As Outlook.ContactItem
    Set Contact = Application.CreateItem(olContactItem)
    
    ' replace specific text with ### in order to split it up into an array
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
    
    Contact.FirstName = messageArray(1)
    Contact.LastName = messageArray(2)
    
    splitArray = Split(messageArray(3), Chr(34))
    Contact.Email1Address = splitArray(UBound(splitArray))
    
    splitArray = Split(messageArray(4), Chr(34))
    Contact.BusinessTelephoneNumber = splitArray(UBound(splitArray))
    
    Contact.CompanyName = messageArray(6)
    Contact.JobTitle = messageArray(7)
    
    splitArray = Split(messageArray(8), Chr(34))
    Contact.BusinessAddressStreet = splitArray(UBound(splitArray))
    
    Contact.BusinessAddressCity = messageArray(9)
    Contact.BusinessAddressState = messageArray(10)
    Contact.BusinessAddressPostalCode = messageArray(11)
    Contact.BusinessAddressCountry = messageArray(12)
    
    splitArray = Split(messageArray(13), vbNewLine)
    Contact.body = "Position: " & splitArray(2) & vbNewLine & _
        "Total payment: " & messageArray(UBound(messageArray))
    
    Contact.Save
End Sub
