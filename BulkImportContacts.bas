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
    
    ' find if folder exists
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
    Dim SubFolder As Outlook.MAPIFolder
    
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
    Dim num As Integer
    num = 1
        
    For Each MyItem In FoundFolder.Items
        'Debug.Print "Sender: " & MyItem.Sender
        'Debug.Print "Body: " & MyItem.Body
        MyItem.SaveAs "C:\Users\Hunter\Documents\out" & num & ".txt", olTXT
        num = num + 1
    Next
End Sub
