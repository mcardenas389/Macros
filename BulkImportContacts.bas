Attribute VB_Name = "Module1"
' imports contacts from emails from a folder
' searches for the folder with FindInFolders()
' calls ImportToContacts() when folder is found
Sub BulkImportContacts()
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
    Dim constantContact As String
    constantContact = "donotreply_eventspot@constantcontact.com"
    Dim paypal As String
    paypal = "service@paypal.com"
    Dim body As String
    Dim counter As Integer
    counter = 1
        
    For Each MyItem In FoundFolder.Items
        ' mark as read if it is unread
        If MyItem.UnRead Then
            MyItem.UnRead = False
        End If
        
        ' If MyItem.sender Is sender Then
            ' Call CreateOrUpdateContact(MyItem.body)
        ' ElseIf MyItem.sender Is paypal Then
        Call UpdatePayment(MyItem.body)
        ' End If
        
        ' for debugging
        MyItem.SaveAs "C:\Users\Hunter\Documents\out" & counter & ".txt", olTXT
        counter = counter + 1
    Next
    
    ' clean up
    Set ContactsFolder = Nothing
End Sub

Sub UpdatePayment(body As String)
    Dim messageArray() As String
    Dim splitArray() As String
    Dim delimitedMessage As String
    Dim lastName As String
    Dim email As String
    Dim Contact As Outlook.ContactItem
    
    delimitedMessage = Replace(body, "Buyer information", "###")
    delimitedMessage = Replace(delimitedMessage, "Instructions from buyer", "###")
    messageArray = Split(delimitedMessage, "###")
    splitArray = Split(messageArray(1), vbNewLine)
    lastName = splitArray(1)
    email = splitArray(2)
    ' email = Split(email, Chr(34))
    
    Debug.Print "Last: " & lastName & vbNewLine & "Email: " & email
    
    splitArray = Split(splitArray(0), Chr(32))
    lastName = splitArray(0)
    
    Set Contact = FindContact2(lastName, email)
    
    If Contact Is Nothing Then
        MsgBox ("no one")
    Else
        MsgBox ("someone")
    End If
    
    ' clean up
    Contact = Nothing
End Sub

' gets contact information from email body and uses this information to populate a contact card
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
        
    ' clean up values and remove unwanted characters
    Dim i As Integer
    For i = 1 To UBound(messageArray)
        ' remove the " mark from the hyperlink
        If i = 3 Or i = 4 Or i = 8 Then
            splitArray = Split(messageArray(i), Chr(34))
            messageArray(i) = splitArray(UBound(splitArray))
        End If
        
        ' remove the newline character and replace it with an empty string
        messageArray(i) = Replace(messageArray(i), vbNewLine, "")
    Next
        
    ' search for contact after collecting the relevant data
    Set Contact = FindContact(messageArray(1), messageArray(2), messageArray(3))
        
    ' if the contact is not found, then create a new contact without prompting the user
    ' if the contact is found, then prompt the user before updating it
    If Contact Is Nothing Then
        Set Contact = Application.CreateItem(olContactItem)
    Else
        ' build prompt
        ' new contact info
        Dim prompt As String
        prompt = "Contact exists!" & vbNewLine & vbNewLine & "New information:" & vbNewLine & _
            "Name: " & messageArray(1) & " " & messageArray(2) & vbNewLine & _
            "Email: " & messageArray(3) & vbNewLine & _
            "Phone: " & messageArray(4) & vbNewLine & _
            "Company: " & messageArray(6) & vbNewLine & _
            "Job Title: " & messageArray(7) & vbNewLine & _
            "Address: " & messageArray(8) & vbNewLine & messageArray(9) & ", " & _
            StateAbbreviation(messageArray(10)) & " " & messageArray(11) & vbNewLine & _
            messageArray(12) & vbNewLine
        
        ' old contact info
        prompt = prompt & vbNewLine & "Old information:" & vbNewLine & _
            "Name: " & Contact.FullName & vbNewLine & _
            "Email: " & Contact.Email1Address & vbNewLine & _
            "Phone: " & Contact.BusinessTelephoneNumber & vbNewLine & _
            "Company: " & Contact.CompanyName & vbNewLine & _
            "Job Title: " & Contact.JobTitle & vbNewLine & _
            "Address: " & Contact.BusinessAddress & vbNewLine & Contact.BusinessAddressCountry & vbNewLine & _
            vbNewLine & "Update with new information?"
            
        If MsgBox(prompt, vbQuestion Or vbYesNo) = vbNo Then
            Set Contact = Nothing
        End If
    End If
    
    ' create or update contact if contact object has been set
    If Not Contact Is Nothing Then
        Contact.firstName = messageArray(1)
        Contact.lastName = messageArray(2)
        Contact.Email1Address = messageArray(3)
        Contact.BusinessTelephoneNumber = messageArray(4)
        Contact.CompanyName = messageArray(6)
        Contact.JobTitle = messageArray(7)
        Contact.BusinessAddressStreet = messageArray(8)
        Contact.BusinessAddressCity = messageArray(9)
        Contact.BusinessAddressState = StateAbbreviation(messageArray(10))
        Contact.BusinessAddressPostalCode = messageArray(11)
        Contact.BusinessAddressCountry = messageArray(12)
        
        ' notes
        Contact.body = Year(Date) & " Regional Conference" & vbNewLine & _
            "Position: " & messageArray(13) & vbNewLine
        
        ' save contact data
        Contact.Save
    End If
    
    ' clean up
    Set Contact = Nothing
End Sub

' searches for a contact using a given first name, last name, and email
' returns contact object if it is found and Nothing if it is not found
Function FindContact(firstName As String, lastName As String, email As String)
    Dim filter As String
    filter = "[FullName] = " & Chr(34) & firstName & " " & lastName & Chr(34) & _
        " And [E-mail] = " & Chr(34) & email & Chr(34)
        
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

' searches for a contact using a given first name, last name, and email
' returns contact object if it is found and Nothing if it is not found
Function FindContact2(lastName As String, email As String)
    Dim filter As String
    filter = "[LastName] = " & Chr(34) & lastName & Chr(34) & _
        " And [E-mail] = " & Chr(34) & email & Chr(34)
        
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
' returns the state abbreviation
' returns the original string if it is not found
Function StateAbbreviation(stateName As String)
    Dim sn As String
    sn = UCase(stateName)
    
    StateAbbreviation = stateName
    
    If sn = "ALABAMA" Then
        StateAbbreviation = "AL"
    ElseIf sn = "ALASKA" Then
        StateAbbreviation = "AK"
    ElseIf sn = "ARIZONA" Then
        StateAbbreviation = "AZ"
    ElseIf sn = "ARKANSAS" Then
        StateAbbreviation = "AR"
    ElseIf sn = "CALIFORNIA" Then
        StateAbbreviation = "CA"
    ElseIf sn = "COLORADO" Then
        StateAbbreviation = "CO"
    ElseIf sn = "CONNECTICUT" Then
        StateAbbreviation = "CT"
    ElseIf sn = "DELAWARE" Then
        StateAbbreviation = "DE"
    ElseIf sn = "FLORIDA" Then
        StateAbbreviation = "FL"
    ElseIf sn = "GEORGIA" Then
        StateAbbreviation = "GA"
    ElseIf sn = "HAWAII" Then
        StateAbbreviation = "HI"
    ElseIf sn = "IDAHO" Then
        StateAbbreviation = "ID"
    ElseIf sn = "ILLINOIS" Then
        StateAbbreviation = "IL"
    ElseIf sn = "INDIANA" Then
        StateAbbreviation = "IN"
    ElseIf sn = "IOWA" Then
        StateAbbreviation = "IA"
    ElseIf sn = "KANSAS" Then
        StateAbbreviation = "KS"
    ElseIf sn = "KENTUCKY" Then
        StateAbbreviation = "KY"
    ElseIf sn = "LOUISIANA" Then
        StateAbbreviation = "LA"
    ElseIf sn = "MAINE" Then
        StateAbbreviation = "ME"
    ElseIf sn = "MARYLAND" Then
        StateAbbreviation = "MD"
    ElseIf sn = "MASSACHUSETTS" Then
        StateAbbreviation = "MA"
    ElseIf sn = "MICHIGAN" Then
        StateAbbreviation = "MI"
    ElseIf sn = "MINNESOTA" Then
        StateAbbreviation = "MN"
    ElseIf sn = "MISSISSIPPI" Then
        StateAbbreviation = "MS"
    ElseIf sn = "MISSOURI" Then
        StateAbbreviation = "MO"
    ElseIf sn = "MONTANA" Then
        StateAbbreviation = "MT"
    ElseIf sn = "NEBRASKA" Then
        StateAbbreviation = "NE"
    ElseIf sn = "NEVADA" Then
        StateAbbreviation = "NV"
    ElseIf sn = "NEW HAMPSHIRE" Then
        StateAbbreviation = "NH"
    ElseIf sn = "NEW JERSEY" Then
        StateAbbreviation = "NJ"
    ElseIf sn = "NEW MEXICO" Then
        StateAbbreviation = "NM"
    ElseIf sn = "NEW YORK" Then
        StateAbbreviation = "NY"
    ElseIf sn = "NORTH CAROLINA" Then
        StateAbbreviation = "NC"
    ElseIf sn = "NORTH DAKOTA" Then
        StateAbbreviation = "ND"
    ElseIf sn = "OHIO" Then
        StateAbbreviation = "OH"
    ElseIf sn = "OKLAHOMA" Then
        StateAbbreviation = "OK"
    ElseIf sn = "OREGON" Then
        StateAbbreviation = "OR"
    ElseIf sn = "PENNSYLVANIA" Then
        StateAbbreviation = "PA"
    ElseIf sn = "RHODE ISLAND" Then
        StateAbbreviation = "RI"
    ElseIf sn = "SOUTH CAROLINA" Then
        StateAbbreviation = "SC"
    ElseIf sn = "SOUTH DAKOTA" Then
        StateAbbreviation = "SD"
    ElseIf sn = "TENNESSEE" Then
        StateAbbreviation = "TN"
    ElseIf sn = "TEXAS" Then
        StateAbbreviation = "TX"
    ElseIf sn = "UTAH" Then
        StateAbbreviation = "UT"
    ElseIf sn = "VERMONT" Then
        StateAbbreviation = "VT"
    ElseIf sn = "VIRGINIA" Then
        StateAbbreviation = "VA"
    ElseIf sn = "WASHINGTON" Then
        StateAbbreviation = "WA"
    ElseIf sn = "WEST VIRGINIA" Then
        StateAbbreviation = "WV"
    ElseIf sn = "WISCONSIN" Then
        StateAbbreviation = "WI"
    ElseIf sn = "WYOMING" Then
        StateAbbreviation = "WY"
    End If
End Function
