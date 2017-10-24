VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4050
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    With ComboBox1
        .AddItem "Today", 0
        .AddItem "Yesterday", 1
        .AddItem "A week", 2
        .AddItem "Two weeks", 3
        .AddItem "30 days", 4
        .ListIndex = 0
    End With
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
    If Len(Me.ComboBox1.Value) = 0 Then
        MsgBox "You must choose a value first!", vbExclamation, "ERROR"
        Exit Sub
    End If
    
    ' check if right option was selected
    Dim i As Integer
    For i = 0 To ComboBox1.ListCount
        If ComboBox1.Value = ComboBox1.List(i) Then
            Exit For
        ElseIf i = ComboBox1.ListCount - 1 Then
            MsgBox "Unknown value " & Chr(34) & ComboBox1.Value & Chr(34) & "!", vbExclamation, "ERROR"
            Exit Sub
        End If
    Next
    
    Module1.choice = ComboBox1.Value
    
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Module1.choice = "Cancel"
    
    Unload Me
End Sub

Private Sub Label1_Click()

End Sub
