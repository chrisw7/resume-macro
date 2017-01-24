VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mainForm 
   Caption         =   "Resume Generator"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4695
   OleObjectBlob   =   "mainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton2_Click()
    'Check for Templates subfolder
    If Dir(mainForm.TextBox2.Value & "\Templates\", vbDirectory) = "" Then
        MsgBox ("Template folder not found - correct Path and try again")
        mainForm.ComboBox1.Clear
    Else
        mainForm.Hide
    End If
End Sub


Private Sub CommandButton3_Click()
    Dim fldr As FileDialog
    'Open file dialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = "C:\Users\"
        .Show
    End With
    
    'Update file path (TODO: catch Dialog cancel)
    filePath = fldr.SelectedItems(1)
    mainForm.TextBox2.Value = filePath
        
    'Fetch templates names (*.docx) & fill combobox
    fileName = Dir(filePath & "\Templates\" & "*.docx")
    Do While fileName <> ""
        mainForm.ComboBox1.AddItem Left(fileName, Len(fileName) - 5)
        fileName = Dir()
    Loop
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent script from running on close
    If CloseMode = 0 Then
        End
    End If
End Sub


