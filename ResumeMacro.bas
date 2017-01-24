Attribute VB_Name = "ResumeMacro"
'Chris Williams | Jan 2017

'Simple script for quickly publishing resumes and/or cover letters as PDFs using a .docx resume/cover template document _
 containing a 1 pg cover letter and resume stored in the \Templates subfolder of the chosen root directory. Note _
 References to the company being applied to must use the "Company" Document Property found under (Insert > Quick Parts > Document Property)

Sub autoPublish()

    'Declare file/page length variables
    Dim filePath, fileName, pkgStr As String
    Dim startPg, endPg As Integer
 
    
    'Reset userform combo box
    mainForm.ComboBox1.Clear
    
    'Define path variables
    'filePath = "C:\Users\Christopher\Dropbox (Personal)\University of Waterloo\Resumes\"
    
    If mainForm.TextBox2.Value <> "" Then
        
        'TODO add Template folder checks
        fileName = Dir(filePath & "Templates\" & "*.docx")
        
        'Fetch templates names (*.docx) & fill combobox
        Do While fileName <> ""
            mainForm.ComboBox1.AddItem Left(fileName, Len(fileName) - 5)
            fileName = Dir()
        Loop
    
    End If
    
    'Open form for user input
    mainForm.Show
    
    
    filePath = mainForm.TextBox2.Value
    fileName = filePath & "\Templates\" & mainForm.ComboBox1.Value & ".docx"
    
    'Check if file to publish is open and in focus
    If ActiveDocument.Name = (mainForm.ComboBox1.Value & ".docx") Then
    Else
        Documents.Open(fileName).Activate
    End If
    
    ActiveDocument.BuiltInDocumentProperties(wdPropertyCompany) = mainForm.TextBox1.Value
    newName = mainForm.TextBox1.Value
    ActiveDocument.Save
    
    'Determine which pages to publish (assuming single page cover)
    If mainForm.OptionButton1.Value Then
        startPg = 1
        endPg = Selection.Information(wdNumberOfPagesInDocument)
        pkgStr = "Package"
    ElseIf mainForm.OptionButton2.Value Then
        startPg = 2
        endPg = Selection.Information(wdNumberOfPagesInDocument)
        pkgStr = "Resume"
    ElseIf mainForm.OptionButton3.Value Then
        startPg = 1
        endPg = 1
        pkgStr = "Cover"
    End If
    
    'Publish PDF
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            filePath & "\[" & "2017" & "] " & newName & "_" & pkgStr & ".pdf", _
            ExportFormat:=wdExportFormatPDF, _
            OpenAfterExport:=True, _
            OptimizeFor:=wdExportOptimizeForPrint, _
            Range:=wdExportFromTo, _
            From:=startPg, _
            To:=endPg, _
            IncludeDocProps:=True, _
            KeepIRM:=False, _
            CreateBookmarks:=wdExportCreateHeadingBookmarks, _
            DocStructureTags:=True, _
            BitmapMissingFonts:=True, _
            UseISO19005_1:=False
            
    'Close file if checked
    If mainForm.CheckBox1.Value Then
        ActiveDocument.Close (wdDoNotSaveChanges)
    End If
End Sub
