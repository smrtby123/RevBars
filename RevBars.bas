Attribute VB_Name = "NewMacros1"
'**************************************************************
'Variables for backing up the current Review settings
Dim CommentsColor_backup As Integer
Dim DeletedTextColor_backup As Integer
Dim DeletedTextMark_backup As Integer
Dim InsertedTextColor_backup As Integer
Dim InsertedTextMark_backup As Integer
Dim MoveFromTextColor_backup As Long
Dim MoveFromTextMark_backup As Integer
Dim MoveToTextColor_backup As Integer
Dim MoveToTextMark_backup As Integer
Dim RevisedLinesMark_backup As Integer
Dim RevisedPropertiesColor_backup As Integer
Dim RevisedPropertiesMark_backup As Integer
Dim RevisionBalloon_backup As Integer
'***************************************************************
Dim isCloud As Boolean 'Checks if the current folder is a cloud drive
Dim currentFolder As String 'Derives the current folder the document is in.  Also used to check if the file is saved locally
Dim docName As String ' Used to store the FileName without an extension
Dim myPath As String 'The full path of the current file
Dim uniqueName As Boolean 'Used to check if the filename PDF already exists in the active folder
Dim slashType As String 'Used to store the correct Slash for the path.  Links get "/", local folders get "\"
Dim fullFile As String 'Used to store the full file name plus extension

Sub Revbars()
Attribute Revbars.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Revbars"
'**********************************************************************************
'This macro sets the proper formatting for markups and exports a PDF file
'that shows only the rev bars on the right hand side and no other markups.
'**********************************************************************************
'Backup current settings for markup views
 On Error GoTo colorError 'Error check for colors that give an overflow error when set to By Author, see below

        CommentsColor_backup = Options.CommentsColor ' Default Value of 0
        DeletedTextColor_backup = Options.DeletedTextColor 'Default Value of -1
        DeletedTextMark_backup = Options.DeletedTextMark 'Default Value of 9
        InsertedTextColor_backup = Options.InsertedTextColor 'Default Value of -1
        InsertedTextMark_backup = Options.InsertedTextMark 'Default Value of 5
        MoveFromTextColor_backup = Options.MoveFromTextColor 'if set to ByAuthor throws an error and overflow issue
        MoveFromTextMark_backup = Options.MoveFromTextMark 'Default Value of 10
        MoveToTextColor_backup = Options.MoveToTextColor 'if set to ByAuthor throws an error and overflow issue
        MoveToTextMark_backup = Options.MoveToTextMark 'Default Value of 5
        RevisedLinesMark_backup = Options.RevisedLinesMark 'Default Value of 2
        RevisedPropertiesColor_backup = Options.RevisedPropertiesColor 'Default Value of -1
        RevisedPropertiesMark_backup = Options.RevisedPropertiesMark 'Default Value of 5
        RevisionBalloon_backup = Options.RevisionsBalloonPrintOrientation 'Default Value of 1


    
   ' ActiveWindow.View.MarkupMode = wdInLineRevisions
    'With ActiveWindow.View.RevisionsFilter
       ' .Markup = wdRevisionsMarkupSimple
  '      .View = wdRevisionsViewFinal
   ' End With
'**********************************************************************************
'Output the PDF as a Save As
  '  With Dialogs(wdDialogFileSaveAs)
 '       .Format = wdFormatPDF
 '       .Show
 '   End With
    
'**********************************************************************************
'Alternate PDF output
'Store Information About Word File
'On Error GoTo xSaveAs
uniqueName = False 'Sets UniqueName to FALSE as the default, and the checks set it to True and execute PDF export
    'UniqueName = FALSE, the PDF already exists and the function has you rename or exit
    'UniqueName = TRUE, there is nothing to overwrite and so exports the PDF to the active directory
myPath = ActiveDocument.FullName
  
isCloud = checkCloud(myPath) 'Check if the file is saved to a cloud location
slashType = checkSlash(myPath) 'Store the correct type of slash for the path, link or local

'Checks for a backslash within the file path.
' If empty, the file isn't saved locally, and a prompt will open to save file]
If InStr(myPath, "\") = 0 And isCloud = False Then 'Check if file is saved locally AND is not a cloud save
   UserAnswer = MsgBox("File Is Not Saved! Click " & _
     "[Yes] to Save As. Click [No] to Exit.", vbYesNoCancel)
      If UserAnswer = vbYes Then
        ShowSaveAsDialog
        myPath = ActiveDocument.FullName 'set the doc path after save
      ElseIf UserAnswer = vbNo Then
        MsgBox "Save File and Try Again"
        Exit Sub
      End If
End If

currentFolder = ActiveDocument.Path & slashType
docName = ActiveDocument.Name
docName = Left(docName, (InStrRev(docName, ".") - 1)) 'gets the name of the file without the extension

'Set full filename to PDF extension to allow for check of existing file
fullFile = currentFolder & docName & ".pdf"
If isCloud = True Then uniqueName = Not CheckUrlExists(fullFile) 'Check if PDF file already exists in cloud link.  If link is valid, Unique set to FALSE
'Does File Already Exist?
Select Case isCloud
 Case True
    Do While uniqueName = False 'separate loop for the cloud save name check
       UserAnswer = MsgBox("File Already Exists! Click " & _
       "[Yes] to override. Click [No] to Rename.", vbYesNoCancel)
            If UserAnswer = vbYes Then
                uniqueName = True
            ElseIf UserAnswer = vbNo Then
                Do
                    'Retrieve New File Name
                    docName = InputBox("Provide New File Name " & _
                    "(will ask again if you provide an invalid file name)", _
                    "Enter File Name", docName)
                     fullFile = currentFolder & docName & ".pdf"
                     uniqueName = Not CheckUrlExists(fullFile)
          'Exit if User Wants To
                If docName = "False" Or docName = "" Then Exit Sub
                Loop While ValidFileName(docName) = False
            Else
                Exit Sub 'Cancel
            End If
     
    Loop
Case False
  Do While uniqueName = False
    If Len(Dir(fullFile)) <> 0 Then
      UserAnswer = MsgBox("File Already Exists! Click " & _
       "[Yes] to override. Click [No] to Rename.", vbYesNoCancel)
      
          If UserAnswer = vbYes Then
            uniqueName = True
          ElseIf UserAnswer = vbNo Then
            Do
                'Retrieve New File Name
                docName = InputBox("Provide New File Name " & _
             "(will ask again if you provide an invalid file name)", _
             "Enter File Name", docName)
          
          'Exit if User Wants To
            If docName = "False" Or docName = "" Then Exit Sub
        Loop While ValidFileName(docName) = False
      Else
        Exit Sub 'Cancel
      End If
    Else
      uniqueName = True
    End If
  Loop
  Case Else
  End Select

    refUpdate

'**********************************************************************************
'This option sets the markup to show only inline, no comment balloons or formatting
 ActiveWindow.View.MarkupMode = wdInLineRevisions

'**********************************************************************************
'Set the options for markup views to hide everything but rev bars on the right hand side
    With Options
        .MoveToTextColor = wdMoveToTextColorNone
        .MoveToTextMark = wdMoveToTextMarkHidden
        .MoveFromTextColor = wdMoveFromTextColorNone
        .MoveFromTextMark = wdMoveFromTextMarkHidden
        .InsertedTextMark = wdInsertedTextMarkNone
        .InsertedTextColor = wdInsertedTextColorNone
        .DeletedTextMark = wdDeletedTextMarkHidden
        .DeletedTextColor = wdDeletedTextColorNone
        .RevisedPropertiesMark = wdRevisedPropertiesMarkNone
        .RevisedPropertiesColor = wdRevisedPropertiesColorNone
        .RevisedLinesMark = wdRevisedLinesMarkRightBorder
        .CommentsColor = wdCommentsColorNone
        .RevisionsBalloonPrintOrientation = wdBalloonPrintOrientationPreserve
    End With
    
'**********************************************************************************
'Automatic link updates sometimes show tracked changes when they refresh
'This loop goes through all range objects and accept tracked changes on fields
'The loop then accepts any tracked changes affecting fields so the marks do not show



'**********************************************************************************
'This turns off tracking formatting, otherwise any format changes will show up as a balloon and mess up the doc
   With ActiveDocument
     .TrackFormatting = False
  '   .TrackRevisions = True
   End With
  
'Save As PDF Document
  On Error GoTo ProblemSaving
    ActiveDocument.ExportAsFixedFormat _
     OutputFileName:=currentFolder & docName & ".pdf", _
     OpenAfterExport:=False, _
     ExportFormat:=wdExportFormatPDF, _
     Item:=wdExportDocumentWithMarkup
  On Error GoTo 0


 
 '   ActiveDocument.ExportAsFixedFormat OutputFileName:= _
  '      Replace(ActiveDocument.FullName, ".docx", ".pdf"), _
   '     ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
    '    wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:= _
     '   wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
     '   CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
     '   BitmapMissingFonts:=True, UseISO19005_1:=False
    With ActiveDocument
      .TrackFormatting = True
    End With
    ActiveDocument.TrackRevisions = True
    With Options
        .CommentsColor = CommentsColor_backup
        .DeletedTextColor = DeletedTextColor_backup
        .DeletedTextMark = DeletedTextMark_backup
        .InsertedTextColor = InsertedTextColor_backup
        .InsertedTextMark = InsertedTextMark_backup
        .MoveFromTextColor = MoveFromTextColor_backup
        .MoveFromTextMark = MoveFromTextMark_backup
        .MoveToTextColor = MoveToTextColor_backup
        .MoveToTextMark = MoveToTextMark_backup
        .RevisedLinesMark = RevisedLinesMark_backup
        .RevisedPropertiesColor = RevisedPropertiesColor_backup
        .RevisedPropertiesMark = RevisedPropertiesMark_backup
    End With
 ActiveWindow.View.MarkupMode = wdInLineRevisions
 ActiveWindow.View.ShowComments = True
 
'Confirm Save To User
  If isCloud = False Then
  With ActiveDocument
    FolderName = Mid(.Path, InStrRev(.Path, "\") + 1, Len(.Path) - InStrRev(.Path, "\"))
  End With
  Else: FolderName = currentFolder
  End If
  
  MsgBox "PDF Saved in the Folder: " & FolderName
Exit Sub
'**********************************************************************************
'Error Handlers
ProblemSaving:
  MsgBox "There was a problem saving your PDF. This is most commonly caused" & _
   " by the original PDF file already being open."
Exit Sub

colorError:
'MsgBox "There was a color backup error"
    MoveFromTextColor_backup = (-1) 'if set to ByAuthor throws an error and overflow issue
    MoveToTextColor_backup = (-1) 'if set to ByAuthor throws an error and overflow issue
Resume Next
'*************************************************************************************
End Sub
Private Sub ShowSaveAsDialog()
'Initiates Save As dialog when the program detects the file isn't saved locally.
  With Dialogs(wdDialogFileSaveAs)
        .format = wdFormatDocument
        .Show
    End With
End Sub
Function ValidFileName(ByVal FileName As String) As Boolean
ValidFileName = Not (FileName Like "*[\/:*?<>|[""]*" Or FileName Like "*]*")
End Function
Function ValidFileName2(FileName As String) As Boolean
'PURPOSE: Determine If A Given Word Document File Name Is Valid
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim TempPath As String
Dim doc As Document

'Determine Folder Where Temporary Files Are Stored
  TempPath = Environ("TEMP")

'Create a Temporary XLS file (XLS in case there are macros)
  On Error GoTo InvalidFileName
    Set doc = ActiveDocument.SaveAs2(ActiveDocument.TempPath & _
     "\" & FileName & ".docx", wdFormatDocument)
  On Error Resume Next

'Delete Temp File
  Kill doc.FullName

'File Name is Valid
  ValidFileName = True

Exit Function

'ERROR HANDLERS
InvalidFileName:
MsgBox "File Name is Invalid"
  ValidFileName = False

End Function
Function checkSlash(xLink As String) As String
If InStr(xLink, "/") <> 0 Then
checkSlash = "/"
ElseIf InStr(xLink, "\") <> 0 Then
checkSlash = "\"
End If
End Function
Function CheckUrlExists(url As String) As Boolean
'*********************************************************
'Check if the PDF exists already for OneDrive
'Duplicate of the local save check
'*********************************************************
    On Error GoTo CheckUrlExists_Error
    
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
 
    xmlhttp.Open "HEAD", url, False
    xmlhttp.send
    
    If xmlhttp.Status = 200 Then
        CheckUrlExists = True 'File exists
    Else
        CheckUrlExists = False 'File does not exist
    End If
    
    Exit Function
    
CheckUrlExists_Error:
MsgBox "Link Check Error"
    CheckUrlExists = False
    
End Function

Function checkCloud(xLink As String) As Boolean
'**********************************************************
'Check if the current path (xLink) is a cloud save location.
'Local folders use "\", links use "/"
'**********************************************************
If InStr(xLink, "http") = 0 Then
    checkCloud = False
Else
    checkCloud = True
End If
End Function

Function refUpdate()
'**********************************************************************************
'Selects entire document and updates all references while tracked changes are off.
'Simulates Ctrl-A + F9
'***********************************************************************************
    ActiveDocument.TrackRevisions = False '[Turn off Tracked Changes]
    Application.ScreenUpdating = False '[Makes it so you can't see the refresh]
    Selection.WholeStory 'Selects entire
    Selection.Fields.Update
    Application.ScreenUpdating = True
End Function


