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
Dim tempDoc As Document 'temp document to delete comments
Dim tempPath As String 'path to temp files locally
Dim tempName As String 'temp name to prevent duplicate errors

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
On Error GoTo 0
    
'**********************************************************************************
uniqueName = False 'Sets UniqueName to FALSE as the default, and the checks set it to True and execute PDF export
    'UniqueName = FALSE, the PDF already exists and the function has you rename or exit
    'UniqueName = TRUE, there is nothing to overwrite and so exports the PDF to the active directory
myPath = ActiveDocument.FullName 'Gets full name of current document
  
isCloud = checkCloud(myPath) 'Check if the file is saved to a cloud location
slashType = checkSlash(myPath) 'Store the correct type of slash for the path, link or local

'Checks for a backslash within the file path.
'If empty, the file isn't saved locally, and a prompt will open to save file
If InStr(myPath, "\") = 0 And isCloud = False Then 'Check if file is saved locally AND is not a cloud save
   UserAnswer = MsgBox("File Is Not Saved! Click " & _
     "[Yes] to Save As. Click [No] to Exit.", vbYesNoCancel)
      If UserAnswer = vbYes Then
        ShowSaveAsDialog
        myPath = ActiveDocument.FullName 'set the new doc path after save
      ElseIf UserAnswer = vbNo Then
        MsgBox "Save File and Try Again"
        Exit Sub
      End If
End If

currentFolder = ActiveDocument.path & slashType 'adds the right slash type to the end of the document path, used to create PDF filesave path
docName = ActiveDocument.Name
docName = Left(docName, (InStrRev(docName, ".") - 1)) 'gets the name of the file without the extension

'Set full filename to PDF extension to allow for check of existing file
fullFile = currentFolder & docName & ".pdf"
If isCloud = True Then uniqueName = Not CheckUrlExists(fullFile) 'Check if PDF file already exists in cloud link.  If link is valid, Unique set to FALSE
'**********************************************************************************
'Loop to rename the file if a PDF already exists.
'Two cases, one for cloud save, one for local save (isCloud is True or False)
On Error GoTo uniqueNameFail
Select Case isCloud
 Case True
    Do While uniqueName = False 'separate loop for the cloud save name check
       UserAnswer = MsgBox("Cloud File Already Exists! Click " & _
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
'Local file save rename loop
Case False
    Do While uniqueName = False
       UserAnswer = MsgBox("Local File Already Exists! Click " & _
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
                uniqueName = Not fileExists(fullFile)
                'Exit if User Wants To
                    If docName = "False" Or docName = "" Then Exit Sub
            Loop While ValidFileName(docName) = False
          Else
            Exit Sub 'Cancel
          End If
    Loop
End Select
On Error GoTo 0

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
'Comments do not export correctly and so need to be deleted before the PDF is created
'Creates a temp file copy of the active doc, deletes all comments, and exports to PDF using the original path and name
On Error GoTo tempSaveFail
tempName = docName & "-temp"
tempPath = (Environ("TEMP") & "\" & tempName & ".docx")
If fileExists(tempPath) <> False Then
Application.Documents(tempPath).Activate
ActiveDocument.Close SaveChanges:=WdSaveOptions.wdDoNotSaveChanges
Kill (tempPath)
Application.Documents(myPath).Activate
End If
'doc.Application.Activate
Set doc = Documents.Add(ActiveDocument.FullName)
ActiveDocument.SaveAs2 FileName:=tempPath, _
    FileFormat:=wdFormatDocumentDefault, AddToRecentFiles:=False
On Error GoTo 0
doc.ActiveWindow.Visible = False
On Error GoTo noComments
ActiveDocument.DeleteAllComments
On Error GoTo 0
'**********************************************************************************
'Automatic link updates sometimes show tracked changes when they refresh
'Runs the refUpdate function to refresh the cross-references, TOC, etc without tracked changes
refUpdate
'**********************************************************************************
'Save As PDF Document
On Error GoTo ProblemSaving
    ActiveDocument.ExportAsFixedFormat _
     OutputFileName:=fullFile, _
     OpenAfterExport:=False, _
     ExportFormat:=wdExportFormatPDF, _
     Item:=wdExportDocumentWithMarkup
On Error GoTo 0
'Closes the temporary document
ActiveDocument.Close SaveChanges:=WdSaveOptions.wdDoNotSaveChanges
Kill (tempPath) 'Delete Temp File
'Activates the original doc
Application.Documents(myPath).Activate
'**********************************************************************************
'Resets the tracked changes settings based on the backups

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
'**********************************************************************************
'Confirm Save To User
  If isCloud = False Then
  With ActiveDocument
    FolderName = Mid(.path, InStrRev(.path, "\") + 1, Len(.path) - InStrRev(.path, "\"))
  End With
  Else: FolderName = currentFolder 'sets just to URL
  FolderName = Replace(FolderName, "%", " ") 'replace % characters from URL with regular spaces for readability
  End If
  
  MsgBox "PDF Saved in the Folder: " & FolderName
Exit Sub
'**********************************************************************************
'Error Handlers

ExitSub:
    Exit Sub

colorError:
    MoveFromTextColor_backup = (-1) 'if set to ByAuthor throws an error and overflow issue
    MoveToTextColor_backup = (-1) 'if set to ByAuthor throws an error and overflow issue
Resume Next

uniqueNameFail:
MsgBox "Error with Updated Name, Check path and try again" & Err.Description
Resume ExitSub

tempSaveFail:
MsgBox "There was an issue saving the temporary file.: " & Err.Number & Err.Description
ActiveDocument.Close SaveChanges:=WdSaveOptions.wdDoNotSaveChanges
Kill (tempPath)
Resume ExitSub

noComments:
MsgBox "No Comments To Delete"
Resume Next

ProblemSaving:
  MsgBox "There was a problem saving your PDF. This is most commonly caused" & _
   " by the original PDF file already being open."
Resume ExitSub

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
Function fileExists(path As String)
    If Len(Dir(path)) <> 0 Then
        fileExists = True
    Else
        fileExists = False
    End If
End Function


Function checkCloud(xLink As String) As Boolean
'**********************************************************
'Check if the current path (xLink) is a cloud save location.
'Local folders use "\", links use "http"
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
    Selection.WholeStory 'Selects entire doc
    Selection.Fields.Update 'Replicates F9
    Application.ScreenUpdating = True
End Function
