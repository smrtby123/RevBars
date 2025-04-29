'Option Explicit
Sub RevBars()

'***************************************************************
Dim userAnswer As Integer 'message box response variable
Dim isCloud As Boolean 'Checks if the current folder is a cloud drive
Dim currentFolder As String 'Derives the current folder the document is in.  Also used to check if the file is saved locally
Dim docName As String ' Used to store the FileName without an extension
Dim myPath As String 'The full path of the current file
Dim uniqueName As Boolean 'Used to check if the filename PDF already exists in the active folder
Dim slashType As String 'Used to store the correct Slash for the path.  Links get "/", local folders get "\"
Dim fullFile As String 'Used to store the full file name plus extension
Dim tempDoc As Object 'Temp file for PDF export
Dim tempPath As String 'path to temp files locally
Dim tempName As String 'temp name to prevent duplicate errors
Dim exportDoc As Object 'Doc that needs exporting
'**********************************************************************************
'This macro sets the proper formatting for markups and exports a PDF file
'that shows only the rev bars on the right hand side and no other markups.
'**********************************************************************************
uniqueName = False 'Sets UniqueName to FALSE as the default, and the checks set it to True and execute PDF export
    'UniqueName = FALSE, the PDF already exists and the function has you rename or exit
    'UniqueName = TRUE, there is nothing to overwrite and so exports the PDF to the active directory
currentFolder = ActiveDocument.path
If currentFolder = vbNullString And isCloud = False Then 'Check if file is saved locally AND is not a cloud save
'Checks if the ActiveDoc path is null. If there is no path, the file isn't saved locally, and a prompt will open to save file
   userAnswer = MsgBox("File Is Not Saved! Click " & _
     "[Yes] to Save As. Click [No] to Exit.", vbYesNoCancel)
      If userAnswer = vbYes Then
        ShowSaveAsDialog
      ElseIf userAnswer = vbNo Then
        MsgBox "Save File and Try Again"
        Exit Sub
      End If
End If
myPath = ActiveDocument.FullName 'Gets full name of current document
currentFolder = ActiveDocument.path
isCloud = checkCloud(myPath) 'Check if the file is saved to a cloud location
If isCloud = False Then
Set exportDoc = GetObject(myPath)
End If
slashType = checkSlash(myPath) 'Store the correct type of slash for the path, link or local
docName = Left$(ActiveDocument.Name, (InStrRev(ActiveDocument.Name, ".") - 1)) 'gets the name of the file without the extension


tempPath = (Environ$("TEMP") & "\" & docName & ".docx") 'Saves file to the windows default temp folder, other locations give a write error.
If fileExists(tempPath) <> False Then 'Check if Temp File already exists
Set tempDoc = GetObject(tempPath) 'Makes the previous temp doc the active document
tempDoc.Close SaveChanges:=WdSaveOptions.wdDoNotSaveChanges 'Closes old file if still open, happens sometimes with errors
Kill (tempPath) 'Delete existing temp file
End If
On Error GoTo tempSaveFail
Set tempDoc = Documents.Add(myPath) 'Sets the current doc as the active temp doc
tempDoc.SaveAs2 FileName:=tempPath, _
    FileFormat:=wdFormatDocumentDefault, AddToRecentFiles:=False 'Saveas in the temp location
On Error GoTo 0
tempDoc.ActiveWindow.Visible = False 'Makes it so you can't see the temp file when it reopens

'*************************************************************************************************
'Set full filename to PDF extension to allow for check of existing file
On Error GoTo uploadFail
fullFile = currentFolder & slashType & docName & ".pdf"
If isCloud = True Then
    uniqueName = Not CheckUrlExists(fullFile) 'Check if PDF file already exists in cloud link.  If link is valid, Unique set to FALSE
Else
    uniqueName = Not fileExists(fullFile) 'Checks in the original folder for existing PDF if the file is not a cloud link
End If

'**********************************************************************************
'Loop to rename the file if a PDF already exists.
'Two cases, one for cloud save, one for local save (isCloud is True or False)
On Error GoTo uniqueNameFail
Select Case isCloud
 Case True
    Do While uniqueName = False 'separate loop for the cloud save name check
       userAnswer = MsgBox("Cloud PDF Already Exists! Click " & _
       "[Yes] to override. Click [No] to Rename.", vbYesNoCancel)
            If userAnswer = vbYes Then
                uniqueName = True
            ElseIf userAnswer = vbNo Then
                Do
                    'Retrieve New File Name
                    docName = InputBox("Provide New File Name " & _
                    "(will ask again if you provide an invalid file name)", _
                    "Enter File Name", docName)
                     fullFile = currentFolder & docName & ".pdf"
                    'Exit if User Wants To
                If docName = "False" Or docName = vbNullString Then Exit Sub
                Loop While ValidFileName(docName) = False
            Else
                Exit Sub 'Cancel
            End If
    Loop
'Local file save rename loop
Case False
    Do While uniqueName = False
       userAnswer = MsgBox("Local PDF Already Exists! Click " & _
       "[Yes] to override. Click [No] to Rename.", vbYesNoCancel)
      
          If userAnswer = vbYes Then
            uniqueName = True
          ElseIf userAnswer = vbNo Then
            Do
                'Retrieve New File Name
                docName = InputBox("Provide New File Name " & _
                    "(will ask again if you provide an invalid file name)", _
                    "Enter File Name", docName)
                fullFile = currentFolder & docName & ".pdf"
                uniqueName = Not fileExists(fullFile)
                'Exit if User Wants To
                    If docName = "False" Or docName = vbNullString Then Exit Sub
            Loop While ValidFileName(docName) = False
          Else
            Exit Sub 'Cancel
          End If
    Loop
End Select
On Error GoTo 0

'**********************************************************************************
'This option sets the markup to show only inline, no comment balloons or formatting
 tempDoc.ActiveWindow.View.MarkupMode = wdInLineRevisions

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

On Error GoTo noComments
tempDoc.DeleteAllComments
On Error GoTo 0
'**********************************************************************************
'Fixes formatting so that there are only single spaces after periods.
'**********************************************************************************
With tempDoc
With Selection.Find 
 .ClearFormatting 
 .Text = ".  " 
 .Replacement.ClearFormatting 
 .Replacement.Text = ". " 
 .Execute Replace:=wdReplaceAll, Forward:=True, _ 
 Wrap:=wdFindContinue 
End With
'**********************************************************************************
'Automatic link updates sometimes show tracked changes when they refresh
'Runs the refUpdate function to refresh the cross-references, TOC, etc without tracked changes
refUpdate tempDoc
'**********************************************************************************
'Save As PDF Document
On Error GoTo ProblemSaving
    tempDoc.ExportAsFixedFormat _
     OutputFileName:=fullFile, _
     OpenAfterExport:=False, _
     ExportFormat:=wdExportFormatPDF, _
     Item:=wdExportDocumentWithMarkup
On Error GoTo 0
'Closes the temporary document
tempDoc.Close SaveChanges:=WdSaveOptions.wdDoNotSaveChanges
Kill (tempPath) 'Delete Temp File
'Activates the original doc
'Application.Documents(myPath).Activate
'**********************************************************************************
'Confirm Save To User
  If isCloud = False Then
  With exportDoc
    FolderName = Mid$(.path, InStrRev(.path, "\") + 1, Len(.path) - InStrRev(.path, "\"))
  End With
  Else: FolderName = currentFolder 'sets just to URL
  FolderName = Replace(FolderName, "%", " ") 'replace % characters from URL with regular spaces for readability
  End If
  
  MsgBox "PDF Saved in the Folder: " & FolderName
'**********************************************************************************
'Reset formatting for tracked changes
      With Options
        .InsertedTextMark = wdInsertedTextMarkColorOnly
        .InsertedTextColor = wdByAuthor
        .DeletedTextMark = wdDeletedTextMarkStrikeThrough
        .DeletedTextColor = wdByAuthor
        .RevisedPropertiesMark = wdRevisedPropertiesMarkColorOnly
        .RevisedPropertiesColor = wdBlue
        .RevisedLinesMark = wdRevisedLinesMarkRightBorder
        .CommentsColor = wdAuto
        .RevisionsBalloonPrintOrientation = wdBalloonPrintOrientationPreserve
    End With
    ActiveWindow.View.RevisionsMode = wdInLineRevisions
    With Options
        .MoveFromTextMark = wdMoveFromTextMarkStrikeThrough
        .MoveFromTextColor = wdByAuthor
        .MoveToTextMark = wdMoveToTextMarkColorOnly
        .MoveToTextColor = wdBlue
        .InsertedCellColor = wdCellColorLightBlue
        .MergedCellColor = wdCellColorLightYellow
        .DeletedCellColor = wdCellColorPink
        .SplitCellColor = wdCellColorLightOrange
    End With
    With ActiveDocument
        .TrackMoves = True
        .TrackFormatting = True
    End With
   
  
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
tempDoc.Close SaveChanges:=WdSaveOptions.wdDoNotSaveChanges
Kill (tempPath)
Resume ExitSub

uploadFail:
MsgBox "There was an issue uploading the file: " & Err.Number & Err.Description
tempDoc.Close SaveChanges:=WdSaveOptions.wdDoNotSaveChanges
Kill (tempPath)
Resume ExitSub

noComments:
'MsgBox "No Comments To Delete" [Uncomment for Debug]
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
        .Format = wdFormatDocumentDefault
        .Show
    End With
End Sub
Private Function ValidFileName(ByVal FileName As String) As Boolean
ValidFileName = Not (FileName Like "*[\/:*?<>|[""]*" Or FileName Like "*]*")
End Function
Private Function checkSlash(ByVal xLink As String) As String
If InStr(xLink, "/") <> 0 Then
checkSlash = "/"
ElseIf InStr(xLink, "\") <> 0 Then
checkSlash = "\"
End If
End Function
Private Function CheckUrlExists(ByVal url As String) As Boolean
'*********************************************************
'Check if the PDF exists already for OneDrive
'Duplicate of the local save check
'*********************************************************
    On Error GoTo CheckUrlExists_Error
    
    Dim tempPage As Object
    Set tempPage = CreateObject("MSXML2.XMLHTTP")
 
    tempPage.Open "HEAD", url, False
    tempPage.send
    
    CheckUrlExists = tempPage.Status = 200 'Checks if response from file URL
   
    Exit Function
    
CheckUrlExists_Error:
MsgBox "Link Check Error"
CheckUrlExists = False
Exit Function
    
End Function
Private Function fileExists(ByVal path As String) As Boolean
    fileExists = Len(Dir(path))
 End Function
Private Function checkCloud(ByVal xLink As String) As Boolean
'**********************************************************
'Check if the current path (xLink) is a cloud save location.
'Local folders use "\", links use "http"
'**********************************************************
checkCloud = InStr(xLink, "http") = 1
End Function

Private Sub refUpdate(ByVal actDoc As Object)
'**********************************************************************************
'Selects entire document and updates all references while tracked changes are off.
'Simulates Ctrl-A + F9
'***********************************************************************************
    actDoc.TrackRevisions = False 'Turn off Tracked Changes
    Application.ScreenUpdating = False 'Makes it so you can't see the refreshing screen
    Selection.WholeStory 'Selects entire doc
    Selection.Fields.Update 'Replicates F9 to refresh links and references
    Application.ScreenUpdating = True
End Sub
'***************************************************************

Sub ResetSettings()
''***************************************************************
' Macro to reset the tracked changes settings back to preferred default
''***************************************************************
'
    With Options
        .InsertedTextMark = wdInsertedTextMarkColorOnly
        .InsertedTextColor = wdByAuthor
        .DeletedTextMark = wdDeletedTextMarkStrikeThrough
        .DeletedTextColor = wdByAuthor
        .RevisedPropertiesMark = wdRevisedPropertiesMarkColorOnly
        .RevisedPropertiesColor = wdBlue
        .RevisedLinesMark = wdRevisedLinesMarkRightBorder
        .CommentsColor = wdAuto
        .RevisionsBalloonPrintOrientation = wdBalloonPrintOrientationPreserve
    End With
    ActiveWindow.View.RevisionsMode = wdInLineRevisions
    With Options
        .MoveFromTextMark = wdMoveFromTextMarkStrikeThrough
        .MoveFromTextColor = wdByAuthor
        .MoveToTextMark = wdMoveToTextMarkColorOnly
        .MoveToTextColor = wdBlue
        .InsertedCellColor = wdCellColorLightBlue
        .MergedCellColor = wdCellColorLightYellow
        .DeletedCellColor = wdCellColorPink
        .SplitCellColor = wdCellColorLightOrange
    End With
    With ActiveDocument
        .TrackMoves = True
        .TrackFormatting = True
    End With
End Sub
