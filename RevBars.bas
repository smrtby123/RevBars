Attribute VB_Name = "NewMacros1"

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
Dim CurrentFolder As String
Dim FileName As String
Dim myPath As String
Dim UniqueName As Boolean

Sub Revbars()
Attribute Revbars.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Revbars"
'This macro sets the proper formatting for markups and exports a PDF file
'that shows only the rev bars on the right hand side.

'**********************************************************************************
'backup current settings for markup views
        CommentsColor_backup = Options.CommentsColor
        DeletedTextColor_backup = Options.DeletedTextColor
        DeletedTextMark_backup = Options.DeletedTextMark
        InsertedTextColor_backup = Options.InsertedTextColor
        InsertedTextMark_backup = Options.InsertedTextMark
        MoveFromTextColor_backup = Options.MoveFromTextColor
        MoveFromTextMark_backup = Options.MoveFromTextMark
        MoveToTextColor_backup = Options.MoveToTextColor
        MoveToTextMark_backup = Options.MoveToTextMark
        RevisedLinesMark_backup = Options.RevisedLinesMark
        RevisedPropertiesColor_backup = Options.RevisedPropertiesColor
        RevisedPropertiesMark_backup = Options.RevisedPropertiesMark
        RevisionBalloon_backup = Options.RevisionsBalloonPrintOrientation

'**********************************************************************************
'This option sets the markup to show only inline, no comment balloons or formatting
 ActiveWindow.View.MarkupMode = wdInLineRevisions

'**********************************************************************************
'Set the options for markup views to hide everything but rev bars
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
' This loop goes through all range objects and accept tracked changes on fields
'The loop then accepts any tracked changes affecting fields so the marks do not show
ActiveDocument.TrackRevisions = False
Application.ScreenUpdating = False
Dim Story As Range, oFld As Field, oRev As Revision, Rng As Range
With ActiveDocument
  For Each Story In .StoryRanges
    For Each oRev In Story.Revisions
      For Each oFld In oRev.Range.Fields
        oFld.ShowCodes = True
        Set Rng = oFld.Code
        With Rng
          .MoveEndUntil cset:=Chr(21), Count:=wdForward
          .MoveEndUntil cset:=Chr(19), Count:=wdBackward
          .End = .End + 1
          .Start = .Start - 1
          oFld.ShowCodes = False
          .Revisions.AcceptAll
        End With
      Next
    Next
  Next
End With
Set Rng = Nothing
Application.ScreenUpdating = True
'**********************************************************************************
'This turns off tracking formatting, otherwise any format changes will show up as a balloon and mess up the doc
   With ActiveDocument
     .TrackFormatting = False
  '   .TrackRevisions = True
   End With
    
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
 UniqueName = False
  myPath = ActiveDocument.FullName
  CurrentFolder = ActiveDocument.Path & "\"
  FileName = Mid(myPath, InStrRev(myPath, "\") + 1, _
   InStrRev(myPath, ".") - InStrRev(myPath, "\") - 1)

'Does File Already Exist?
  Do While UniqueName = False
    DirFile = CurrentFolder & FileName & ".pdf"
    If Len(Dir(DirFile)) <> 0 Then
      UserAnswer = MsgBox("File Already Exists! Click " & _
       "[Yes] to override. Click [No] to Rename.", vbYesNoCancel)
      
      If UserAnswer = vbYes Then
        UniqueName = True
      ElseIf UserAnswer = vbNo Then
        Do
          'Retrieve New File Name
            FileName = InputBox("Provide New File Name " & _
             "(will ask again if you provide an invalid file name)", _
             "Enter File Name", FileName)
          
          'Exit if User Wants To
            If FileName = "False" Or FileName = "" Then Exit Sub
        Loop While ValidFileName(FileName) = False
      Else
        Exit Sub 'Cancel
      End If
    Else
      UniqueName = True
    End If
  Loop
  
'Save As PDF Document
  On Error GoTo ProblemSaving
    ActiveDocument.ExportAsFixedFormat _
     OutputFileName:=CurrentFolder & FileName & ".pdf", _
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
  With ActiveDocument
    FolderName = Mid(.Path, InStrRev(.Path, "\") + 1, Len(.Path) - InStrRev(.Path, "\"))
  End With
  
  MsgBox "PDF Saved in the Folder: " & FolderName
Exit Sub
 'Error Handlers
ProblemSaving:
  MsgBox "There was a problem saving your PDF. This is most commonly caused" & _
   " by the original PDF file already being open."
Exit Sub

End Sub

Function ValidFileName(FileName As String) As Boolean
'PURPOSE: Determine If A Given Word Document File Name Is Valid
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim TempPath As String
Dim doc As Document

'Determine Folder Where Temporary Files Are Stored
  TempPath = Environ("TEMP")

'Create a Temporary XLS file (XLS in case there are macros)
  On Error GoTo InvalidFileName
    Set doc = ActiveDocument.SaveAs2(ActiveDocument.TempPath & _
     "\" & FileName & ".doc", wdFormatDocument)
  On Error Resume Next

'Delete Temp File
  Kill doc.FullName

'File Name is Valid
  ValidFileName = True

Exit Function

'ERROR HANDLERS
InvalidFileName:
'File Name is Invalid
  ValidFileName = False

End Function




