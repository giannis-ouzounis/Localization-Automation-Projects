Attribute VB_Name = "BT_with_CAT_Prep"
Option Explicit

Sub LV01a_BT_with_CAT_Tool_prep_batch()

'Pick source & save-to folders ➜ loop through Word docs ➜ run per-file macro ➜ save copy.
'Accept all changes, hide headers, insert blank column for back-translation, hide source rows.
'Unhide everything, strip headers, delete extra columns, add header row (“TRANSLATION / BACKTRANSLATION”), hide source rows.
'uhideAllText shows all hidden text. HideOddRowsFromXTMBilDoc hides every other row (source).


'Set Variables
    Dim strPath As String
    Dim strSavePath As String
    Dim strFile As String
    Dim docA As Document
    Dim docName As String
    'Timer variables
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    Dim MinutesElapsed As String
   
'On Error GoTo ErrHandler
    
    'Remember time when macro starts
    StartTime = Timer
   
'Ask for folder location with files to be processed.
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "SELECT FOLDER WITH XTM BILINGUAL WORD DOCS."
        If .Show = False Then
            MsgBox "You didn't select a folder.", vbInformation
            Exit Sub
        End If
        strPath = .SelectedItems(1)
    End With
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    
'Ask for folder location to save files to after process.
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "SELECT LOCATION TO SAVE EDITED FILES."
        If .Show = False Then
            MsgBox "You didn't select a folder.", vbInformation
            Exit Sub
        End If
        strSavePath = .SelectedItems(1)
    End With
    If Right(strSavePath, 1) <> "\" Then
        strSavePath = strSavePath & "\"
    End If
    'Test to see if 2nd location is the same as first.
    'Has to be here so that both strings have been through the same process
    If strSavePath = strPath Then
        MsgBox "Please select a different location the source files.", vbExclamation
        Exit Sub
    End If
    
    'No screen updates
    'Application.ScreenUpdating = False
    
    'Allows all doc, docx & docm files to be opened
    strFile = Dir(strPath & "*.doc?")
    
    'Loop
    Do While strFile <> ""
        Set docA = Documents.Open(strPath & strFile)
        docName = ActiveDocument.Name
        
'*****Call the macro you want to batch run here!*****
        'Add the macro as private sub below.
        Call LV01b_BT_with_CAT_Tool_prep(docA)
                
        'Save document with original name to new loction.
        docA.SaveAs2 FileName:=strSavePath & docName
        docA.Close
        'Next line is required to loop to the next document.
        strFile = Dir
    Loop
        
    'Screen updates
    Application.ScreenUpdating = True
    
        'Determine how many seconds code took to run
         MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
        
        'Notification that the loop has completed processing all files.
        MsgBox "All files have been processed." & vbCr & "Macro run time: " & MinutesElapsed, vbInformation

    Exit Sub
      
ErrHandler:
'EH code to provide more info about errors.
    Select Case Err.Number
        Case 5174
            MsgBox prompt:="Sorry, the required file cannot be found. Please check that all file names match." & _
                            vbCr & vbCr & docName & vbCr & "(" & Err.Number & " - " & Err.Description & ")", _
                            Buttons:=vbCritical
        Case 5914
            MsgBox prompt:="Input file has the wrong number of columns. Please use XTM BLT export file. This macro will now end. " & _
                            "(5914 - Missing the correct number of columns.)", _
                            Buttons:=vbCritical
        Case Else
            MsgBox prompt:="Sorry, an error has occurred. Please check that all file names match. Please screen shot this message and notify LE." & _
                            vbCr & vbCr & docName & vbCr & "(" & Err.Number & " - " & Err.Description & ")", _
                            Buttons:=vbCritical
    End Select
    
End Sub

Private Sub LV01b_BT_with_CAT_Tool_prep(docA As Document)

    Dim newCol As Column
    Dim newRow As row
    Dim i As Integer
    
    'No screen updates
    Application.ScreenUpdating = False
    
    'Activate docA to work on it
    docA.Activate

    'turn off track changes
    docA.TrackRevisions = False
    'Accept all changes in docA
    docA.Revisions.AcceptAll
    
    'Hide text in document headers
    docA.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader

    For i = 1 To docA.Sections.count
        Selection.WholeStory
        Selection.Font.Hidden = True

    If i = docA.Sections.count Then GoTo Line1

    docA.ActiveWindow.ActivePane.View.NextHeaderFooter

Line1:
Next
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
    'add space for the BT to go
    Set newCol = docA.Tables(1).Columns.Add(BeforeColumn:=docA.Tables(1).Columns(4))
    
    'Amend the Target column location after creating new empty columns
    docA.Tables(1).Columns(5).Select
    Selection.Copy
    docA.Tables(1).Columns(4).Select
HammerHeader:
    On Error GoTo FailHeader
        Selection.PasteAndFormat (wdFormatOriginalFormatting)
        GoTo HeaderPasteOK
FailHeader:
    Err.Clear
    On Error GoTo -1
    On Error GoTo 0
    DoEvents
    GoTo HammerHeader
HeaderPasteOK:
    On Error GoTo 0
    
    'remove extra columns
    docA.Tables(1).Columns(5).Delete
    
    'Autofit table to Window
    docA.Tables(1).Select
    docA.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
    'Make sure view is Print, as weblayout or read causes an error with the next step
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    
    'Select and hide all text
        Selection.WholeStory
        Selection.Font.Hidden = True
        
        'Unhide the column that contains the Source for the BT
        docA.Tables(1).Columns(5).Select
        Selection.Font.Hidden = False
    
    Call HideOddRowsFromXTMBilDoc(docA)

End Sub

Sub LV02a_BT_with_CAT_Tool_prepForEQS_batch()
'@Welocalize LS 2020
'Edited by Adrià Aleu & Stephanie Pietz

'First is the multi file saving macor to presever the original files.
'This Macro will take an XTM word Bilingual file removed the source text and
'move the target into the Left column, adding an empty right column for
'new translation.


'Set Variables
    Dim strPath As String
    Dim strSavePath As String
    Dim strFile As String
    Dim docA As Document
    Dim docName As String
    'Timer variables
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    Dim MinutesElapsed As String
   
'On Error GoTo ErrHandler
    
    'Remember time when macro starts
    StartTime = Timer
   
'Ask for folder location with files to be processed.
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "SELECT FOLDER WITH XTM BIL WORD DOCS."
        If .Show = False Then
            MsgBox "You didn't select a folder.", vbInformation
            Exit Sub
        End If
        strPath = .SelectedItems(1)
    End With
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    
'Ask for folder location to save files to after process.
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "SELECT LOCATION TO SAVE EDITED FILES."
        If .Show = False Then
            MsgBox "You didn't select a folder.", vbInformation
            Exit Sub
        End If
        strSavePath = .SelectedItems(1)
    End With
    If Right(strSavePath, 1) <> "\" Then
        strSavePath = strSavePath & "\"
    End If
    'Test to see if 2nd location is the same as first.
    'Has to be here so that both strings have been through the same process
    If strSavePath = strPath Then
        MsgBox "Please select a different location the source files.", vbExclamation
        Exit Sub
    End If
    
    'No screen updates
    'Application.ScreenUpdating = False
    
    'Allows all doc, docx & docm files to be opened
    strFile = Dir(strPath & "*.doc?")
    
    'Loop
    Do While strFile <> ""
        Set docA = Documents.Open(strPath & strFile)
        docName = ActiveDocument.Name
        
'*****Call the macro you want to batch run here!*****
        'Add the macro as private sub below.
        Call LV02b_BT_with_CAT_Tool_prepForEQS(docA)
                
        'Save document with original name to new loction.
        docA.SaveAs2 FileName:=strSavePath & docName
        docA.Close
        'Next line is required to loop to the next document.
        strFile = Dir
    Loop
        
    'Screen updates
    Application.ScreenUpdating = True
    
        'Determine how many seconds code took to run
         MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
        
        'Notification that the loop has completed processing all files.
        MsgBox "All files have been processed." & vbCr & "Macro run time: " & MinutesElapsed, vbInformation

    Exit Sub
      
ErrHandler:
'EH code to provide more info about errors.
    Select Case Err.Number
        Case 5174
            MsgBox prompt:="Sorry, the required file cannot be found. Please check that all file names match." & _
                            vbCr & vbCr & docName & vbCr & "(" & Err.Number & " - " & Err.Description & ")", _
                            Buttons:=vbCritical
        Case 5914
            MsgBox prompt:="Input file has the wrong number of columns. Please use XTM BLT export file. This macro will now end. " & _
                            "(5914 - Missing the correct number of columns.)", _
                            Buttons:=vbCritical
        Case Else
            MsgBox prompt:="Sorry, an error has occurred. Please check that all file names match. Please screen shot this message and notify LE." & _
                            vbCr & vbCr & docName & vbCr & "(" & Err.Number & " - " & Err.Description & ")", _
                            Buttons:=vbCritical
    End Select
    
End Sub

Private Sub LV02b_BT_with_CAT_Tool_prepForEQS(docA As Document)

    Dim newCol As Column
    Dim newRow As row
    Dim i As Integer
    
    'No screen updates
    Application.ScreenUpdating = False
    
    'Activate docA to work on it
    docA.Activate
    
    'Unhide all text text in document headers
    
    Call uhideAllText
    Call HideOddRowsFromXTMBilDoc(docA)

    'remove header and footers
    docA.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader

    For i = 1 To docA.Sections.count
        Selection.WholeStory
        Selection.Delete

    If i = docA.Sections.count Then GoTo Line1

    docA.ActiveWindow.ActivePane.View.NextHeaderFooter

Line1:
Next

    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

    'remove unneeded columns
    docA.Tables(1).Columns(6).Delete
    docA.Tables(1).Columns(3).Delete
    
    'Autofit table
    docA.Tables(1).Select
    docA.Tables(1).AutoFitBehavior (wdAutoFitWindow)

    'Make sure view is Print, as weblayout or read causes an error with the next step
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If

    'Add new empty row at the top of the table to add column names
    Set newRow = docA.Tables(1).Rows.Add(BeforeRow:=docA.Tables(1).Rows(1))

    'Adding column names to each column in the newly created first row
    With docA.Tables(1).Rows(1)
        .Cells(3).Range.Text = "TRANSLATION"
        .Cells(4).Range.Text = "BACKTRANSLATION"
    End With

End Sub

Private Sub uhideAllText()

Dim shp As Variant

'Unhide all text within Word document.
'*******************************************'
    'Uhnide all body text.
    
'    Range = WholeStory
'    With Range
'        .Font.Hidden = False
'    End With
    Selection.WholeStory
    With Selection.Font
        .Hidden = False
    End With
'*******************************************'
    'Uhide all text in text boxes.
    For Each shp In ActiveDocument.Shapes
        With shp.TextFrame
            If .HasText Then
                .TextRange.Font.Hidden = False
            End If
        End With
    Next shp
'*******************************************'
'Uhhid all header and footer text
'Source: http://www.vbaexpress.com/forum/archive/index.php/t-59028.html
'Set objRange = HeadersFooters

Dim Sctn As Section, HdFt As HeaderFooter
    
With ActiveDocument
    For Each Sctn In .Sections
    'Find each header
        For Each HdFt In Sctn.Headers
            With HdFt
                If .LinkToPrevious = False Then
                    With .Range
                    'Do range processing here
                        .Font.Hidden = False
                    End With
                End If
            End With
        Next
        
    'Find each footer
    For Each HdFt In Sctn.Footers
        With HdFt
            If .LinkToPrevious = False Then
                With .Range
                'Do range processing here
                    .Font.Hidden = False
                End With
            End If
        End With
        Next
    Next
End With
'*******************************************'
    'uhnide footnotes & end notes

End Sub

Sub HideOddRowsFromXTMBilDoc(doc As Document)

Dim i As Integer
Dim z As Integer
Dim intNumOfRows As Integer
Dim tbl As Table
Dim RowTable As Object

    'No screen updates
    Application.ScreenUpdating = False


Set tbl = doc.Tables(1)

   Set RowTable = tbl.Rows
   intNumOfRows = tbl.Rows.count
    
    'loop from the last row in the table to the 1st row, stepping 2 each time
    For i = intNumOfRows To 1 Step -2
        tbl.Rows(i).Select
        Selection.Font.Hidden = True
    Next
End Sub
