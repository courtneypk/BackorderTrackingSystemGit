VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Drug Shortage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ImportMBO_Click()

    DoCmd.RunSavedImportExport ("Import-MBO")

    Response = MsgBox("Import completed successfully")

End Sub

Private Sub SelectChargeFile_Click()
    Dim Message, Title, Default, MyValue

    newFile = ""
    
    newFile = getFileName() 'If file is selected, then the While statement is exited.
    If newFile = "" Then
        tryAgain = MsgBox("No file name was selected. If you want to try again click Yes, if not click No", vbYesNo)
        If tryAgain = vbYes Then
            newFile = getFileName()
        Else
            Exit Sub
        End If
    End If
    
    ' Use Split to put the filename into the last element of an array.
    filePath = Split(newFile, "\", -1, vbTextCompare)

    ' Get number of elements in the array
    arrayLength = UBound(filePath)
    
    ' Get just the filename from the last element in the array, put it onto the form for reference.
    newFilename = filePath(arrayLength)
    'Me.FileName = newFilename Don't know if we need this or not.
    
    ' Pull just the date encoded in the filename as MMDDYYYY
    tmpDate = Mid(newFilename, 11, 8)
    ' Put the encoded date into a regular date format MM/DD/YYYYY
    tmpDate = Left(tmpDate, 2) & "/" & Mid(tmpDate, 3, 2) & "/" & Right(tmpDate, 4)
    ' MsgBox ("String Date is " & tmpDate)
    
    ' Make sure that the text pulled out from the filename is actually a date; if not
    ' then ask the user for a date.
    If IsDate(tmpDate) Then
        newDate = CDate(tmpDate)
        ' MsgBox ("Date is " & newDate)
        ' Me.fileDate = newDate
    Else
        MsgBox (tmpDate & " is not a date! Please enter date in form.")
        Message = "Enter a date using MM/DD/YYYY format"    ' Set prompt.
        Title = "Charging Data Date"    ' Set title.
        Default = Date     ' Set default.
        ' Display message, title, and default value.
        MyValue = InputBox(Message, Title, Default)
        
        ' Me.fileDate.Enabled = True
        ' DoCmd.OpenForm "NewChargingFileName", acNormal, , , , acDialog, dateOfFile
        newDate = CDate(MyValue)
        ' MsgBox ("Date is " & newDate)
        ' Me.fileDate = newDate
    End If
    
        
    ' Here we first establish if there are any datasets in the database by checking the devIDs table.
    ' If there are no records (database is empty) then set the newID to 1. Otherwise set the newID equal
    ' to the maximum of DatasetID incremented by 1.
    If DCount("[DatasetID]", "prodDatasetIDs") = 0 Then
        newID = 1
    Else
        newID = DMax("[DatasetID]", "prodDatasetIDs") + 1
    End If
    
    ' Okay, now we enable the next button and disable this button
    'Me.PrepChargeFile.Enabled = True
    'Me.SelectChargeFile.Enabled = False
    'Me.PrepChargeFile.SetFocus
    
End Sub

Function getFileName() As String

   Dim fDialog As Object

   Dim varFile As Variant
   Dim importFile As String
 
   ' Clear listbox contents.
   ' Me.fileList.RowSource = ""
 
   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
   
   With fDialog
 
      ' Don't allow user to make multiple selections in dialog box
      .AllowMultiSelect = False
      
      
      .InitialFileName = "\\Dfa1\dfedata$\Test\As Modified\"
             
      ' Set the title of the dialog box.
      .Title = "Please select the new file"
 
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Access Databases", "*.XLSX"
      .Filters.Add "All Files", "*.*"
 
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      
    If .Show = True Then
 
         'Loop through each file selected (there is only one) and add it to our list box.
        For Each varFile In .SelectedItems
            ' Me.fileList.AddItem varFile
            importFile = varFile
        Next
        getFileName = importFile
    Else
         ' MsgBox "You clicked Cancel in the file dialog box."
         getFileName = ""
    End If
      
   End With
    

   Exit Function

End Function

