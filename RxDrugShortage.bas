Attribute VB_Name = "RxDrugShortage"
Option Compare Database

Sub SelectMBOFile()

    Dim filePath As String
    
    
    filePath = "T:\Documents\Projects\Pharmacy\Drug Shortage\Downloaded MBO\MBO_09_15_2016 As Delivered.csv"


End Sub



Sub FileProperties()

    Dim filePath As String
    Dim dateCreated, dateModified As Variant
    
    filePath = "T:\Documents\Projects\Pharmacy\Backorder Tracking System\Downloaded Transaction Log\Transaction Log Report2162017.xlsx"
    ' dateCreated = ShowDateCreated(filePath)
    dateModified = ShowDateModified(filePath)

End Sub


Function ShowDateCreated(filespec)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.getFile(filespec)
    ShowDateCreated = f.dateCreated
End Function

Function ShowDateModified(filespec)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.getFile(filespec)
    ShowDateModified = f.DateLastModified
End Function

Function getFileName() As String
Attribute getFileName.VB_Description = "This function presents the user a file dialog box to select a file and it returns the full path of the file."

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
      
      ' Make the path directly to the folder where the dated folders are located
      .InitialFileName = "\\cifs2\dfpharm$\Materials Management\Backorders\"
             
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
