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

    Dim Response1, Response2, Response3, tryAgain As Integer
    Dim MBOFileNamePath As String
    
    ' DoCmd.RunSavedImportExport ("Import-MBO")
    ' Need to use code for this process at this time
    ' Also, may need to ask in dialog box if user has already run the Excel code on the original CSV file to
    ' get it ready for import.
    
    Response1 = MsgBox("Have you already processed the original MBO CSV file downloaded from ABC?", vbYesNo, "MBO Import")
    If Response1 = vbNo Then
        Response3 = MsgBox("You will need to use the Excel file to process the MBO file first.")
        Exit Sub
    End If
    
    ' This section obtains the name of the file path using function getFileName(), which is located in
    ' the module RxDrugShortage
    
    MBOFileNamePath = ""
    
    MBOFileNamePath = getFileName() 'If file is selected, then the IfThen statement is exited.
    If MBOFileNamePath = "" Then
        tryAgain = MsgBox("No file name was selected. If you want to try again click Yes, if not click No", vbYesNo)
        If tryAgain = vbYes Then
            MBOFileNamePath = getFileName()
        Else
            Exit Sub
        End If
    End If

    
    ' First clear all data from the tmp table "tbl MBO", then transfer the spreadsheet already selected with the path in
    ' MBOFileNamePath
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM [tbl MBO];"
    DoCmd.SetWarnings True
    
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tbl MBO", MBOFileNamePath, True, , False

    ' Okay, now we enable the next button and disable this button
    Me.ImportRxScanInventory.Enabled = True
    Me.ImportMBO.Enabled = False
    Me.ImportRxScanInventory.SetFocus
    
    Response2 = MsgBox("Import completed successfully")

End Sub

