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
    
    MBOFileNamePath = getFileName("Select new MBO file") 'If file is selected, then the IfThen statement is exited.
    If MBOFileNamePath = "" Then
        tryAgain = MsgBox("No file name was selected. If you want to try again click Yes, if not click No", vbYesNo)
        If tryAgain = vbYes Then
            MBOFileNamePath = getFileName("Select new MBO file")
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

Private Sub ImportRxScanInventory_Click()

    Dim Response1, Response2, Response3, tryAgain As Integer
    Dim RxScanInventoryFileNamePath, SQL1, SQL2, SQL3 As String
    
    ' DoCmd.RunSavedImportExport ("Import-MBO")
    ' Need to use code for this process at this time
    ' Also, may need to ask in dialog box if user has already run the Excel code on the original CSV file to
    ' get it ready for import.
    
    Response1 = MsgBox("Have you already downloaded both the Warehouse and non-Warehouse inventory files from RxScan?", vbYesNo, "MBO Import")
    If Response1 = vbNo Then
        Response3 = MsgBox("Please download both files first.")
        Exit Sub
    End If
    
    ' This part imports the inventory files, both Warehouse and Non-Warehouse. The first section gets the name of the
    ' Non-Warehouse file and then imports it. The seconce section gets the name of the Warehouse file and then imports it.
    
    ' This section obtains the name of the file path using function getFileName(), which is located in
    ' the module RxDrugShortage.
    
    RxScanInventoryFileNamePath = ""
    
    RxScanInventoryFileNamePath = getFileName("Select Non-Warehouse file") 'If file is selected, then the IfThen statement is exited.
    If RxScanInventoryFileNamePath = "" Then
        tryAgain = MsgBox("No file name was selected. If you want to try again click Yes, if not click No", vbYesNo)
        If tryAgain = vbYes Then
            RxScanInventoryFileNamePath = getFileName("Select Non-Warehouse file")
        Else
            Exit Sub
        End If
    End If
    
    ' First clear all data from the tmp table "tmpRxScanInventory", then transfer the spreadsheet already selected with the path in
    ' RxScanInventoryFileNamePath, and finally use a SQL statement to pick a subset of columns to be copied into the table "tbl RxScan Inventory Non-Warehouse"
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM [tmpRxScanInventory];"
    DoCmd.SetWarnings True
    
    ' Need to first import entire spreadsheet into a temporary table and then do a select of columns into the final table
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tmpRxScanInventory", RxScanInventoryFileNamePath, True, , False

    DoCmd.SetWarnings False
    SQL1 = "SELECT tmpRxScanInventory.LocationName, tmpRxScanInventory.NDC, tmpRxScanInventory.Description, tmpRxScanInventory.PackageSize, tmpRxScanInventory.PackageUnit, "
    SQL2 = "tmpRxScanInventory.QuantityOnHand, tmpRxScanInventory.AWC, tmpRxScanInventory.UOM, tmpRxScanInventory.TotalCost INTO [tbl RxScan Inventory Non-Warehouse] FROM tmpRxScanInventory;"
    SQL3 = SQL1 & SQL2
    DoCmd.RunSQL SQL3
    DoCmd.SetWarnings True
    

    ' Warehouse file

    RxScanInventoryFileNamePath = ""
    
    RxScanInventoryFileNamePath = getFileName("Select Warehouse file") 'If file is selected, then the IfThen statement is exited.
    If RxScanInventoryFileNamePath = "" Then
        tryAgain = MsgBox("No file name was selected. If you want to try again click Yes, if not click No", vbYesNo)
        If tryAgain = vbYes Then
            RxScanInventoryFileNamePath = getFileName("Select Warehouse file")
        Else
            Exit Sub
        End If
    End If
    
    ' First clear all data from the tmp table "tmpRxScanInventory", then transfer the spreadsheet already selected with the path in
    ' RxScanInventoryFileNamePath, and finally use a SQL statement to pick a subset of columns to be copied into the table "tbl RxScan Inventory Warehouse"
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM [tmpRxScanInventory];"
    DoCmd.SetWarnings True
    
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12, "tmpRxScanInventory", RxScanInventoryFileNamePath, True, , False

    DoCmd.SetWarnings False
    SQL2 = "tmpRxScanInventory.QuantityOnHand, tmpRxScanInventory.AWC, tmpRxScanInventory.UOM, tmpRxScanInventory.TotalCost INTO [tbl RxScan Inventory Warehouse] FROM tmpRxScanInventory;"
    SQL3 = SQL1 & SQL2
    DoCmd.RunSQL SQL3
    DoCmd.SetWarnings True

    ' Okay, now we enable the checkbox for Optional Imports and disable this button
    Me.chkOptionalImports.Enabled = True
    Me.ImportRxScanInventory.Enabled = False
    Me.RunReport.Enabled = True
    Me.chkOptionalImports.SetFocus

    Response2 = MsgBox("Imports completed successfully")

End Sub


Private Sub chkOptionalImports_Click()

    If Me.chkOptionalImports = True Then
        Me.ImportTransactionLog.Enabled = True
        Me.ImportNDCLookup.Enabled = True
    Else
        Me.ImportTransactionLog.Enabled = False
        Me.ImportNDCLookup.Enabled = False
    End If
    
End Sub

Private Sub RunReport_Click()

    DoCmd.OpenReport "Projected Supply vs Delay", acViewPreview
    

End Sub
