Attribute VB_Name = "RxDrugShortage"
Option Compare Database

Sub SelectMBOFile()

    Dim filePath As String
    
    filePath = "T:\Documents\Projects\Pharmacy\Drug Shortage\Downloaded MBO\MBO_09_15_2016 As Delivered.csv"


End Sub

Function getFile()



End Function

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

