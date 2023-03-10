Option Explicit
' =============================================================================
' Export VBA Modules to file
' 
' Date  : 2019/01/21
' Auther: Kunihisa Abukawa (@kabukawa)
' 
' Require:
'   Microsoft Excel
'   Microsoft Windows
' Usage:
'   Drop Files on this if you want export VBA Modules to file.
'     or
'   cscript [Excel file names...]
' =============================================================================

' Check Arguments(Excel File Path) is not set.
' -----------------------------------------------------------------------------
If WScript.Arguments.Count = 0 Then
    MsgBox "Drop Files on this if you want export VBA Modules to file.", _
            vbOkOnly + vbInformation, "Information"
    WScript.Quit
End If

' Export VBA Modules
' -----------------------------------------------------------------------------
With CreateObject("Excel.Application")
    Dim fileName
    For Each fileName In WScript.Arguments
        With .Workbooks.Open(fileName,0,True,,,,True)
            outPutModule fileName & "_modules", .VBProject
            .Close False
        End With
    Next
End With

MsgBox "All Excel VBA Modules exported to file.", _
        vbOkOnly + vbInformation, "Information"
WScript.Quit


' =============================================================================
' Subroutines/Functions
' =============================================================================

' Clean output folder
' -----------------------------------------------------------------------------
Sub CleanOutput(dirName)
    With CreateObject("Scripting.FileSystemObject")
        If .FolderExists(dirName) Then
            .DeleteFolder dirName, True
        End If
        .CreateFolder dirName
    End With
End Sub

' Output VBA Modules to file.
' -----------------------------------------------------------------------------
Sub outPutModule(dirName, objProject)
    Dim objModule
    CleanOutput dirName
    If objProject.Protection = 0 Then
        For Each objModule In objProject.VBComponents
            objModule.Export BuildExpPath(dirName, objModule)
        Next
    End If
End Sub

' Build export module file path.
' -----------------------------------------------------------------------------
Function BuildExpPath(dirName, objModule)
    BuildExpPath = dirName & "\" & objModule.Name & GetExt(objModule.Type)
End Function

' Get file extention from module type.
' -----------------------------------------------------------------------------
Function GetExt(objType)
    Select Case objType
        Case 1      : GetExt = ".bas"
        Case 3      : GetExt = ".frm"
        Case Else   : GetExt = ".cls"
    End Select
End Function