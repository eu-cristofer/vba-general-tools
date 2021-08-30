Attribute VB_Name = "Tool_ModulesExport"
Option Explicit
'
' Development tool to export all VBAComponents
' from a VBA project.
'
' Before run this macro change ***LABEL*** for a valid user folder
' or set the constant strPath with a valid address
'
Public Sub ExportModules()
    Const strWB As String = "Book1.xlsm"
    Const strPath As String = "/Users/***LABEL***/Desktop/ExportedFiles"
    Dim strFile As String
    Dim i As Integer
    Dim oVBP As Object ' VBProject
    Dim oComp As Object ' VBComponent
    Dim oSheet As Object
    Dim oNewWB As Object

    MkDir strPath ' Create output folder

    Set oVBP = Application.Workbooks(strWB).VBProject

    i = 1
    Set oNewWB = Application.Workbooks.Add
    Set oSheet = oNewWB.Sheets.Add
    oSheet.Name = "ModList"

    ' Select each kind of component
    For Each oComp In oVBP.VBComponents
        Select Case oComp.Type
            Case 1 ' vext_ct_StdModule
                strFile = oComp.Name & ".bas"
            Case 2 ' vext_ct_ClassModule
                strFile = oComp.Name & ".cls"
            Case 3 ' vbext_ct_MSForm
                strFile = oComp.Name & ".frm"
            Case 100 ' vbext_ct_Document
                strFile = oComp.Name & ".xlsx"
        End Select

        ' Export component
        If oComp.Type <> 100 Then
            oComp.Export strPath & Application.PathSeparator & strFile
            oSheet.Cells(i, 1) = strFile
            i = i + 1
        End If
    Next oComp
End Sub
