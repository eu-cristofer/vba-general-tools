Attribute VB_Name = "SubCategorizeFunction"
' Force explicit variable declaration
Option Explicit

Sub CategorizeFunctionTriangleArea()
'
' Categorizes the function Triangle Area
'
    ' Function arguments description
    Dim strDescription(1 To 3) As String
    strDescription(1) = "A side of  the triangle."
    strDescription(2) = "B side of  the triangle."
    strDescription(3) = "Theta angle in degrees between A e B."
   
    '  Function's categorization
    Application.MacroOptions _
        Macro:="TriangleArea", _
        Description:="Function to compute the area of a triangle.", _
        Category:=3, _
        ArgumentDescriptions:=strDescription

End Sub

