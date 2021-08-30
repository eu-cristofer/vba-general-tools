Attribute VB_Name = "FunctionTriangleArea"
' Force explicit variable declaration
Option Explicit

Function TriangleArea(A_side As Integer, B_side As Integer, theta As Double)
Attribute TriangleArea.VB_Description = "Function to compute the area of a triangle."
Attribute TriangleArea.VB_ProcData.VB_Invoke_Func = " \n3"
'
' Function to compute the area of a triangle.
'
    Dim alpha As Double
    
    alpha = WorksheetFunction.Radians(theta)
    
    TriangleArea = 0.5 * A_side * B_side * Sin(alpha)

End Function
