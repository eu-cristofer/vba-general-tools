Attribute VB_Name = "CheckOSX"
Option Explicit

Function bCheckOSX() As Boolean

    If Application.OperatingSystem Like "*Mac*" Then
        bCheckOSX = True
    Else
        bCheckOSX = False
    End If

End Function
