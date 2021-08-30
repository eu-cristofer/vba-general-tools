Attribute VB_Name = "CheckOSX"
Option Explicit

' Code by Cristofer Costa
' http://github.com/eu-cristofer/VBA_GeneralTools

Function bCheckOSX() As Boolean
'
' Function to check if OS is OSX
'
    If Application.OperatingSystem Like "*Mac*" Then
        bCheckOSX = True
    Else
        bCheckOSX = False
    End If

End Function
