Attribute VB_Name = "Module1"

Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" ( _
     ByVal hwnd As Long, _
     ByVal szApp As String, _
     ByVal szOtherStuff As String, _
     ByVal hIcon As Long) As Long

