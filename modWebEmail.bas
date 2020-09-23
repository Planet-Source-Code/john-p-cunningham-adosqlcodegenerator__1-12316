Attribute VB_Name = "modWebEmail"
Option Explicit

Public Const URL = "http://users.ids.net/~johnpc"
Public Const email1 = "jerry_m_barnes@hotmail.com"

Public Const email2 = "johnpc@ids.net"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Sub gotoweb()
Dim Success As Long

Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub
Public Sub sendemail1()
Dim Success As Long

Success = ShellExecute(0&, vbNullString, "mailto:" & email1, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub
Public Sub sendemail2()
Dim Success As Long

Success = ShellExecute(0&, vbNullString, "mailto:" & email2, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub

