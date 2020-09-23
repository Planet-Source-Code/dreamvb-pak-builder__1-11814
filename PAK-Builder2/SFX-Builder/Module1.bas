Attribute VB_Name = "Module1"
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Function WindowOnTop(mHwnd As Long, OnTop As Boolean)
If OnTop = True Then
     SetWindowPos mHwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
Else
     SetWindowPos mHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
End If

End Function
Public Function FolderExists(ByVal Foldername As String) As Integer
If Dir(Foldername, vbDirectory) = "" Then FolderExists = 0 Else FolderExists = 1

End Function
