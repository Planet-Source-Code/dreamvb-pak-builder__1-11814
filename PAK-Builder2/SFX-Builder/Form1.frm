VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Self Extracter"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   3450
      TabIndex        =   6
      Top             =   1410
      Width           =   1050
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Help"
      Height          =   315
      Left            =   2325
      TabIndex        =   5
      Top             =   1410
      Width           =   1050
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&About"
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   1410
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Build"
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   1410
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Borowse"
      Height          =   350
      Left            =   3540
      TabIndex        =   2
      Top             =   630
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1185
      TabIndex        =   1
      Top             =   660
      Width           =   2250
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dreams PAK Self Extracter Builder 1.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      TabIndex        =   7
      Top             =   150
      Width           =   3750
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":27A2
      Top             =   30
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   -75
      X2              =   2850
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   -90
      X2              =   2835
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PAK Filename"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   690
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PAK_Data As String
Dim SFX_Data As String
Dim MainData As String

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOPENFILENAME As OPENFILENAME) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Sub CenterForm(Frm As Form)
With Frm
    .Top = (Screen.Height - Frm.Height) / 2
    .Left = (Screen.Width - Frm.Width) / 2
End With
    
End Sub
Public Function OpenFile(mTitle, mFileType, mFileExt As String) As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = mFileType + Chr(0) + mFileExt
        ofn.lpstrFile = Space(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path
        ofn.lpstrTitle = mTitle
        ofn.FLAGS = 0
       
        a = GetOpenFileName(ofn)
        If (a) Then
                OpenFile = Trim(ofn.lpstrFile)
        End If
        
 End Function
Private Sub Command1_Click()
    Text1.Text = OpenFile("Open PAK File", "PKG Files", "*.pkg")
    
End Sub
Function GetFileNameFromPath(Pathname As String) As String
Dim Xpos As Integer
 
 If Right(Pathname, 1) = "\" Then
   Exit Function
   Else
   For m_count = 1 To Len(Pathname)
    ch = Mid(Pathname, m_count, 1)
     
     If ch = "\" Then
        Xpos = m_count
     End If
     
    Next
     GetFileNameFromPath = Mid(Pathname, Xpos + 1, Len(Pathname))
     End If
     
End Function

Private Sub Command2_Click()
Dim SFX_Header As String
Dim NewName As String

Module1.WindowOnTop Form1.hwnd, False

SFX_Header = "<SFX>"
 
 If Len(Text1.Text) = 0 Then MsgBox "You must first slect a PAK filename", vbCritical: Exit Sub
  If FileExists(App.Path & "\SFX\DD1.SFX") Then
    NewName = GetFileNameFromPath(Text1.Text)
        NewName = Left(NewName, Len(NewName) - 3) + "exe"
        
    Open Text1.Text For Binary As #1
    Open App.Path & "\SFX\DD1.SFX" For Binary As #2
    Open "c:\PakSfx\" & NewName For Binary As #3
    
        PAK_Data = Space(LOF(1))
        SFX_Data = Space(LOF(2))
        Get #1, , PAK_Data
        Get #2, , SFX_Data
        
        Close #1
        Close #2
    End If
    
    MainData = MainData & SFX_Data & SFX_Header & PAK_Data
    Put #3, , MainData
    Close #3
    
    SFX_Data = ""
    PAK_Data = ""
    MainData = ""
    NewName = ""
        MsgBox "Self Extractable has beem saved to C:\PakSfx\", vbInformation
        Module1.WindowOnTop Form1.hwnd, True
        
End Sub

Private Sub Command3_Click()
Module1.WindowOnTop Form1.hwnd, False
    MsgBox "Dreams PAK Self Extracter Builder 1.1 Free Add-on", vbInformation
Module1.WindowOnTop Form1.hwnd, True

End Sub
Public Function FileExists(ByVal Filename As String) As Integer
If Dir(Filename) = "" Then FileExists = 0 Else FileExists = 1

End Function
Private Sub Command4_Click()
Dim Msg As String
Module1.WindowOnTop Form1.hwnd, False
    Msg = Msg + "To Build the selft extarcter click the Borowse" & vbCrLf
    Msg = Msg + "and select the PAK file you whish to compile" & vbCrLf
    Msg = Msg + "Then click the build button"
        MsgBox StrConv(Msg, vbProperCase), vbInformation
        Module1.WindowOnTop Form1.hwnd, True
        
End Sub

Private Sub Command5_Click()
End
 
End Sub



Private Sub Form_Load()
On Error GoTo terr
Module1.WindowOnTop Form1.hwnd, True
 CenterForm Form1
  
  If Not Module1.FolderExists("C:\PakSfx\") Then
    MkDir "C:\PakSfx\"
  End If
terr:
Err.Clear


End Sub

Private Sub Form_Resize()
    Line1.X2 = Form1.Width
    Line2.X2 = Form1.Width
    
End Sub
