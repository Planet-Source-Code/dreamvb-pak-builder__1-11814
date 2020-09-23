VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PAK Builder 2"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4830
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   5460
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1260
         TabIndex        =   10
         Text            =   "MyPakfile.pkg"
         Top             =   3915
         Width           =   1680
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "...."
         Height          =   330
         Left            =   4455
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3540
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   3540
         Width           =   3180
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "E&xit"
         Height          =   390
         Left            =   2655
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2865
         Width           =   1230
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&About"
         Height          =   390
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2865
         Width           =   1230
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Compile"
         Height          =   390
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2865
         Width           =   1230
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   1875
         Pattern         =   "*.txt"
         TabIndex        =   3
         Top             =   270
         Width           =   3465
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   90
         TabIndex        =   2
         Top             =   765
         Width           =   1755
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   285
         Width           =   1755
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   90
         TabIndex        =   12
         Top             =   4320
         Width           =   5280
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PAK Filename"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   3945
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Save Folder"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   3570
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Data As String

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type
Function Encode(TString As String) As String
Dim I_Count As Integer
    For I_Count = 1 To Len(TString)
        letter = Mid(TString, I_Count, 1)
        Mid(TString, I_Count, 1) = Chr(Asc(letter) Xor Hex(128))
    Next
        Encode = TString

End Function
Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
      Offset = InStr(RetPath, Chr$(0))
      GetFolder = Left$(RetPath, Offset - 1)
    End If
End Function
Function FixPath(Flb As FileListBox) As String
Dim TPath As String

If Right(Flb.Path, 1) = "\" Then
    TPath = Flb.Path
    Else
    TPath = TPath & Flb.Path & "\"
End If
    FixPath = TPath
    
End Function
Private Function TimeOut(nSecons As Single)
 Dim m_Sec
  m_Sec = Timer
  Do While Timer - m_Sec < nSecons
     DoEvents
     Loop
     
End Function

Sub BuildPakFile(PakFilename As String)
Open PakFilename For Binary As #2
    Put #2, , Data
    Close #2
    
End Sub
Private Sub Command1_Click()
Dim StrBuff As String
Dim F_Name As String
 On Error Resume Next
If Len(Text1.Text) < 3 Or Len(Text2.Text) < 4 Then MsgBox "Please select a save path name or filename", vbCritical: Exit Sub

For i = 0 To File1.ListCount - 1
a = FixPath(File1) & File1.List(i)
    F_Name = F_Name & File1.List(i) & "|"
TimeOut 0.1
        Label3.Caption = "Adding " & a

    Open a For Binary As #1
        Data = Space(LOF(1))
        Get #1, , Data
        StrBuff = StrBuff & Encode(Data) & "<--POS-->"
        Data = StrBuff & "<File>" & F_Name
    Close #1
Next
    BuildPakFile Text1.Text & Text2.Text
        Label3.Caption = "All " & i & " files have been added"
 If Err Then Err.Clear
 
End Sub

Private Sub Command2_Click()
Dim Msg As String
Msg = Msg + "Dream PAK Builder Program" + vbCrLf
Msg = Msg + "This is the main PAK Builder program" + vbCrLf
Msg = Msg + "Programmed by Ben Jones" + vbCrLf
    MsgBox Msg, vbInformation
    

End Sub

Private Sub Command3_Click()
Dim Answer
    Answer = _
        MsgBox("Do you wish to quit the program now", _
        vbYesNo)
        
            If Answer = vbYes Then
                End
            Else
            End If
            
End Sub

Private Sub Command4_Click()
    Text1.Text = GetFolder(Form1.hWnd, "Please select Your Folder")
    
End Sub



Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    Dir1.Path = Drive1.Drive
    If Err Then Err.Clear
    
    
End Sub

Private Sub Form_Load()
On Error Resume Next
    MkDir App.Path & "\PakFiles"
        Text1.Text = App.Path & "\PakFiles\"
        
End Sub

