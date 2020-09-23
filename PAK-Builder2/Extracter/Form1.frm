VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PAK Extracter 2"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4635
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   5460
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
         Caption         =   "&Extract"
         Height          =   390
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2865
         Width           =   1230
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   1875
         Pattern         =   "*.pkg"
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
         TabIndex        =   10
         Top             =   4140
         Width           =   5280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Save Files To"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   3570
         Width           =   975
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
Sub SaveFile(StrData As String, TFilename As String)
Open Text1.Text & TFilename For Binary As #3
    Put #3, , Encode(StrData)
    Close #3
    
End Sub
Private Sub Command1_Click()
Dim Z_Filename As String
Dim Counter, XPos, YPos, FileN As Long
Dim StrLine, StrBuffer As String
Dim s As String
Dim SplitString As String
Dim Filename As Collection
Set Filename = New Collection

Z_Filename = FixPath(File1) & File1.Filename
On Error Resume Next

If Len(Text1.Text) < 5 Then
    MsgBox "Please select a folder to extract the files to", vbInformation: Exit Sub
End If

Open Z_Filename For Binary As #1

s = Space(LOF(1))
Get #1, , s
Close #1

' used to find the filenames

FileN = InStr(s, "<File>") ' Filename Name Tag
If FileN Then
       StrLine = Mid(s, FileN + 6, Len(s))
End If

' Used to Split up all the filenames

For XPos = 1 To Len(StrLine)
ch1 = Mid(StrLine, XPos, 1)
     SplitString = SplitString & ch1
     If InStr(SplitString, "|") Then
        g = Left(SplitString, Len(SplitString) - 1): SplitString = ""
        Filename.Add (g)
     End If
Next

'Extarct all the files

    For YPos = 1 To Len(s) ' Get the length of the file
        ch2 = Mid(s, YPos, 1) ' Grab Each char
        StrBuffer = StrBuffer & ch2
         If InStr(StrBuffer, "<--POS-->") Then ' This is were each file will end
         StrBuffer = Left(StrBuffer, Len(StrBuffer) - 9) ' Remove File Pointer
         Counter = Counter + 1 ' This tell use how many file are within the pak file
         TimeOut 0.1
         Label3.Caption = "Extracting " & Filename(Counter)
         SaveFile StrBuffer, Filename(Counter) ' Save each file
         StrBuffer = "" 'We Clear this out affter each file has been saved
         End If
    Next
    Label3.Caption = "All " & Counter & " Files have been Extracted"
    
    If Err Then
        MsgBox "Please Select a filename form the list", vbInformation: Label3.Caption = ""
    End If

End Sub

Private Sub Command2_Click()
Dim Msg As String
Msg = Msg + "Dreams PAK File Extracter 2" + vbCrLf
Msg = Msg + "This is a part of the PAK File Builder Program" + vbCrLf
Msg = Msg + "Programmed by Ben Jones" + vbCrLf

MsgBox Msg, vbInformation


End Sub

Private Sub Command3_Click()
Dim Answer
    Answer = _
        MsgBox("Do you wish to quit this program now", _
        vbYesNo)
        
        If Answer = vbYes Then
            End
            Else
            End If
            
End Sub

Private Sub Command4_Click()
    Text1.Text = GetFolder(Form1.hWnd, "Select a folder to Extract files to")
    
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
    MkDir App.Path & "\Files"
        Text1.Text = App.Path & "\Files\"
        
End Sub

