VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dreams PAK Self-Extracter"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      ItemData        =   "Form1.frx":1272
      Left            =   105
      List            =   "Form1.frx":1274
      TabIndex        =   8
      Top             =   1665
      Width           =   3825
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Help"
      Height          =   350
      Left            =   3990
      TabIndex        =   7
      Top             =   1980
      Width           =   990
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&About"
      Height          =   350
      Left            =   3990
      TabIndex        =   6
      Top             =   1545
      Width           =   990
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   350
      Left            =   3990
      TabIndex        =   5
      Top             =   1125
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Extract"
      Height          =   350
      Left            =   3990
      TabIndex        =   4
      Top             =   705
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Other Folder"
      Height          =   345
      Left            =   2700
      TabIndex        =   2
      Top             =   1035
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   165
      TabIndex        =   1
      Text            =   "C:\WINDOWS\TEMP\"
      Top             =   1065
      Width           =   2490
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   195
      TabIndex        =   10
      Top             =   2325
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current Files to Extract = "
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   210
      TabIndex        =   9
      Top             =   1395
      Width           =   1770
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "Form1.frx":1276
      Top             =   105
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Extract All Files To"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   810
      Width           =   1305
   End
   Begin VB.Label Label1 
      Height          =   405
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Folder_Name As String
Dim StrBuff As String


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
Sub SaveData(DataBlock As String, zFilename As String)
Open AddBackSlash(Text1.Text) & zFilename For Binary As #2
    Put #2, , Encode(DataBlock)
    Close #2
    
End Sub

Sub Extract()
Dim M_Filename As Collection
Dim M_Count As Long
Dim Xpos As Long
Dim YPos As Long
Dim Zpos As Long
Dim StrLine As String
Dim StrSplit As String
Dim G As String
Dim TBuffer As String

    Set M_Filename = New Collection
    
    Xpos = InStr(StrBuff, "<File>")
        If Xpos Then
            StrLine = Mid(StrBuff, Xpos + 6, Len(StrBuff))
        End If
        
        For M_Count = 1 To Len(StrLine)
            ch = Mid(StrLine, M_Count, 1)
             StrSplit = StrSplit & ch
              If InStr(StrSplit, "|") Then
                  G = Left(StrSplit, Len(StrSplit) - 1): StrSplit = ""
                  M_Filename.Add (G)
                  
              End If
        Next
        
        ' Extract all files here
        
        For YPos = 1 To Len(StrBuff)
         ch2 = Mid(StrBuff, YPos, 1)
            TBuffer = TBuffer & ch2
             If InStr(TBuffer, "<--POS-->") Then
                 TBuffer = Left(TBuffer, Len(TBuffer) - 9)
                    Zpos = Zpos + 1
                        SaveData TBuffer, M_Filename(Zpos)
                TBuffer = ""
             End If
          Next
          Label4.Caption = StrConv("All " & Zpos & " files have been extracted to " & Text1.Text, vbProperCase)
          
          StrBuff = ""
          G = ""
          Xpos = 0: YPos = 0: Zpos = 0: M_Count = 0
          
End Sub

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
Function AddBackSlash(Pathname As String) As String
Dim TBackSlash As String

If Not Right(Pathname, 1) = "\" Then
    TBackSlash = Pathname & "\"
    Else
    TBackSlash = Pathname
End If
    AddBackSlash = TBackSlash
    
End Function
Function CheckFileIsSFX(Filename As String)
Dim MainData As String
Dim M_Start As Long
Dim M_FileStart As Long
Dim File_Names As String
Dim M_Split As String
Dim M_Count As Long
Dim Data As String


    Open Filename For Binary As #1
     MainData = Space(LOF(1))
     Get #1, , MainData
     Close #1
     
      M_Start = InStr(MainData, "<SFX>")
        If M_Start Then
             Data = Mid(MainData, M_Start + 5, Len(MainData))
             M_FileStart = InStr(Data, "<File>")
                If M_FileStart Then
                    File_Names = Mid(Data, M_FileStart + 6, Len(Data))
                    
                    For M_Count = 1 To Len(File_Names)
                        ch = Mid(File_Names, M_Count, 1)
                        M_Split = M_Split & ch
                             
                             If InStr(M_Split, "|") Then
                                List1.AddItem Left(M_Split, Len(M_Split) - 1): M_Split = ""
                             End If
                            Next
                             Label3.Caption = Label3.Caption & List1.ListCount
                                StrBuff = Mid(MainData, M_Start + 5, Len(MainData))
                                'Extract
                    End If
                    
            Else
            MsgBox "Inviald SFX File", vbCritical
            End
            End If
            
            M_Count = 0
            File_Names = ""
            MainData = ""
            M_Split = ""
            Data = ""
            
            
End Function
Sub CenterForm(Frm As Form)
With Frm
    .Top = (Screen.Height - Frm.Height) / 2
    .Left = (Screen.Width - Frm.Width) / 2
End With
    
End Sub

Private Sub Command1_Click()
Text1.Text = GetFolder(Form1.hwnd, "&Please Select A Folder")
 If Len(Text1.Text) < 3 Then
    Text1.Text = Folder_Name
    End If
    
End Sub

Private Sub Command2_Click()
AddBackSlash Text1.Text
    Extract
    
End Sub

Private Sub Command3_Click()
End

End Sub

Private Sub Command4_Click()
Dim msg As String
Module1.WindowOnTop Form1.hwnd, False
    msg = msg & "This is a Self Extracting exe" & vbCrLf
    msg = msg & "Createed with the PAK Self Extracter Builder 1.1" & vbCrLf
    msg = msg & "This program is part of Dreams PAK-Builder and PAK - Extracter Program" & vbCrLf
    msg = msg & "Created by Ben Jones"
        MsgBox StrConv(msg, vbProperCase), vbInformation
Module1.WindowOnTop Form1.hwnd, True

End Sub

Private Sub Command5_Click()
Dim msg As String
Module1.WindowOnTop Form1.hwnd, False
    msg = msg & "To Extarct all the files form this exe" & vbCrLf
    msg = msg & "Click the [Other Folder] and select" & vbCrLf
    msg = msg & "the path were you whish to extarct the files then click the Extract button"
    MsgBox StrConv(msg, vbProperCase), vbInformation
        Module1.WindowOnTop Form1.hwnd, True
        
End Sub

Private Sub Form_Load()
Module1.WindowOnTop Form1.hwnd, True

CenterForm Form1
    Folder_Name = Text1.Text
    
 CheckFileIsSFX AddBackSlash(App.Path) & App.EXEName & ".exe"
 
Form1.Caption = Form1.Caption & " - " & App.EXEName & ".exe"
    Label1.Caption = "To extract all files form  " & App.EXEName & ".exe" & _
    "to the folder in the edit field below just Click Extract"
    
    
End Sub

