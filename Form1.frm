VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "AVI with Loop"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000C&
      Height          =   2175
      Left            =   1680
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   5
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Resume"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


Dim mssg As String * 255


   Public Function GetShortName(ByVal sLongFileName As String) As String
       Dim lRetVal As Long, sShortPathName As String, iLen As Integer
       'Set up buffer area for API function call return
       sShortPathName = Space(255)
       iLen = Len(sShortPathName)

       'Call the function
       lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
       'Strip away unwanted characters.
       GetShortName = Left(sShortPathName, lRetVal)
   End Function
Private Sub Command1_Click()
  Dim MInfo As String
  Dim ShortName As String
Screen.MousePointer = 11

CommonDialog1.CancelError = True
On Error GoTo EH1

CommonDialog1.Filter = "Video (*.avi)|*.avi"
CommonDialog1.Flags = &H80000 Or &H1000
CommonDialog1.ShowOpen

'#####################################################
'IMPORTANT!!!!!!!! all MCI Commands cannot see LONG FileNames!
'Therefore we must convert it to Short Name Format (This applies to MIDI, AVI and WAVs too)
  ShortName = GetShortName(CommonDialog1.filename)
'#####################################################



i = mciSendString("close all", 0&, 0, 0)
'#####################################################
Last$ = Picture1.hWnd & " Style " & &H40000000
ToDo$ = "open " & ShortName & " Type avivideo Alias video1 parent " & Last$
i = mciSendString(ToDo$, 0&, 0, 0)
i = mciSendString("put video1 window at 40 40 60 60", 0&, 0, 0)
'#####################################################

Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Screen.MousePointer = 0

Exit Sub

EH1:

Screen.MousePointer = 0
If Err = 32755 Then Err.Clear: Exit Sub
MsgBox Err.Description, vbExclamation, "ERR #" & Err
End Sub

Private Sub Command2_Click()
 i = mciSendString("play video1 from 0", 0&, 0, 0)
End Sub


Private Sub Command3_Click()
 i = mciSendString("pause video1", 0&, 0, 0)
End Sub


Private Sub Command4_Click()
 i = mciSendString("stop video1", 0&, 0, 0)
End Sub



Private Sub Command5_Click()
 i = mciSendString("resume video1", 0&, 0, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
 i = mciSendString("close video1", 0&, 0, 0)


End Sub


