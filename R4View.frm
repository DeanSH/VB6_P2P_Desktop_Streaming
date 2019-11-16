VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form R4View 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "R4 Screen View"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   4920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   4920
   End
   Begin MSWinsockLib.Winsock Listener 
      Left            =   4560
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4560
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock NSock 
      Left            =   4560
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   4095
      TabIndex        =   12
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Host / Connect  ( For Purpose Of Screen Viewing )"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "R4View.frx":0000
         Left            =   1800
         List            =   "R4View.frx":0031
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Image Quality Controlled By Host! 1-100"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Host Session"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Start"
         Height          =   255
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Text            =   "1234"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Text            =   "85"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture Quality ??:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "The Pass Code ??:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Forwarded Port ?:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "The Network IP?:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ready To Use!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   2060
      Width           =   4335
   End
End
Attribute VB_Name = "R4View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sending As Boolean
Private Sending2 As Boolean
Private Receiving As Boolean
Private ShotSize As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Check1_Click()
If Check1.Value = 1 Then
Combo1.Enabled = True
Else
Combo1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
If Check1.Value = 1 Then Combo1.Enabled = False
Check1.Enabled = False
Command1.Enabled = False
Command2.Enabled = True
Sending = True
Sending2 = False
Receiving = False
On Error Resume Next
Winsock1.Close
NSock.Close
Listener.Close
Kill App.Path & "\R4SV.jpg"
If Check1.Value = 1 Then
Listener.LocalPort = Text2
Label2.Caption = "Listening.. Awaiting Viewer!"
Listener.Listen
Else
Label2.Caption = "Connecting To.. " & Text1
Winsock1.Connect Text1, Text2
End If
End Sub

Private Sub Command2_Click()
Check1.Enabled = True
If Check1.Value = 1 Then Combo1.Enabled = True
On Error Resume Next
Command2.Enabled = False
Command1.Enabled = True
Sending = True
Sending2 = False
Receiving = False
Winsock1.Close
NSock.Close
Listener.Close
Unload R4Screen
Label2.Caption = "All Activity Has Stopped!"
Kill App.Path & "\R4SV.jpg"
End Sub

Private Sub Form_Load()
Combo1.Text = "30"
R4Screen.Show
'If App.PrevInstance = True Then
'On Error Resume Next
'MsgBox "Opps, Sorry! Only One Copy Of R4 Screen View Can Be Open Running At A Time!!", , "Cannot Run Multiple Copies!"
'DoEvents
'End
'End If
End Sub

Private Function TakeAnSendShot()
If NSock.State = 7 Then
Sending = True
Sending2 = False
Set Picture1.Picture = CaptureScreen()
DoEvents
SaveJPG Picture1, App.Path & "\R4SV.jpg", Combo1
Else
Command2_Click
DoEvents
Label2.Caption = "Activity Stopped! Socket Error!"
End If
End Function

Private Function LoadShotReceived()
On Error Resume Next
If Winsock1.State = 7 Then
R4Screen.Show
R4Screen.Image1.Picture = LoadPicture(App.Path & "\R4SV.jpg")
DoEvents
Winsock1.SendData "Next||Pic"
Else
Command2_Click
DoEvents
Label2.Caption = "Activity Stopped! Socket Error!"
End If
End Function

Public Sub SendPhotoSize(pPath As String)
On Error Resume Next
If NSock.State = 7 Then
    Sending = False
    Debug.Print "Sending Ping..."
    NSock.SendData "Ping||Alive"
Else
Command2_Click
DoEvents
Label2.Caption = "Activity Stopped! Socket Error!"
End If
End Sub

Private Sub SendPhoto(pPath As String)
'On Error GoTo Down:
Dim i As Long
Dim B() As Byte
'==============
    Sending2 = True
    Debug.Print "Sending File..."

Open pPath For Binary As #1

For i = 1 To LOF(1)
    If Loc(1) >= LOF(1) Then
        Debug.Print "Sent File Completely!!"
        GoTo Down:
        Exit Sub
    End If
    If NSock.State <> 7 Then
        Debug.Print "File Socket Error!"
        GoTo Down:
        Exit Sub
    End If
    If LOF(1) < 1024 * 8 Or LOF(1) - Loc(1) < 1024 * 8 Then
        ReDim B(LOF(1) - Loc(1) - 1)
    Else
        ReDim B(1024 * 8 - 1)
    End If
    Get #1, , B
    'If Sending2 = False Then GoTo Down
    Debug.Print "Sending File Data..."
    Sending = False
    NSock.SendData B
    DoEvents
    wait (5)
    Pause "0.01"
Next
DoEvents
Down:
'On Error Resume Next
Debug.Print "File Closed"
Sending2 = False
Sending = True
Close #1

End Sub

Private Sub wait(iSecond As Integer)
On Error Resume Next
Dim Tme As Single
'================
Tme = Timer
Do Until (Timer - Tme >= iSecond) Or Sending = True
    DoEvents
Loop
End Sub

Private Sub Pause(ByVal interval As String)
On Error Resume Next
Dim wait   As Single
  
  wait = Timer
  
  Do While Timer - wait < CSng(interval$)
     DoEvents
 Loop
End Sub

Private Sub NSock_Close()
Check1.Enabled = True
If Check1.Value = 1 Then Combo1.Enabled = True
On Error Resume Next
Sending = True
Sending2 = False
NSock.Close
If Command2.Enabled = True Then Label2.Caption = "Connection Has Closed!!"
Command2.Enabled = False
Command1.Enabled = True
End Sub

Private Sub Listener_ConnectionRequest(ByVal requestID As Long)
'On Error Resume Next
NSock.Close
NSock.LocalPort = Text2
NSock.Accept requestID
DoEvents
Sending = False
wait (1)
Sending = True
Label2.Caption = "Requesting Pass From " & NSock.RemoteHostIP
NSock.SendData "Pass||Request"
End Sub

Private Sub NSock_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
NSock.GetData Data
Select Case Left(Data, 4)

Case "Pass"
'Pause "0.1"
Dim CheckPass As String
CheckPass = Split(Data, "||")(1)
If CheckPass = Text3 Then
Label2.Caption = "Connection Is Active!!!"
DoEvents
GoTo Feed
Else
Sending = False
NSock.SendData "Fail||"
wait "3"
Pause "0.001"
Check1.Enabled = True
If Check1.Value = 1 Then Combo1.Enabled = True
On Error Resume Next
Command2.Enabled = False
Command1.Enabled = True
NSock.Close
Label2.Caption = "Connection DC! Invalid Pass."
End If

Case "Next"
Feed:
Pause "0.01"
TakeAnSendShot
Exit Sub

Case "Send"
Pause "0.001"
SendPhoto App.Path & "\R4SV.jpg"
Exit Sub

Case "Ping"
Pause "0.001"
NSock.SendData "Kill||File"
Exit Sub
End Select
End Sub

Private Sub NSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Check1.Enabled = True
If Check1.Value = 1 Then Combo1.Enabled = True
On Error Resume Next
Sending = True
Sending2 = False
NSock.Close
If Command2.Enabled = True Then Label2.Caption = "Activity Stopped By Error!"
Command2.Enabled = False
Command1.Enabled = True
End Sub

Private Sub NSock_SendComplete()
If Sending = False Then Sending = True
End Sub

Private Sub Timer1_Timer()
If Winsock1.State = 7 Then
If ShotSize = 1 Then
Debug.Print "Receiving Ends"
Receiving = False
LoadShotReceived
Timer1 = False
Exit Sub
End If
ShotSize = ShotSize + 1
Else
Timer1 = False
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Winsock1.State = 7 Then Winsock1.SendData "Send||"
Timer2 = False
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Timer3 = False
SendPhotoSize App.Path & "\R4SV.jpg"
End Sub

Private Sub Winsock1_Close()
Check1.Enabled = True
If Check1.Value = 1 Then Combo1.Enabled = True
On Error Resume Next
Winsock1.Close
Unload R4Screen
If Command2.Enabled = True Then Label2.Caption = "Connection Has Closed!!"
Command2.Enabled = False
Command1.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
Dim FileP As String
Winsock1.GetData Data
DoEvents

If Receiving = True Then
ShotSize = 0
FileP = App.Path & "\R4SV.jpg"
Open FileP For Binary As #2
Put #2, LOF(2) + 1, Data
Close #2
DoEvents
Timer1 = True
Exit Sub
End If

Select Case Left(Data, 4)
Case "Ping"
Pause "0.001"
Label2.Caption = "Connection Is Active!!!"
Winsock1.SendData "Ping||"
Exit Sub
Case "Kill"
Pause "0.001"
Kill App.Path & "\R4SV.jpg"
Receiving = True
Timer2 = True
Exit Sub
Case "Pass"
Pause "0.001"
Label2.Caption = "Pass Code Requested..."
Winsock1.SendData "Pass||" & Text3
Exit Sub
Case "Fail"
Check1.Enabled = True
If Check1.Value = 1 Then Combo1.Enabled = True
On Error Resume Next
Command2.Enabled = False
Command1.Enabled = True
Winsock1.Close
Label2.Caption = "Connection DC! Invalid Pass."
Exit Sub
End Select
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Check1.Enabled = True
If Check1.Value = 1 Then Combo1.Enabled = True
On Error Resume Next
Winsock1.Close
Unload R4Screen
If Command2.Enabled = True Then Label2.Caption = "Activity Stopped By Error!"
Command2.Enabled = False
Command1.Enabled = True
End Sub
