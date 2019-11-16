VERSION 5.00
Begin VB.Form R4Screen 
   BackColor       =   &H00FFFFFF&
   Caption         =   "R4 Screen View (Viewing: 127.0.0.1)"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   6615
      Left            =   0
      Picture         =   "R4Screen.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "R4Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
Image1.Height = Me.Height - 570
Image1.Width = Me.Width - 240
End Sub
