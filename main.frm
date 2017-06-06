VERSION 5.00
Begin VB.Form mainForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1335
   FillColor       =   &H00E0E0E0&
   FillStyle       =   0  'Solid
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "main.frx":C0DE
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton copyBtn 
      Caption         =   "Copy URL"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1350
      Width           =   1335
   End
   Begin VB.Image failImage 
      Height          =   1350
      Left            =   0
      Picture         =   "main.frx":EE38
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image tickImage 
      Height          =   1350
      Left            =   0
      Picture         =   "main.frx":FA22
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim url As String

Private Function getUrl() As String
ip = Resolve("mahya.ydns.eu", Resolve("ns1.ydns.eu", "4.2.2.4"))
If ip = "" Then
getUrl = ""
Exit Function
End If
getUrl = "http://" + ip + ":8092/sys/fa/neoclassic/login/login"
End Function



Private Sub copyBtn_Click()
Clipboard.SetText (url)
End Sub

Private Sub Form_Load()
MakeTopMost (mainForm.hwnd)
End Sub

Private Sub Timer_Timer()
url = getUrl()
'MsgBox url
If url = "" Then
mainForm.Hide
MsgBox "Connecting to dns server failed!", vbOKOnly + vbCritical
End
failImage.Visible = True
Else
openInBrowser (url)
tickImage.Visible = True
copyBtn.Enabled = True
End If

Timer.Enabled = False

End Sub
