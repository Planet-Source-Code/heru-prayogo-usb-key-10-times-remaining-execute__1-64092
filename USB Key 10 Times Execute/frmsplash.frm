VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmsplash 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFC0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   5265
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   2775
         Left            =   -120
         TabIndex        =   6
         Top             =   -120
         Width           =   7095
         ExtentX         =   12515
         ExtentY         =   4895
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   240
      Top             =   480
   End
   Begin frmreg.Xp_ProgressBar Xp_Pro 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      Style           =   1
      ProgressLook    =   1
   End
   Begin VB.Label lblusbname 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label lblsearching 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Please Wait While Searching USB Device...."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5295
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PicPath As String
Private Sub cmdconnect_Click()
frmactivation.Show
Unload Me
End Sub
Private Sub cmdexit_Click()
End
End Sub
Private Sub Form_Load()
cmdExit.Enabled = False
cmdConnect.Enabled = False
PicPath = App.Path & "\splash.html"
WebBrowser1.Navigate PicPath
End Sub
Private Sub Form_Initialize()
 XPStyle            'XP Style form
End Sub
Private Sub Timer1_Timer()
Xp_Pro.value = Xp_Pro.value + 1
If Xp_Pro.value > 99 Then
Timer1.Enabled = False
Call CheckLabelUSB                               'Check Timer Function
End If
End Sub
Function CheckLabelUSB()
'This function Is for check USB Trademark I use San Disk 128 MB
Dim oWMINameSpace As SWbemServices
Dim oUSBDriveSet As SWbemObjectSet
Dim oUSBDrive As SWbemObject
Dim USBName As String                    'Get USB Name
Set oWMINameSpace = GetObject("winmgmts:")
Set oUSBDriveSet = oWMINameSpace.InstancesOf("Win32_DiskDrive")
For Each oUSBDrive In oUSBDriveSet
    On Error Resume Next
    USBName = oUSBDrive.Caption & ""            'USB Label Name
    lblusbname.Caption = USBName
    If lblusbname.Caption = "SanDisk Cruzer Mini USB Device" Then  'I use Sandisk USB 128 MB
    cmdExit.Enabled = False
    cmdConnect.Enabled = True
    lblsearching.Caption = "USB Deviced Is Found.You May Continued Use Software."
    Else
     cmdExit.Enabled = True
     cmdConnect.Enabled = False
     frmsplash.Caption = "Please Connect Your USB Device"
     lblsearching.Caption = "Please Connect Your USB Device.Press Exit Button To Continued." 'If Not My San Disk then
    End If
Next
End Function

