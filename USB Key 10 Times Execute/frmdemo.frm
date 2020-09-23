VERSION 5.00
Begin VB.Form frmdemo 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Registration Demo"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   ClipControls    =   0   'False
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
   ScaleHeight     =   2160
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit Demo"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "License Information"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtserialdemo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtnamademo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtiddemo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Serial Number :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Machine ID :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmdemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
End
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
