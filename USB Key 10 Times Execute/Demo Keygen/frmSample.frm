VERSION 5.00
Begin VB.Form frmKeygen 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Keygen"
   ClientHeight    =   2340
   ClientLeft      =   1320
   ClientTop       =   555
   ClientWidth     =   5145
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
   Icon            =   "frmSample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "&Generate ID"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Generate License Information by Heru Prayogo"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      Begin VB.TextBox txtserial 
         Height          =   615
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtnama 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtId 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Serial Number :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nama :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblHardwareID 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Machine ID :"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmKeygen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmdgen_Click()
Dim Text As String              '
Dim d As String                 'Initialize Variabel
Dim Key As String               '
Dim Enc As New clsTEA
Text = Enc.Encode64(txtId & txtnama)
MD5 = CalculateMD5(Text)
Length = Len(MD5)
d = ""
For I = 1 To Length             'Get Length from string
 Char$ = Mid(MD5, I, 1)
 Code = I + 1
 Code2 = I * Code
 Salt = I * 258880              'Can be changed for other Ex:Salt = i * 12345
Result = (((Asc(Char$) Xor Code) + ((Code2 * Code) + Salt)) Xor Code2)
Logans = Abs(Fix(Fix(Cos(Result)) * 255 + Sin(Result)))
Result = Result + ((Length And I) Or (Length Or I)) + Logans
d = d & Result
Next I
HASH = CalculateMD5(d)
HASH = StrReverse(HASH)
Key = ""
For I = 1 To Len(HASH)                  '
    Alph = Mid(HASH, I, 1)              '
    getrand = (I * 2 + Salt) Mod I      '
        If getrand Mod 2 = 0 Then       '
        Alph = LCase(Alph)              'Calculate Keygen routine
        Else                            '
        Alph = UCase(Alph)              '
        End If                          '
Key = Key & Alph
Next I
lasthash = CalculateMD5(Key)
txtserial.Text = "ACTDEMO" & UCase(lasthash)
End Sub
Private Sub Form_Initialize()
XPStyle
End Sub

