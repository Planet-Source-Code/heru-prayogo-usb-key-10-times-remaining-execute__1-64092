VERSION 5.00
Begin VB.Form frmactivation 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activation System 10 Times Execute"
   ClientHeight    =   3105
   ClientLeft      =   1305
   ClientTop       =   540
   ClientWidth     =   5895
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
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset To Zero"
      Height          =   375
      Left            =   3960
      TabIndex        =   28
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4200
      TabIndex        =   27
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdRunDemo 
      Caption         =   "&Run Demo"
      Height          =   375
      Left            =   2520
      TabIndex        =   26
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register Software"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Width           =   2295
   End
   Begin frmreg.Xp_ProgressBar Xp_Pro 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      ProgressLook    =   1
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "License Information"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   5655
      Begin VB.Label chkserial 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label chkend 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   5295
      End
      Begin VB.Label chknama 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label chkmachineid 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Serial Number :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nama :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Machine ID :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label chkbegin 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Registration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5655
      Begin VB.Timer TimerDbg 
         Interval        =   10000
         Left            =   120
         Top             =   1320
      End
      Begin VB.TextBox txtnama 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txtserial 
         Height          =   555
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtId 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblRegistrasiKey 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Serial Number :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
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
   Begin VB.Label lblusbname 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   5520
      Width           =   4815
   End
   Begin VB.Label proexecutions 
      BackColor       =   &H00FFC0C0&
      Caption         =   "---"
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Executns:"
      Height          =   255
      Left            =   1680
      TabIndex        =   21
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Executns:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label rawexecutions 
      BackColor       =   &H00FFC0C0&
      Caption         =   "---"
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "out of executions"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblcdname 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6480
      Width           =   1815
   End
End
Attribute VB_Name = "frmactivation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code is for people to who using keygeneration routines to
'protect their application.It used WMI Scripts can be found at microsoft homepage.
'Besides MD5 Secure hash algorithma and Tean Encoding also used Enigma Mode Encryption.

Private Sub cmdexit_Click()
Unload Me           'Exit Form
End Sub
Private Sub cmdregister_Click()
If txtnama.Text = "" Or txtserial.Text = "" Then
MsgBox "Please Insert A Name And Serial Number.", vbInformation, "Information"
Exit Sub
End If
Dim Text As String
Dim d As String
Dim Key As String                   'Initialize variabel
Dim Enc As New clsTEA
On Error Resume Next
Text = Enc.Encode64(txtId & txtnama)                'Encoding with TEAN then calculate between Machine id and Name
MD5 = CalculateMD5(Text)
Length = Len(MD5)                   'Take from the length of MD5
d = ""
For i = 1 To Length
 Char$ = Mid(MD5, i, 1)
 Code = i + 1
 Code2 = i * Code
 Salt = i * 258880                  'Can be changed Ex:Salt = i * 12345
Result = (((Asc(Char$) Xor Code) + ((Code2 * Code) + Salt)) Xor Code2)
Logans = Abs(Fix(Fix(Cos(Result)) * 255 + Sin(Result)))     'Logans is for the results
Result = Result + ((Length And i) Or (Length Or i)) + Logans
d = d & Result
Next i
HASH = CalculateMD5(d)
HASH = StrReverse(HASH)
Key = ""
For i = 1 To Len(HASH)                  '
    Alph = Mid(HASH, i, 1)              '
    getrand = (i * 2 + Salt) Mod i      '
        If getrand Mod 2 = 0 Then       '
        Alph = LCase(Alph)              'Calculate Hash Keygen
        Else                            '
        Alph = UCase(Alph)              '
        End If                          '
Key = Key & Alph
Next i
lasthash = CalculateMD5(Key)
If txtserial.Text = "ACTDEMO" & UCase(lasthash) Then          'If valid the goto message
    MsgBox "Registration information correct." & _
   vbCrLf & "Please Restart Demo Activation System.", vbInformation + vbOKOnly, "Registered"
Close #1
Open App.Path & "\" & "license.key" For Output As #1            '
Print #1, "--- Begin of Demo Licensed by Heru Prayogo ---"       '
Print #1, EnigmaEncrypt(txtId.Text)                              '>>>>>> Print Out License <<<<<<<<<
Print #1, EnigmaEncrypt(txtnama.Text)                          '
Print #1, EnigmaEncrypt(txtserial.Text)                         '
Print #1, "--- End of Demo Licensed v1.1 by Heru Prayogo ---"    '
Close #1
End
Else                        'If failed registration then goto message
   MsgBox "Registration Failed. Please Check Your Information.", vbCritical, ("Registration")
Kill App.Path & "\" & "license.key"
    End If
End Sub
Private Sub cmdreset_Click()
Close #1
Open "C:\windows\system32\write.exe" For Binary As #1
Put #1, FileLen("C:\windows\system32\write.exe"), 0         'Open wordpad for write binary
Xp_Pro.value = 0
Label2.Caption = "0"
End Sub
Private Sub cmdrundemo_Click()
Clipboard.SetText Label2.Caption
If Xp_Pro.value < 9 Then
frmdemo.Show
Me.Hide             'Check for license can`t be unload
Else
MsgBox "The trial has expired.Please register the software to continue using it.", vbCritical, "Trial Expired"
End
End If
frmdemo.txtiddemo.Text = txtId.Text
frmdemo.txtnamademo.Text = "DEMO VERSION"
frmdemo.txtserialdemo.Text = "UNREGISTERED VERSION"
End Sub
Private Sub Form_Load()
Dim oWMINameSpace As SWbemServices
Dim oUSBDriveSet As SWbemObjectSet
Dim oUSBDrive As SWbemObject
Dim USBSerial As String      'Get String CDSerial
Dim USBName As String
Set oWMINameSpace = GetObject("winmgmts:")
Set oUSBDriveSet = oWMINameSpace.InstancesOf("Win32_USBHub")
For Each oUSBDrive In oUSBDriveSet
    On Error Resume Next
    USBSerial = oUSBDrive.PNPDeviceID & ""              'Use Pnp Device ID to calculate machine id
    txtId.Text = KeyGen(USBSerial, "DEMOMachineIDKey", 3) 'Machine ID Serial
Next
Call CheckValidKey                      'Try to open license.key and Xor it
'===============================
'start reading write.exe file
'===============================
Close #1
Dim readtrialnumber As Byte
On Error GoTo Err
Open "C:\windows\system32\write.exe" For Binary As #1
Get #1, FileLen("C:\windows\system32\write.exe"), readtrialnumber
Close #1
Clipboard.Clear
Clipboard.SetText readtrialnumber
rawexecutions.Caption = Chr(Clipboard.GetText)
Clipboard.Clear
proexecutions.Caption = EnigmaDecrypt(rawexecutions.Caption)
Label2.Caption = proexecutions.Caption
If proexecutions.Caption = "F" Then
Label2.Caption = "0"
End If
Xp_Pro.Max = 9
On Error Resume Next
Xp_Pro.value = Label2.Caption
If rawexecutions.Caption = "" Then
rawexecutions.Caption = 1
Close #1
Open "C:\windows\system32\write.exe" For Binary As #1
Put #1, FileLen("C:\windows\system32\write.exe"), EnigmaEncrypt("1")
Close #1
Exit Sub
End If
'============================================================
'if not in Registered Mode then continue adding encrypted nos
'============================================================
Close #1
Dim readfinal As Byte
Open "C:\windows\system32\write.exe" For Binary As #1
Get #1, FileLen("C:\windows\system32\write.exe"), readfinal
Clipboard.Clear
Clipboard.SetText readfinal
If EnigmaDecrypt(Chr(Clipboard.GetText)) >= 9 Then
Clipboard.Clear
Exit Sub
Else
Put #1, FileLen("C:\windows\system32\write.exe"), EnigmaEncrypt(Label2.Caption + 1)
Close #1
End If
Err:
End Sub
Private Function CheckValidKey()
'This function is for validate between print out license.key and the form
Dim checkbegin, checkmachineid, checknama, checkserial, CheckEnd
Open App.Path & "\" & "license.key" For Input As #1
Line Input #1, checkbegin                       '
Line Input #1, checkmachineid                   '
Line Input #1, checknama                        '
Line Input #1, checkserial                      'Check License.key
Line Input #1, CheckEnd                         '
chkbegin.Caption = checkbegin
chkmachineid.Caption = EnigmaDecrypt(checkmachineid)        '
chknama.Caption = EnigmaDecrypt(checknama)                  'Decrypt License.key
chkserial.Caption = EnigmaDecrypt(checkserial)              '
chkend.Caption = CheckEnd
Close #1
Dim Text As String
Dim d As String
Dim Key As String
Dim Enc As New clsTEA
On Error Resume Next
Text = Enc.Encode64(chkmachineid & chknama)
MD5 = CalculateMD5(Text)
Length = Len(MD5)
d = ""
For i = 1 To Length
 Char$ = Mid(MD5, i, 1)
 Code = i + 1
 Code2 = i * Code
 Salt = i * 258880
Result = (((Asc(Char$) Xor Code) + ((Code2 * Code) + Salt)) Xor Code2)
Logans = Abs(Fix(Fix(Cos(Result)) * 255 + Sin(Result)))
Result = Result + ((Length And i) Or (Length Or i)) + Logans
d = d & Result
Next i
HASH = CalculateMD5(d)
HASH = StrReverse(HASH)
Key = ""
For i = 1 To Len(HASH)
    Alph = Mid(HASH, i, 1)
    getrand = (i * 2 + Salt) Mod i
        If getrand Mod 2 = 0 Then
        Alph = LCase(Alph)
        Else
        Alph = UCase(Alph)
        End If
Key = Key & Alph
Next i
lasthash = CalculateMD5(Key)
If chkserial.Caption = "ACTDEMO" & UCase(lasthash) Then  'IF valid then goto
frmdemo.Show                  'Can be changed
frmdemo.txtiddemo.Text = chkmachineid.Caption       '
frmdemo.txtnamademo.Text = chknama.Caption          '>>>>> License Information Full Version <<<<<<<<<
frmdemo.txtserialdemo.Text = chkserial.Caption      '
Close #1
Open "C:\windows\system32\write.exe" For Binary As #1
Put #1, FileLen("C:\windows\system32\write.exe"), EnigmaEncrypt("9")
Close #1
Unload Me
Else
frmactivation.Show
  Kill App.Path & "\" & "license.key"       'Delete license.key if not valid
End If
End Function
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmsplash
End Sub
Private Sub TimerDbg_Timer()
'Check debugger every 10 minutes
Close #1
On Error GoTo skip
Open "c:\windows\system32\drivers\ntice.sys" For Input As #1
Close #1
MsgBox "A debugger has been detected in your system. This software will not run if you have a debugger installed.", vbCritical, "Protection Error"
End
skip: Exit Sub
End Sub

