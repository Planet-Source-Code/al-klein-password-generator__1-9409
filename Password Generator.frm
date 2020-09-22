VERSION 5.00
Begin VB.Form Password 
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   Icon            =   "Password Generator.frx":0000
   LinkTopic       =   "Password Generator"
   ScaleHeight     =   1980
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSerial 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1020
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4380
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton btnCopy 
      Caption         =   "Copy code to Clipboard"
      Height          =   375
      Left            =   2460
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtLength 
      Height          =   285
      Left            =   1500
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton btnGetCode 
      Caption         =   "Get Code"
      Height          =   255
      Left            =   4380
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton btnCheckCode 
      Caption         =   "Check Code"
      Height          =   255
      Left            =   4380
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   1020
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1020
      TabIndex        =   4
      Top             =   780
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "Disk Serial#"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   420
      Width           =   855
   End
   Begin VB.Label lblResult 
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   5835
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Enter Code Length"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Enter Code"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter name"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   810
      Width           =   810
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Regcode Generator v2.0
'Author: Dustin DAvis
'Bootleg Software Inc.
'http://www.warpnet.org/bsi
'
'I created this code and put it up to help people create
'registration codes for their programs. It is very simple
'to use and to figure out!
'Do not steal this code!! Please give me proper credit for it
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Amritanshu Gupta (tanshu@i.am) -- I Changed the txtName_keypress event
'Removed the bug -- Pl. give me some credit too
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Improved the encoding algorithm - see the Readme file
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Dim regCode$, inCode$, Code%(25), serialnum&
Private Sub btnCheckCode_Click()
Dim temp As String
Dim length As Integer

  length = CInt(txtLength.Text)
  temp = Left$(regCode, length)
  If txtCode.Text = temp Then
    MsgBox "Pass"
  Else
    MsgBox "Failed"
  End If

End Sub

Private Sub btnGetCode_Click()
Dim length As Integer

  'Make sure we have something to encode first
  If txtSerial.Text = "" Then
    i% = MsgBox("You must enter a serial number before you can create a registration code", 16)
    txtSerial.SetFocus
    Exit Sub
  ElseIf txtName.Text = "" Then
    i% = MsgBox("You must enter a name before you can create a registration code", 16)
    txtName.SetFocus
    Exit Sub
  End If
  
  Call Setup ' Get the registration code
  length = CInt(txtLength.Text)
  lblResult.Caption = "Reg Code for: " & txtName.Text & " - " & Left$(regCode, length)

End Sub

Private Sub btnCopy_Click()
Dim temp As String
Dim length As Integer

  length = CInt(txtLength.Text)
  temp = Left$(regCode, length)
  Clipboard.SetText temp

End Sub

Private Sub btnClear_Click()

'Clear everything

  Clipboard.SetText ""
  regCode = ""
  txtName.Text = ""
  txtCode.Text = ""
  txtSerial.Text = ""
  lblResult.Caption = ""
  txtSerial.SetFocus

End Sub

Private Sub Form_GotFocus()
  
  txtSerial.SetFocus

End Sub

Private Sub Form_Load()
  
  serialnum = 0
  Call GetDiskSerial
End Sub

Private Sub Setup()
Dim e$, i%, j%
Const upperbound% = 255
Const lowerbound = 33
  
  serialnum = Val(txtSerial.Text)
  inCode = ""
  regCode = ""
    
  License = Int(Sqr(1.732 * serialnum))
  While License > 32767
    License = Int(License / 1.895)
  Wend
  i = CInt(License)
  i = Rnd(1 - i)
  For i = 0 To 25
    Code(i) = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
  Next i

  regCode = ""
  e$ = txtName.Text
  While Len(e$) < 10
    e$ = " " & e$
  Wend
  If Len(e$) > 25 Then
    j = 25
  Else
    j = Len(e$)
  End If
  txtLength.Text = Str(j)
  For i = 1 To j
      regCode = regCode & (Asc(Mid(e$, i, 1)) Xor Code(i))
  Next


End Sub

Private Sub txtLength_Change()
Dim temp As String
Dim length As Integer

If Trim(txtLength.Text) > "0" Then 'If it's "", the next line generates an error
                                   'you could also use error trapping here
    length = CInt(txtLength.Text)
    lblResult.Caption = "Reg Code for: " & txtName.Text & " - " & Left$(regCode, length)
Else
    DoEvents
End If
End Sub
Private Sub GetDiskSerial()
Dim volbuf$, sysname$, sysflags&, componentlength&, res&, License$, i%, j%
Const upperbound% = 255
Const lowerbound = 33

    volbuf$ = String$(256, 0)
    sysname$ = String$(256, 0)

    res = GetVolumeInformation("C:\", volbuf$, 255, serialnum, _
            componentlength, sysflags, sysname$, 255)

    License = Int(Sqr(1.895 * serialnum))
    While License > 32767
      License = Int(License / 1.732)
    Wend
'    License = Int(Sqr(serialnum))
    'License = Int(License)
    'AddText ("HD's serial number got: " & serialnum)
    i = CInt(License)
    i = Rnd(1 - i)
    For i = 0 To 25
      Code(i) = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    Next i

End Sub
