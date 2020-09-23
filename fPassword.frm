VERSION 5.00
Begin VB.Form fPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Generator"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRunTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   1560
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   3600
      TabIndex        =   13
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtRandLen 
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Text            =   "5"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdRand 
         Caption         =   "Charset"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   24
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdRand 
         Caption         =   "Target"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtMaxLen 
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Text            =   "8"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtMinLen 
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Text            =   "1"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblCap 
         Caption         =   "of length"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblCap 
         Caption         =   "Generate Random:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblCap 
         Caption         =   "Max Length"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCap 
         Caption         =   "Min Length"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraStats 
      Caption         =   "Statistics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
      Begin VB.TextBox txtReason 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   28
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtCurPass 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtPassLen 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtRunTime 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtTotalPass 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtPassPerSec 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblCap 
         Caption         =   "Active State"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblCap 
         Caption         =   "Current Password"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCap 
         Caption         =   "Password Length"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCap 
         Caption         =   "Running Time"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblCap 
         Caption         =   "Total Passwords"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblCap 
         Caption         =   "Passwords / Sec"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Password Generation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   27
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtCharSet 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "abcdefghijklmnopqrstuvwxyz"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtTarget 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Text            =   "target"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblCap 
         Caption         =   "Character Set"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblCap 
         Caption         =   "Target"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "fPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' project: pPassword
' date   : 2003-26-09
' author : redbird77
' email  : redbird77@earthlink.net
' www    : http://home.earthlink.net/~redbird77

' about  : See related document - Readme.txt.

Option Explicit

Private m_lRunTime         As Long
Private m_sActiveStates(3) As String
Private WithEvents m_oPass As cPassword
Attribute m_oPass.VB_VarHelpID = -1

Private Sub cmdRand_Click(Index As Integer)

Dim i As Integer, s As String

    Dim iLen As Integer
    
    iLen = Val(txtRandLen.Text)
    
    If iLen <= 0 Then
        MsgBox "Length of Charset must be greater than 0", vbExclamation
        Exit Sub
    End If

    ' Generate a random pass/charset between Chr(33) & Chr(126).
    ' It's not necessary to stay in those boundries though.
    For i = 1 To iLen
        s = s & Chr$(Int(94 * Rnd + 33))
    Next
    
    If Index Then
        txtCharSet.Text = s
    Else
        txtTarget.Text = s
    End If
    
End Sub

Private Sub cmdStart_Click()

    With m_oPass
    
        .CharacterSet = txtCharSet.Text
        .Target = txtTarget.Text
        
        ' TODO: Add property validation.
        .MaxLength = Val(txtMaxLen.Text)
        .MinLength = Val(txtMinLen.Text)
        
        m_lRunTime = 0
        tmrRunTime.Enabled = Not tmrRunTime.Enabled
        cmdStart.Enabled = False
        cmdStop.Enabled = True
        txtReason.Text = ""
        
        Do
            .Generate
        Loop Until .ActiveState <> asIsActive
        
        tmrRunTime.Enabled = False
        txtReason.Text = m_sActiveStates(.ActiveState)
        cmdStart.Enabled = True
        cmdStop.Enabled = False
        txtCurPass.Text = .CurrentPassword

    End With

End Sub

Private Sub cmdStop_Click()
    cmdStart.Enabled = True
    cmdStop.Enabled = False
    m_oPass.ActiveState = asUserCancelled
End Sub

Private Sub Form_Load()

    Randomize
    m_sActiveStates(0) = "Is Active": m_sActiveStates(1) = "User Cancelled"
    m_sActiveStates(2) = "Password Found": m_sActiveStates(3) = "Max Exceeded"
    
    Set m_oPass = New cPassword
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' TODO: Fix this!
    If m_oPass.ActiveState = asIsActive Then
        MsgBox "Stop password generation before you exit.", vbExclamation
        Cancel = 1
        Exit Sub
    End If
    
    Set m_oPass = Nothing
    Set fPassword = Nothing
End Sub


Private Sub m_oPass_PasswordsPerSecond(Passwords As Long)

    ' IDEA: move the counting aspect out of class and into form to speed up?
    txtCurPass.Text = m_oPass.CurrentPassword
    txtPassLen.Text = Len(txtCurPass.Text)
    txtPassPerSec.Text = CStr(Passwords)
    
End Sub

Private Sub m_oPass_TotalPasswords(Passwords As Long)

    txtTotalPass.Text = CStr(Passwords)
    
End Sub

Private Sub tmrRunTime_Timer()

    m_lRunTime = m_lRunTime + 1
    txtRunTime.Text = SecToTime(m_lRunTime)
    
End Sub

Private Function SecToTime(lRunTime As Long) As String
'This function modified from the Planet-Source-Code post "Fast BruteForce Class Example" by §e7eN.

Dim lHr As Long, lMin As Long, lSec As Long
    
    lSec = lRunTime Mod 60
    lMin = Int(lRunTime / 60)
    lHr = Int(lMin / 60)

    SecToTime = Format$(lHr, "00") & ":" & Format$(lMin, "00") & ":" & Format$(lSec, "00")
    
End Function
