VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ActiveStatesConstants
    asIsActive
    asUserCancelled
    asPasswordFound
    asMaxExceeded
End Enum

' Properties.
Private m_sCharSet      As String
Private m_sCurPass      As String
Private m_sTarget       As String
Private m_iMaxLength    As Integer
Private m_iMinLength    As Integer
Private m_asActiveState As ActiveStatesConstants

' Events.  The concept/implementaion of the events taken directly from the
' Planet-Source-Code post "Fast BruteForce Class Example" by §e7eN.
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=48276&lngWId=1
Public Event PasswordsPerSecond(Passwords As Long)
Public Event TotalPasswords(Passwords As Long)

' Module-level variables.
Private m_lPassPerSec As Long
Private m_lPassTime   As Long
Private m_lTotalPass  As Long
Private m_asCharSet() As String
Private m_bFound      As Boolean

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Property Let CharacterSet(nv As String)

Dim i As Integer

    m_sCharSet = nv
    
    ReDim m_asCharSet(Len(m_sCharSet) - 1)
    
    ' Create an array of characters based on the charset.
    For i = 0 To UBound(m_asCharSet)
        m_asCharSet(i) = Mid$(m_sCharSet, i + 1, 1)
    Next
        
End Property

Property Get CharacterSet() As String
    CharacterSet = m_sCharSet
End Property

Property Let CurrentPassword(nv As String)
    m_sCurPass = nv
End Property

Property Get CurrentPassword() As String
    CurrentPassword = m_sCurPass
End Property

Property Let Target(nv As String)
    m_sTarget = nv
End Property

Property Get Target() As String
    Target = m_sTarget
End Property

' TODO: Add property validation.
Property Let MaxLength(nv As Integer)
    m_iMaxLength = nv
End Property

Property Get MaxLength() As Integer
    MaxLength = m_iMaxLength
End Property

Property Let MinLength(nv As Integer)
    m_iMinLength = nv
End Property

Property Get MinLength() As Integer
    MinLength = m_iMinLength
End Property

Property Let ActiveState(nv As ActiveStatesConstants)
    m_asActiveState = nv
End Property

Property Get ActiveState() As ActiveStatesConstants
    ActiveState = m_asActiveState
End Property

Private Sub GetPasswordsStats()

    m_lPassPerSec = m_lPassPerSec + 1

    If GetTickCount - m_lPassTime >= 1000 Then
        RaiseEvent PasswordsPerSecond(m_lPassPerSec)
            
        m_lPassTime = GetTickCount
        m_lTotalPass = m_lTotalPass + m_lPassPerSec
        m_lPassPerSec = 0
        RaiseEvent TotalPasswords(m_lTotalPass)
    End If
    
End Sub

Public Sub Generate()

Dim iLen As Integer, bMaxExceeded As Boolean

    ' Reset variables/flags.
    iLen = m_iMinLength
    m_bFound = False
    m_asActiveState = asIsActive
    m_lTotalPass = 0
    
    ' Generate passwords while:
    ' 1. Password does not match target.
    ' 2. Password length does not exceed user-defined maximum.
    ' 3. User does not cancel operation.

    Do
        pMakePermsWithRepeats iLen
        iLen = iLen + 1
        If iLen > m_iMaxLength Then bMaxExceeded = True
    Loop Until m_bFound Or bMaxExceeded Or m_asActiveState <> asIsActive
    
    ' Determine why password generation stopped.
    If m_bFound Then
        m_asActiveState = asPasswordFound
    ElseIf bMaxExceeded Then
        m_asActiveState = asMaxExceeded
    End If
    
End Sub


Private Sub pMakePermsWithRepeats(ByVal iOutputLen As Integer)

' The main sub.

Dim iBase As Integer, i As Double
Dim dMaxPass As Double, sField As String

    ' Get number-system base.
    iBase = Len(m_sCharSet)
    
    ' Ex: base=3, passwords range from 000-222.
    dMaxPass = iBase ^ iOutputLen - 1
    
    sField = String$(iOutputLen, Left$(m_sCharSet, 1))

    For i = 0 To dMaxPass
    
        ' Convert the decimal number to a password.
        m_sCurPass = Right$(sField & DecToPass(i, iBase), iOutputLen)

        ' TODO: Un-hardcode refresh rate.
        If i Mod 5000 = 0 Then DoEvents

        If m_sCurPass = m_sTarget Then
            m_bFound = True
            Exit Sub
        ' This seems kludgy, but is necessary is user cancels.
        ElseIf m_asActiveState <> asIsActive Then
            Exit Sub
        End If

        GetPasswordsStats
    Next
    
End Sub

Private Function DecToPass(ByVal dDec As Double, _
                           ByVal iBase As Integer) As String

    ' d Mod Base will be in range [0 .. base-1]
    
Dim s As String

    ' This loop is like the basic DecToBin conversion, but it has been modified
    ' to be more generic DecToAnyBase.
    
    ' The actual conversion to a base n number is bypassed, instead the
    ' number in each place is used as an index in the charset array.
    
    ' Ex: dec=4,base=3, then s=011 and using those numbers as indexes
    ' w/charset="abc" you would return "abb".
    
    ' loop ? times: 4\3=1,1\3=0, int(log(dDec)/log(iBase))
    Do
        s = m_asCharSet(dDec Mod iBase) & s
        dDec = Int(dDec / iBase)
    Loop While dDec

    DecToPass = s
    
End Function

Private Sub Class_Initialize()
    m_asActiveState = asIsActive
End Sub



















'' LEFTOVER STUFF
' Passwords of k objects from set of n objects
'            n!
'n_C_k = ----------
'        k!(n - k)!

'OLD METHOD
'        NumInBase = Right$(sField & DecToBaseN2(i, iBase), iOutputLen)
'        For j = 1 To iOutputLen
'            letter = Asc(Mid$(NumInBase, j, 1))
'            If letter < 97 Then
'                 letter = letter - 48
'            Else
'                letter = letter - 87
'            End If
'            Mid$(sReturn, j, 1) = Mid$(m_sCharSet, letter + 1, 1)
'        Next

' OLD STUFF
'Private Sub MakeCombinations(ByVal sSet As String)
'
'Dim i As Long, j As Integer, sBin As String, s As String
'Dim iLen As Integer
'
'    iLen = Len(sSet)
'
'    For i = 1 To 2 ^ iLen - 1
'        sBin = DecToBinL(i, iLen)
'        For j = 1 To iLen
'            If Mid$(sBin, j, 1) = "1" Then s = s & Mid$(sSet, j, 1)
'        Next
'        List1.AddItem s
'        s = ""
'    Next
'
'End Sub

'Private Function DecToBinL(ByVal d As Long, ByVal w As Integer) As String

'Dim b As String
'
'    Do
'        b = (CInt((d / 2) = Int(d / 2)) + 1) & b
'        d = Int(d * 0.5)
'    Loop While d
'
'    DecToBinL = Right$(String(w, "0") & b, w)
'
'End Function

'Private Function DecToBase(ByVal d As Long, ByVal b As Integer, w As Integer) As String
'
'Dim s As String
'
'    Do
'        s = m_sDigits(d Mod b) & s
'        d = Int(d / b)
'    Loop While d
'
'    DecToBase = Right$(String(w, "0") & s, w)
'
'End Function
