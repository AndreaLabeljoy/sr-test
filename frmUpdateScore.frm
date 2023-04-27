VERSION 5.00
Begin VB.Form frmUpdateScore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update score"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmUpdateScore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtHomeTeam 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   1650
      MaxLength       =   2
      TabIndex        =   1
      Top             =   330
      Width           =   900
   End
   Begin VB.TextBox txtAwayTeam 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   1650
      MaxLength       =   2
      TabIndex        =   3
      Top             =   870
      Width           =   900
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   3210
      TabIndex        =   4
      Top             =   1590
      Width           =   1395
   End
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   4710
      TabIndex        =   5
      Top             =   1590
      Width           =   1395
   End
   Begin VB.Label lblHomeTeam 
      Alignment       =   1  'Right Justify
      Caption         =   "Home team:"
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   1470
   End
   Begin VB.Label lblGuestTeam 
      Alignment       =   1  'Right Justify
      Caption         =   "Guest team:"
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   930
      Width           =   1470
   End
End
Attribute VB_Name = "frmUpdateScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : frmUpdateScore
' Author : Andrea De Filippo
' Date   : 2023-04-27 16:08
' Purpose: Score setter form
'---------------------------------------------------------------------------------------
Option Explicit

' Private variables
Private m_bSuccess As Boolean
Private m_oGame As clsGame
Private m_iNewHome As Integer
Private m_iNewAway As Integer

' Message box title
Private Const MSGBOX_TITLE As String = "Football World Cup Score Board"

Public Property Set Game(oVal As clsGame)
' Date   : 2023-04-27 16:11
' Purpose: Init reference game

    Set m_oGame = oVal

End Property

Public Property Get NewHomeScore() As Integer
' Date   : 2023-04-27 16:15
' Purpose: Return validated new home score
    
    NewHomeScore = m_iNewHome

End Property

Public Property Get NewAwayScore() As Integer
' Date   : 2023-04-27 16:16
' Purpose: Return validated new away score
    
    NewAwayScore = m_iNewAway

End Property

Private Sub cmdEsc_Click()
    ' Abort
    Me.Hide
    
End Sub

Public Property Get Success() As Boolean
' Date   : 2023-04-27 15:18
' Purpose: Return Success flag

    Success = m_bSuccess
    
End Property

Private Sub cmdOk_Click()
' Date   : 2023-04-27 16:13
' Purpose: Validate scores and hide
Dim sErr As String
Dim iHome As Integer
Dim iAway As Integer

    On Error GoTo ERR_HANDLE
    
    ' Sanity checks
    If Trim$(txtHomeTeam.Text) = vbNullString Then
        txtHomeTeam.SetFocus
        MsgBox "Enter a valid score for the home team (" & m_oGame.HomeTeamName & ").", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If
    If Trim$(txtAwayTeam.Text) = vbNullString Then
        txtAwayTeam.SetFocus
        MsgBox "Enter a valid score for the guest team (" & m_oGame.AwayTeamName & ").", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If
    If (pCheckNumericString(txtHomeTeam.Text) = False) Or _
       (pCheckNumericString(txtAwayTeam.Text) = False) Then
        MsgBox "Only numbers can be entered.", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If

    ' Get scores
    iHome = Val(txtHomeTeam.Text)
    iAway = Val(txtAwayTeam.Text)
    
    ' Check scores
    If iHome < 0 Or iAway < 0 Then
        MsgBox "Negative numbers are not allowed.", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If
    If iHome < m_oGame.HomeTeamScore Then
        txtHomeTeam.SetFocus
        MsgBox "The home team score (" & m_oGame.HomeTeamName & ") must be equal or greater than the current score (" & m_oGame.HomeTeamScore & ").", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If
    If iAway < m_oGame.AwayTeamScore Then
        txtAwayTeam.SetFocus
        MsgBox "The guaest team score (" & m_oGame.AwayTeamName & ") must be equal or greater than the current score (" & m_oGame.AwayTeamScore & ").", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If

    ' Scores validates
    m_iNewHome = iHome
    m_iNewAway = iAway

    ' Ok
    m_bSuccess = (m_oGame.HomeTeamScore <> m_iNewHome) Or (m_oGame.AwayTeamScore <> m_iNewAway)
    Me.Hide
    On Error GoTo 0
    Exit Sub

ERR_HANDLE:

    ' Errors
    MsgBox Err.Description, vbCritical, MSGBOX_TITLE

End Sub

Private Function pCheckNumericString(sText As String) As Boolean
' Date   : 2023-04-27 16:18
' Purpose: Check that a string only contains numbers
Dim lChr As Long
Dim sChr As String
Dim iAscii As Integer

    On Error GoTo ERR_HANDLE

    ' Sanity check
    If sText = vbNullString Then Exit Function
    
    ' Loop characters
    For lChr = 1 To Len(sText)
        sChr = Mid$(sText, lChr, 1)
        iAscii = Asc(sChr)
        Select Case iAscii
            Case &H30 To &H39
                ' Ok
            Case Else
                ' Not ok
                Exit Function
        End Select
    Next lChr
    
    ' Ok
    pCheckNumericString = True
    On Error GoTo 0
    Exit Function

ERR_HANDLE:

    ' Clear errors
    Err.Clear
End Function

Private Sub Form_Load()
' Date   : 2023-04-27 16:12
' Purpose: Load controls

    txtHomeTeam.Text = m_oGame.HomeTeamScore
    txtAwayTeam.Text = m_oGame.AwayTeamScore
    lblHomeTeam.Caption = m_oGame.HomeTeamName & ":"
    lblGuestTeam.Caption = m_oGame.AwayTeamName & ":"

End Sub
