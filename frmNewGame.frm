VERSION 5.00
Begin VB.Form frmNewGame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New game"
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
   Icon            =   "frmNewGame.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEsc 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   4710
      TabIndex        =   5
      Top             =   1590
      Width           =   1395
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
   Begin VB.TextBox txtAwayTeam 
      Height          =   405
      Left            =   1650
      TabIndex        =   3
      Top             =   870
      Width           =   4455
   End
   Begin VB.TextBox txtHomeTeam 
      Height          =   405
      Left            =   1650
      TabIndex        =   1
      Top             =   330
      Width           =   4455
   End
   Begin VB.Label lblGuestTeam 
      AutoSize        =   -1  'True
      Caption         =   "Guest team:"
      Height          =   225
      Left            =   585
      TabIndex        =   2
      Top             =   930
      Width           =   945
   End
   Begin VB.Label lblHomeTeam 
      AutoSize        =   -1  'True
      Caption         =   "Home team:"
      Height          =   225
      Left            =   540
      TabIndex        =   0
      Top             =   390
      Width           =   990
   End
End
Attribute VB_Name = "frmNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : frmNewGame
' Author : Andrea De Filippo
' Date   : 2023-04-27 15:10
' Purpose: Add and start a new game
'---------------------------------------------------------------------------------------

' Private variables
Private m_bSuccess As Boolean

' Message box title
Private Const MSGBOX_TITLE As String = "Football World Cup Score Board"

Option Explicit

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
' Date   : 2023-04-27 15:13
' Purpose: Confirm and exit
Dim sErr As String

    On Error GoTo ERR_HANDLE

    ' Sanity checks
    If Trim$(txtHomeTeam.Text) = vbNullString Then
        MsgBox "Home team name is not valid.", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If
    If Trim$(txtAwayTeam.Text) = vbNullString Then
        MsgBox "Guest team name is not valid.", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If
    If Trim$(txtAwayTeam.Text) = Trim$(txtHomeTeam.Text) Then
        MsgBox "Home team name and Guest team name cannot be the same.", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If

    ' Ok
    m_bSuccess = True
    Me.Hide
    On Error GoTo 0
    Exit Sub

ERR_HANDLE:

    ' Errors
    MsgBox Err.Description, vbCritical, MSGBOX_TITLE

End Sub

Private Sub txtHomeTeam_GotFocus()
    ' Select all
    txtHomeTeam.SelStart = 0
    txtHomeTeam.SelLength = Len(txtHomeTeam.Text)
End Sub

Private Sub txtAwayTeam_GotFocus()
    ' Select all
    txtAwayTeam.SelStart = 0
    txtAwayTeam.SelLength = Len(txtAwayTeam.Text)
End Sub

