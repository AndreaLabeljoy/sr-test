VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Football World Cup Score Board"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16470
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
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   710
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1098
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwGames 
      Height          =   7395
      Left            =   420
      TabIndex        =   6
      Top             =   2430
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   13044
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   2556103
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Game"
         Object.Width           =   21167
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Score"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Start"
         Object.Width           =   38100
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "End"
         Object.Width           =   38100
      EndProperty
   End
   Begin VB.CommandButton cmdEndGame 
      Caption         =   "End game"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11970
      TabIndex        =   4
      Top             =   1770
      Width           =   1815
   End
   Begin VB.CommandButton cmdUpdateScore 
      Caption         =   "Update score"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10020
      TabIndex        =   3
      Top             =   1770
      Width           =   1815
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New game"
      Height          =   495
      Left            =   8070
      TabIndex        =   2
      Top             =   1770
      Width           =   1815
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1665
      Left            =   0
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1098
      TabIndex        =   0
      Top             =   0
      Width           =   16470
      Begin VB.Line lneSep 
         BorderColor     =   &H00E0E0E0&
         X1              =   26
         X2              =   1082
         Y1              =   86
         Y2              =   86
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Football World Cup Score Board - Code test Andrea De Filippo"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00888888&
         Height          =   375
         Left            =   570
         TabIndex        =   1
         Top             =   870
         Width           =   7860
      End
      Begin VB.Image imgLogo 
         Height          =   900
         Left            =   330
         Picture         =   "frmMain.frx":335A7
         Top             =   30
         Width           =   6000
      End
   End
   Begin VB.Label lblNoGames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No games present. Click ""New Game"" to start"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00888888&
      Height          =   300
      Left            =   570
      TabIndex        =   5
      Top             =   1680
      UseMnemonic     =   0   'False
      Width           =   5145
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : frmMain
' Author : Andrea De Filippo
' Date   : 2023-04-27 13:00
' Purpose: SportRadarScoreBoardLib example.
'---------------------------------------------------------------------------------------

Option Explicit

' Score library instance
Private WithEvents m_oScoreLib As clsScoreBoard
Attribute m_oScoreLib.VB_VarHelpID = -1

' Message box title
Private Const MSGBOX_TITLE As String = "Football World Cup Score Board"

Private Sub cmdEndGame_Click()
' Date   : 2023-04-27 15:49
' Purpose: End an on-going game
Dim sErr As String
Dim oItm As ListItem
Dim oGame As clsGame
Dim sPrompt As String

    On Error GoTo ERR_HANDLE
    
    ' Get selected game
    Set oItm = lvwGames.SelectedItem

    ' Sanity check
    If oItm Is Nothing Then
        MsgBox "Select an on-going game.", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If
    
    ' Get game and double check its state
    Set oGame = m_oScoreLib.GetGame(oItm.Tag)
    If oGame.GameIsOnGoing = False Then
        MsgBox "Select an on-going game.", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If
    
    ' Confirm
    sPrompt = "End game " & oGame.HomeTeamName & " - " & oGame.AwayTeamName & "?"
    If MsgBox(sPrompt, vbQuestion Or vbOKCancel) <> vbOK Then Exit Sub

    ' End corresponding game
    m_oScoreLib.EndGame oItm.Tag

    ' Ok
    On Error GoTo 0
    Exit Sub

ERR_HANDLE:

    ' Errors
    MsgBox Err.Description, vbCritical, MSGBOX_TITLE

End Sub

Private Sub cmdNewGame_Click()
' Date   : 2023-04-27 15:16
' Purpose: Add a new game
Dim sErr As String
Dim f As frmNewGame

    On Error GoTo ERR_HANDLE

    ' Show form
    Set f = New frmNewGame
    f.Show vbModal, Me
    
    ' Check results
    If f.Success Then
        m_oScoreLib.StartNewGame f.txtHomeTeam.Text, f.txtAwayTeam.Text
    End If
    
    ' Clean up
    Unload f
    Set f = Nothing

    ' Ok
    On Error GoTo 0
    Exit Sub

ERR_HANDLE:

    If f Is Nothing = False Then
        Unload f
        Set f = Nothing
    End If

    ' Errors
    MsgBox Err.Description, vbCritical, MSGBOX_TITLE

End Sub

Private Sub cmdUpdateScore_Click()
' Date   : 2023-04-27 16:26
' Purpose: Update score for ther selected game
Dim sErr As String
Dim oItm As ListItem
Dim oGame As clsGame
Dim f As frmUpdateScore

    On Error GoTo ERR_HANDLE

    ' Get selected game
    Set oItm = lvwGames.SelectedItem

    ' Sanity check
    If oItm Is Nothing Then
        MsgBox "Select an on-going game.", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If
    
    ' Get game and double check its state
    Set oGame = m_oScoreLib.GetGame(oItm.Tag)
    If oGame.GameIsOnGoing = False Then
        MsgBox "Select an on-going game.", vbExclamation, MSGBOX_TITLE
        Exit Sub
    End If
    
    ' Init form
    Set f = New frmUpdateScore
    Set f.Game = oGame
    
    ' Show
    f.Show vbModal, Me
    
    ' Check result
    If f.Success Then
        ' Update scores
        m_oScoreLib.UpdateScore oGame.GameGuid, f.NewHomeScore, f.NewAwayScore
    End If
    
    ' Clean up
    Unload f
    Set f = Nothing

    ' Ok
    On Error GoTo 0
    Exit Sub

ERR_HANDLE:

    If f Is Nothing = False Then
        Unload f
        Set f = Nothing
    End If

    ' Errors
    MsgBox Err.Description, vbCritical, MSGBOX_TITLE

End Sub

Private Sub Form_Load()
' Date   : 2023-04-27 13:09
' Purpose: Load controls

    lvwGames.ColumnHeaders(1).Width = 200
    lvwGames.ColumnHeaders(2).Width = 60
    lvwGames.ColumnHeaders(3).Width = 60
    lvwGames.ColumnHeaders(4).Width = 60
    
    ' Init score library
    Set m_oScoreLib = New clsScoreBoard


End Sub

Private Sub Form_Resize()
' Date   : 2023-04-27 13:01
' Purpose: Move and resize controls on UI
Const MARGIN_SIZE As Long = 32
Dim lWdt As Long

    On Error GoTo ERR_HANDLE

    ' Resize separator line
    lWdt = picTop.ScaleWidth - (MARGIN_SIZE * 2)
    If lWdt < 0 Then lWdt = 0
    lneSep.X1 = MARGIN_SIZE
    lneSep.X2 = MARGIN_SIZE + lWdt
    
    ' Align listview to the left
    pResizeListView
    
    ' Align buttons top
    cmdEndGame.Top = lvwGames.Top
    cmdNewGame.Top = lvwGames.Top
    cmdUpdateScore.Top = lvwGames.Top

    ' Ok
    On Error GoTo 0
    Exit Sub

ERR_HANDLE:

    ' Disregard errors, non vital
    Err.Clear

End Sub

Private Sub pResizeListView()
Dim lWdt As Long

    ' Align listview to the left
    lWdt = 500
    If lvwGames.ListItems.Count = 0 Then
        lWdt = 0
        ' Show no games label.
        ' Need to shrink the listview for it to appear, Visible property is not enough (refresh bug?)
        lblNoGames.Visible = True
        lvwGames.Visible = False
    Else
        lblNoGames.Visible = False
        lvwGames.Visible = True
    End If
    lvwGames.Move 0, picTop.Height, lWdt, Me.ScaleHeight - picTop.Height

End Sub

Private Sub lvwGames_ItemClick(ByVal Item As ListItem)
' Date   : 2023-04-27 15:24
' Purpose: Update UI

    pActivateButtons

End Sub

Private Sub m_oScoreLib_GameChanged(EventType As SportRadarScoreBoardLib.GameEventType, Game As SportRadarScoreBoardLib.clsGame)
' Date   : 2023-04-27 13:37
' Purpose: Changes in games detected -> Update list and score board
' Note   : Point 4 of the test is always performed on the list whenever there's a change
Dim sCurrentSel As String
Dim oGame As clsGame
Dim oItm As ListItem

    On Error GoTo ERR_HANDLE
    
    ' Store currently selected items
    If EventType = gt_NewGameStarted Then
        ' Select the newly created item
        sCurrentSel = Game.GameGuid
    Else
        ' Re-select when possible
        If lvwGames.SelectedItem Is Nothing = False Then
            sCurrentSel = lvwGames.SelectedItem.Tag
        End If
    End If
    
    ' Update list
    lvwGames.ListItems.Clear
    For Each oGame In m_oScoreLib.GetSortedList
        Set oItm = lvwGames.ListItems.Add
        oItm.Text = oGame.HomeTeamName & " - " & oGame.AwayTeamName
        oItm.Tag = oGame.GameGuid
        oItm.SubItems(1) = oGame.HomeTeamScore & " - " & oGame.AwayTeamScore
        oItm.SubItems(2) = Format$(oGame.StartTime, "HH:nn")
        If oGame.GameIsOnGoing Then
            oItm.SubItems(3) = "Playing"
        Else
            oItm.SubItems(3) = Format$(oGame.EndTime, "HH:nn")
        End If
        If sCurrentSel = oGame.GameGuid Then
            ' Select old item
            lvwGames.SelectedItem = oItm
        End If
    Next oGame
    
    ' Update UI
    pResizeListView
    pActivateButtons
    If lvwGames.Visible Then lvwGames.SetFocus
    
    ' Ok
    On Error GoTo 0
    Exit Sub

ERR_HANDLE:

    ' Show error
    MsgBox Err.Description, vbCritical, MSGBOX_TITLE

End Sub

Private Sub pActivateButtons()
' Date   : 2023-04-27 15:21
' Purpose: Update state of buttons
Dim oGame As clsGame
Dim oItm As ListItem

    On Error GoTo ERR_HANDLE
    
    ' Get selection
    Set oItm = lvwGames.SelectedItem
    
    ' Sanity check
    If oItm Is Nothing Then
        cmdEndGame.Enabled = False
        cmdUpdateScore.Enabled = False
        Exit Sub
    End If
    
    ' Get game
    Set oGame = m_oScoreLib.GetGame(oItm.Tag)
    
    ' Set buttons based on game state
    cmdEndGame.Enabled = oGame.GameIsOnGoing
    cmdUpdateScore.Enabled = oGame.GameIsOnGoing

    ' Ok
    On Error GoTo 0
    Exit Sub

ERR_HANDLE:

    ' Clear errors - Non vital
    Err.Clear
    
End Sub
