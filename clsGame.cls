VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' File   : clsGame
' Author : Andrea De Filippo
' Date   : 2023-04-27 10:57
' Purpose: Holds information about a game
'---------------------------------------------------------------------------------------
Option Explicit

' Private variables
Private m_sGameGuid As String       ' Unique identifier of this game.
Private m_sHomeTeamName As String   ' Name of home team.
Private m_sAwayTeamName As String   ' Name of away team.
Private m_iHomeTeamScore As Integer ' Score of home team.
Private m_iAwayTeamScore As Integer ' Score of away team.
Private m_bGameIsOnGoing As Boolean ' Flag indicating whether the game is playing.
Private m_dGameStartTime As Date    ' Start date/time of game.
Private m_dGameEndTime As Date      ' End date/time of game.

Private Sub Class_Initialize()
' Date   : 2023-04-27 11:38
' Purpose: Init variable. Assumes that the game starts when the instance is created.

    ' Create a new GUID
    m_sGameGuid = GetNewGUID()
    
    ' Init start time
    m_dGameStartTime = Now()
    
    ' Flag as on-going
    m_bGameIsOnGoing = True
    
End Sub

Public Property Get GameGuid() As String
' Date   : 2023-04-27 11:20
' Purpose: Return the GUID identifying this game

    GameGuid = m_sGameGuid

End Property

Public Property Get HomeTeamName() As String
' Date   : 2023-04-27 11:22
' Purpose: Return the Home team name

    HomeTeamName = m_sHomeTeamName

End Property

Friend Property Let HomeTeamName(ByVal sVal As String)
' Date   : 2023-04-27 11:23
' Purpose: Store Home team name. For init only, not available outside library.
    
    ' Sanity check
    If Trim$(sVal) = vbNullString Then
        Err.Raise &HFFF, , "Team name cannot be empty."
    End If
    
    m_sHomeTeamName = sVal

End Property

Public Property Get AwayTeamName() As String
' Date   : 2023-04-27 11:25
' Purpose: Return the Away team name

    AwayTeamName = m_sAwayTeamName

End Property

Friend Property Let AwayTeamName(ByVal sVal As String)
' Date   : 2023-04-27 11:28
' Purpose: Store Away team name. For init only, not available outside library.
    
    ' Sanity check
    If Trim$(sVal) = vbNullString Then
        Err.Raise &HFFF, , "Team name cannot be empty."
    End If
    
    m_sAwayTeamName = sVal

End Property

Public Property Get HomeTeamScore() As Integer
' Date   : 2023-04-27 11:29
' Purpose: Return home team score

    HomeTeamScore = m_iHomeTeamScore
    
End Property

Friend Property Let HomeTeamScore(ByVal iVal As Integer)
' Date   : 2023-04-27 11:30
' Purpose: Update Home team score. For clsScoreBoard.UpdateScore only, not available outside library.
Dim sErr As String

    ' Sanity checks
    If CheckScore(iVal, m_iHomeTeamScore, sErr) = False Then
        Err.Raise &HFFF, , sErr
    End If

    m_iHomeTeamScore = iVal

End Property

Public Property Get AwayTeamScore() As Integer
' Date   : 2023-04-27 11:32
' Purpose: Return Away team score

    AwayTeamScore = m_iAwayTeamScore
    
End Property

Friend Property Let AwayTeamScore(ByVal iVal As Integer)
' Date   : 2023-04-27 11:34
' Purpose: Update Away team score. For clsScoreBoard.UpdateScore only, not available outside library.
Dim sErr As String

    ' Sanity checks
    If CheckScore(iVal, m_iAwayTeamScore, sErr) = False Then
        Err.Raise &HFFF, , sErr
    End If
    
    m_iAwayTeamScore = iVal

End Property

Public Property Get StartTime() As Date
' Date   : 2023-04-27 14:00
' Purpose: Return the start time of the game

    StartTime = m_dGameStartTime

End Property

Public Property Get EndTime() As Date
' Date   : 2023-04-27 14:01
' Purpose: Return the end time of the game

    EndTime = m_dGameEndTime

End Property

Public Property Get GameIsOnGoing() As Boolean
' Date   : 2023-04-27 11:34
' Purpose: Return the current state of the game

    GameIsOnGoing = m_bGameIsOnGoing
    
End Property

Friend Function EndGame()
' Date   : 2023-04-27 11:50
' Purpose: End the game. For clsScoreBoard.EndGame only, not available outside library.

    ' Sanity check
    If m_bGameIsOnGoing = False Then Err.Raise &HFFF, , "Game has already ended."

    ' Change game state
    m_bGameIsOnGoing = False
    
    ' Store end game date
    m_dGameEndTime = Now()
    

End Function

Friend Function CheckScore(iNewScore As Integer, iOldScore As Integer, ByRef sErr As String) As Boolean
' Date   : 2023-04-27 12:09
' Purpose: Check before updating a score

    ' Check that game is still on.
    If m_bGameIsOnGoing = False Then
        sErr = "Cannot change then score after a game as concluded."
        Exit Function
    End If
    
    ' Check for negative value
    If iNewScore < 0 Then
        sErr = "Score cannot be negative."
        Exit Function
    End If
    
    ' Ok
    CheckScore = True

End Function

Friend Function SortKey() As String
' Date   : 2023-04-27 14:05
' Purpose: Return a sort key for sorting by total score and start time
Dim sScore As String
Dim sStart As String

    ' Build fixed-length score key
    sScore = Format$(m_iHomeTeamScore + m_iAwayTeamScore, "000")
    sStart = Format$(m_dGameStartTime, "yyyymmddhhnnss")
    SortKey = sScore & sStart

End Function

