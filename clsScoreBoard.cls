VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScoreBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' File   : clsScoreBoard
' Author : Andrea De Filippo
' Date   : 2023-04-27 10:47
' Purpose: Implementation of the Football World Cup Score Board as a simple library.
'---------------------------------------------------------------------------------------
Option Explicit

' Collection of games
Private m_colGames As New Collection

' Event type
Public Enum GameEventType
    gt_NewGameStarted
    gt_GameEnded
    gt_ScoreChanged
End Enum

' public events
Public Event GameChanged(EventType As GameEventType, Game As clsGame)

Public Function GetGame(sGameGuid As String) As clsGame
' Date   : 2023-04-27 15:57
' Purpose: Return a specific game in the collection

    ' Get the game
    Set GetGame = m_colGames(sGameGuid) ' Will raise an error if the game is not in the collection

End Function

Public Function StartNewGame(ByVal sHomeTeamName As String, ByVal sAwayTeamName As String) As String
' Date   : 2023-04-27 10:49
' Purpose: Start a new game
' Return : The GUID of the newly started game
Dim oGame As clsGame

    ' Pre-process team names
    sHomeTeamName = Trim$(sHomeTeamName)
    sAwayTeamName = Trim$(sAwayTeamName)

    ' Check for an on-going game with one of the teams being passed.
    For Each oGame In m_colGames
        If oGame.GameIsOnGoing Then
            If oGame.HomeTeamName = sHomeTeamName Or _
               oGame.AwayTeamName = sHomeTeamName Then
                Err.Raise &HFFF, , "Home team is already playing."
            End If
            If oGame.HomeTeamName = sAwayTeamName Or _
               oGame.AwayTeamName = sAwayTeamName Then
                Err.Raise &HFFF, , "Away team is already playing."
            End If
        End If
    Next oGame
    
    ' Check for equal team names
    If sHomeTeamName = sAwayTeamName Then
        Err.Raise &HFFF, , "Home team name and Away team name cannot be the same."
    End If

    ' Create a new game and start it.
    Set oGame = New clsGame
    
    ' Init team names
    oGame.HomeTeamName = sHomeTeamName
    oGame.AwayTeamName = sAwayTeamName
    
    ' No errors with team names, we can add the new game to the collection.
    ' Use Game guid as collection key
    m_colGames.Add oGame, oGame.GameGuid
    
    ' Return the new game guid
    StartNewGame = oGame.GameGuid
    
    ' Raise event
    RaiseEvent GameChanged(gt_NewGameStarted, oGame)

End Function

Public Function EndGame(sGameGuid As String) As Boolean
' Date   : 2023-04-27 10:53
' Purpose: End a game
' Return : True on success
Dim oGame As clsGame

    ' Get the game
    Set oGame = m_colGames(sGameGuid) ' Will raise an error if the game is not in the collection
    
    ' End game
    oGame.EndGame
    
    ' Raise event
    RaiseEvent GameChanged(gt_GameEnded, oGame)
    
    ' Signal success
    EndGame = True

End Function

Public Function UpdateScore(sGameGuid As String, iHomeTeamScore As Integer, iAwayTeamScore As Integer) As Boolean
' Date   : 2023-04-27 10:56
' Purpose: Update the score of an on-going game
' Return : True on success
Dim oGame As clsGame
Dim sErr As String

    ' Get the game
    Set oGame = m_colGames(sGameGuid) ' Will raise an error if the game is not in the collection
    
    ' Skip if the new scores are the same as current ones
    If oGame.HomeTeamScore = iHomeTeamScore And _
       oGame.AwayTeamScore = iAwayTeamScore Then
        Exit Function
    End If
    
    ' Check scores
    If oGame.CheckScore(iHomeTeamScore, oGame.HomeTeamScore, sErr) = False Then
        Err.Raise &HFFF, , sErr
    End If
    If oGame.CheckScore(iAwayTeamScore, oGame.AwayTeamScore, sErr) = False Then
        Err.Raise &HFFF, , sErr
    End If

    ' It's ok to update scores
    oGame.HomeTeamScore = iHomeTeamScore
    oGame.AwayTeamScore = iAwayTeamScore
        
    ' Raise event
    RaiseEvent GameChanged(gt_ScoreChanged, oGame)
    
    ' Signal success
    UpdateScore = True

End Function

Public Function GetSortedList() As Collection
' Date   : 2023-04-27 13:50
' Purpose: Returns the list of games sorted by total score and start time

    ' Return
    Set GetSortedList = BubbleSort(m_colGames)

End Function
