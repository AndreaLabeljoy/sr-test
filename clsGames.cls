VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGames"
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
        
    ' Update sorting
    pUpdateSorting

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
    
    ' Update sorting
    pUpdateSorting
        
    ' Raise event
    RaiseEvent GameChanged(gt_ScoreChanged, oGame)
    
    ' Signal success
    UpdateScore = True

End Function

Private Sub pUpdateSorting()
' Date   : 2023-04-27 13:50
' Purpose: Bubble sort the collection
Dim oNewGames As Collection

    ' Return
    Set oNewGames = BubbleSort(Me)
    pCleanupCollection
    Set m_colGames = oNewGames

End Sub

Private Sub pCleanupCollection()
' Date   : 2023-04-28 11:32
' Purpose: Clean up
Dim lItm As Long

    On Error GoTo ERR_HANDLE

    For lItm = m_colGames.Count To 1 Step -1
        m_colGames.Remove lItm
    Next lItm
    Set m_colGames = Nothing

    ' Ok
    On Error GoTo 0
    Exit Sub

ERR_HANDLE:

    ' Clear errors - Non-vital
    Err.Clear

End Sub

Private Sub Class_Terminate()
' Date   : 2023-04-28 11:32
' Purpose: Clean up

    pCleanupCollection
    
End Sub

' Make class a collection class
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colGames.[_NewEnum]
End Property

Public Property Get Item(Key) As clsGame
Attribute Item.VB_UserMemId = 0
    Set Item = m_colGames(Key)
End Property

Public Property Get Count() As Long
    Count = m_colGames.Count
End Property

Public Sub Remove(Key)
    m_colGames.Remove Key
End Sub

