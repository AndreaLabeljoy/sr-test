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
' File   : clsWorldCupScoreBoard
' Author : Andrea De Filippo
' Date   : 2023-04-27 10:47
' Purpose: Implementation of the Football World Cup Score Board as a simple library.
'---------------------------------------------------------------------------------------

Option Explicit

Public Function StartNewGame(sHomeTeamName As String, sAwayTeamName As String) As String
' Date   : 2023-04-27 10:49
' Purpose: Start a new game
' Return : The GUID of the newly started game



End Function

Public Function FinishGame(sGameGuid As String) As Boolean
' Date   : 2023-04-27 10:53
' Purpose: End a game
' Return : True on success



End Function

Public Function UpdateScore(sGameGuid As String, iHomeTeamScore As Integer, iAwayTeamScore As Integer) As Boolean
' Date   : 2023-04-27 10:56
' Purpose: Update the score of an on-going game
' Return : True on success


End Function

