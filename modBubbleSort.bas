Attribute VB_Name = "modBubbleSort"
'---------------------------------------------------------------------------------------
' File   : modBubbleSort
' Author : Andrea De Filippo
' Date   : 2023-04-27 14:10
' Purpose: Bubble sorting for listing games in requested order
' Source : https://www.codestack.net/visual-basic/algorithms/data/sorting/ (modified)
'---------------------------------------------------------------------------------------
Option Explicit

Private Declare Function StrCmpLogicalW Lib "shlwapi" (ByVal s1 As String, ByVal s2 As String) As Integer

Public Function BubbleSort(colGames As Collection) As Collection
Dim i As Integer
Dim j As Integer
Dim tempVal As String
Dim asKeys() As String
Dim oGame As clsGame
Dim asSort() As String

    ' Always init output
    Set BubbleSort = New Collection
        
    ' Sanity checks
    If colGames Is Nothing Then Exit Function
    If colGames.Count = 0 Then
        Exit Function
    End If
    If colGames.Count = 1 Then
        BubbleSort.Add colGames(1)
        Exit Function
    End If

    ' Convert collection to array
    ReDim asKeys(colGames.Count - 1)
    For i = 1 To colGames.Count
        Set oGame = colGames(i)
        asKeys(i - 1) = oGame.SortKey & vbTab & oGame.GameGuid ' Sort by SortkKey first then append guid
    Next i
    
    ' Bubble sort
    For i = 0 To UBound(asKeys)
        For j = i To UBound(asKeys)
            If StrCmpLogicalW(asKeys(i), asKeys(j)) = -1 Then
                tempVal = asKeys(j)
                asKeys(j) = asKeys(i)
                asKeys(i) = tempVal
            End If
        Next j
    Next i
    
    ' Build sorted collection
    For i = 0 To UBound(asKeys)
        asSort = Split(asKeys(i), vbTab)
        Set oGame = colGames(asSort(1))
        BubbleSort.Add oGame, asSort(1)
    Next i
    
End Function
