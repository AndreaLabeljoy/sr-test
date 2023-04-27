Attribute VB_Name = "modGuid"
'---------------------------------------------------------------------------------------
' File   : modGuid
' Author : Andrea De Filippo
' Date   : 2023-04-27 11:17
' Purpose: Random GUID creation
' Source : https://www.vbforums.com/showthread.php?592964-VB6-Random-GUID-Generator
'---------------------------------------------------------------------------------------


Option Explicit

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Sub CoCreateGuid Lib "ole32.dll" (ByRef pguid As Guid)
Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long

Public Function GetNewGUID() As String
    Dim MyGUID As Guid
    Dim GUIDByte() As Byte
    Dim GuidLen As Long
    
    CoCreateGuid MyGUID
    
    ReDim GUIDByte(80)
    GuidLen = StringFromGUID2(VarPtr(MyGUID.Data1), VarPtr(GUIDByte(0)), UBound(GUIDByte))
    
    GetNewGUID = Left(GUIDByte, GuidLen - 1)
    GetNewGUID = Replace$(GetNewGUID, "{", vbNullString)
    GetNewGUID = Replace$(GetNewGUID, "}", vbNullString)
    
End Function
