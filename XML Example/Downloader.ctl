VERSION 5.00
Begin VB.UserControl Downloader 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Complete(ByRef URL As String, ByRef Data As String, ByRef Key As String)
Public Event Progress(ByRef URL As String, ByRef Key As String, ByVal BytesDone As Long, ByVal BytesTotal As Long, ByVal Status As AsyncStatusCodeConstants)

Public Enum DownloaderCache
    [Always download] = vbAsyncReadForceUpdate
    [Get cache copy only] = vbAsyncReadOfflineOperation
    [Update cached copy only] = vbAsyncReadResynchronize
    [Use cache if no connection] = vbAsyncReadGetFromCacheIfNetFail
End Enum

Private m_Keys As String

Private Function Private_AddKey(ByRef Key As String) As Boolean
    ' see if we do not have the key
    Private_AddKey = InStr(m_Keys, vbNullChar & Key & vbNullChar) = 0
    ' we can add it
    If Private_AddKey Then
        m_Keys = m_Keys & Key & vbNullChar
    End If
End Function

Private Sub Private_RemoveKey(ByRef Key As String)
    ' remove the key
    m_Keys = Replace(m_Keys, vbNullChar & Key & vbNullChar, vbNullChar)
End Sub

Public Sub Start(ByRef URL As String, Optional ByVal CacheMode As DownloaderCache = [Always download], Optional ByVal Key As String)
    ' use URL as key if no key is given
    If LenB(Key) = 0 Then Key = URL
    ' do we already have this key?
    If Not Private_AddKey(Key) Then
        ' cancel the old one
        CancelAsyncRead Key
    End If
    ' begin download process
    AsyncRead URL, vbAsyncTypeByteArray, Key, CacheMode
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    Dim strData As String
    ' get Variant byte array to byte string (needs StrConv to Unicode for displaying in a textbox)
    If AsyncProp.BytesRead Then strData = AsyncProp.Value Else strData = vbNullString
    ' redirect information
    RaiseEvent Complete(AsyncProp.Target, strData, AsyncProp.PropertyName)
    ' remove the key
    Private_RemoveKey AsyncProp.PropertyName
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    With AsyncProp
        ' redirect event information
        If LenB(.PropertyName) Then
            RaiseEvent Progress(.Target, .PropertyName, .BytesRead, .BytesMax, .StatusCode)
        Else
            RaiseEvent Progress(.Target, vbNullString, .BytesRead, .BytesMax, .StatusCode)
        End If
    End With
End Sub

Private Sub UserControl_Initialize()
    m_Keys = vbNullChar
End Sub
