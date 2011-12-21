Attribute VB_Name = "modDeclarations"
Option Explicit

' HtmlParserVB6 - A XML/HTML DOM-parser for VB6
' Copyright (C) 2011  Kristian. S Stangeland
'
' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
'
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA

' API-calls used primarily to access/manipulate strings faster
Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal lpString As Long, ByVal lLen As Long) As String
Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long
Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (pDst As Any, ByVal ByteLen As Long)
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' A HTML-property
Public Type HTMLProperty
    Name As String
    Value As String
    ID As Boolean
End Type

' The data type holding enities
Public Type Entity
    lngName() As Integer
    lngText() As Integer
End Type

' User data
Public Type UserData
    Key As String
    Data As Object
    Handler As Object
End Type

' Used to hold all HTML-tags
Public Type HTMLElement
    TagName As String
    Properties() As HTMLProperty
    PropertyCount As Long
    Parent As Long
    Children() As Long
    ChildCount As Long
    NodeType As NodeType
    Position As Long
    Value As String
    EndTag As Boolean
    Reject As Boolean
    UserData() As UserData
    UserDataCount As Long
End Type

' Used in conjunction with the below
Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

' The safe-array structure used in the fast string access
Public Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0) As SAFEARRAYBOUND
End Type

Public Const LeftSquareBracket As Integer = 91
Public Const RightSquareBracket As Integer = 93
Public Const QuotationMark As Integer = 34
Public Const Apostrophe As Integer = 39
Public Const LessThan As Integer = 60
Public Const GreaterThan As Integer = 62
Public Const SpaceChar As Integer = 32
Public Const Slash As Integer = 47
Public Const MinusSign As Integer = 45
Public Const EqualSign As Integer = 61
Public Const QuestionMark As Integer = 63
Public Const ExclamationMark As Integer = 33
Public Const Ampersand As Integer = 38
Public Const NumberSign As Integer = 35
Public Const Semicolon As Integer = 59
Public Const SmallX As Integer = 120
Public Const SmallA As Integer = 97
Public Const SmallF As Integer = 102
Public Const SmallZ As Integer = 122
Public Const LargeA As Integer = 65
Public Const LargeF As Integer = 70
Public Const LargeZ As Integer = 90
Public Const IntZero As Integer = 48
