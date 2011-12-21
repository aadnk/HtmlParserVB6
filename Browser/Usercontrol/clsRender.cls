VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

' Used to draw text
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, _
 ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

' The current drawing position
Public CurrentX As Long
Public CurrentY As Long

' The control and the object to draw to
Public Control As usrBrowser
Public Destination As PictureBox

Public Sub DrawDocument(Document As clsDocument)

    Dim HTML As Object, Body As Object, Font As StdFont, List As clsNodeList, bFound As Boolean
    
    ' First of all, get the HTML-tag
    For Each HTML In Document.DocumentElement.ChildNodes
        If LCase(HTML.NodeName) = "html" And HTML.NodeType = DOMElement Then
            
            ' Yup, we've found a node
            bFound = True
            
            ' We're finished
            Exit For
            
        End If
    Next
    
    ' If we couln't locate the html element, just assume that it exists
    If Not bFound Then
        ' We'll use the document element
        Set HTML = Document.DocumentElement
    End If
    
    ' Then get the body element
    Set List = HTML.GetElementsByTagName("body")
    
    ' If this dosen't exists, we'll assume that it does
    If List.Lenght < 0 Then
        ' We'll think of it as the html-element
        Set Body = HTML
    Else
        ' If not, use the first matching element
        Set Body = List.Item(0)
    End If
    
    ' Clear the position variables
    CurrentX = 0
    CurrentY = 0
    
    ' Clear the destination to paint to
    Destination.Cls
    
    ' Set the default font-settings
    Set Font = CloneFont(Control.DefaultFont)
    
    ' Start the drawing
    DrawElement Body, Font

End Sub

Public Sub DrawElement(Element As clsElement, Font As StdFont)

    Dim sText As String, sChar As String, Tell As Long, Child As Object, sLastChar As String

    ' Use this font settings
    Control.UseFont Font

    ' Firstly, different things is to be done depending on this element's type
    Select Case Element.NodeType
    Case DOMText
    
        ' Firstly, get the text of this node
        sText = Element.NodeValue
    
        ' Clear variable
        sLastChar = ""
    
        ' We need to draw the text, character for character
        For Tell = 1 To Len(sText)
        
            ' Get the character to draw
            sChar = Mid(sText, Tell, 1)
        
            ' Ignore certain characters
            If sChar <> Chr(10) And sChar <> Chr(13) And sChar <> Chr(9) And Not (sLastChar = sChar And sChar = " ") Then
        
                ' Draw the character
                TextOut Destination.hdc, CurrentX, CurrentY, sChar, 1
                
                ' Calculate the new X-position
                CurrentX = CurrentX + Destination.TextWidth(sChar)
        
                ' If we have gone to far, ...
                If CurrentX >= Destination.ScaleWidth Then
                
                    ' ... end this line
                    NewLine 1
                    
                End If
                
            End If
            
            ' Save this character
            sLastChar = sChar
        
        Next
    
    Case DOMElement
    
        ' Do different things based on the name of the node
        Select Case Element.NodeName
        Case "br" ' Forced line break
        
            ' Jump one line
            NewLine 1
        
        Case "b", "strong"
        
            ' Make the font bold
            Font.Bold = True
            
        Case "i", "em"
        
            ' Make the font italic
            Font.Italic = True
        
        End Select
        
        ' Draw all it's subelements
        For Each Child In Element.ChildNodes
        
            ' Draw the element
            DrawElement Child, CloneFont(Font)
            
            ' Still use this font setting
            Control.UseFont Font
        
        Next
        
        ' Actions to be done after processing the element's children
        Select Case Element.NodeName
        Case "p" ' A paragraph, jump two lines
        
            NewLine 2
            
        End Select

    End Select

End Sub

Private Function NewLine(Optional Amout As Long = 1)

    ' Simply reposition the writing "head"
    CurrentY = CurrentY + Destination.TextHeight("A") * Amout
    CurrentX = 0

End Function

Private Function CloneFont(refFont As StdFont) As StdFont

    ' Firstly, create the new class instance
    Set CloneFont = New StdFont
    
    ' Then set all its settings
    With CloneFont
        .Bold = refFont.Bold
        .Charset = refFont.Charset
        .Italic = refFont.Italic
        .Name = refFont.Name
        .Size = refFont.Size
        .Strikethrough = refFont.Strikethrough
        .Underline = refFont.Underline
        .Weight = refFont.Weight
    End With

End Function
