VERSION 5.00
Begin VB.UserControl usrBrowser 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9435
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   629
   Begin VB.PictureBox picBrowser 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   0
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "usrBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

' The default font to use
Public DefaultFont As StdFont

' Allocate the rendering engine
Public RenderEngine As New clsRender

' The document to draw
Public Document As clsDocument

' The main parser
Private XMLParser As clsParser

Public Sub Navigate(URL As String)

    ' Firstly, download and get the documennt
    Set Document = XMLParser.ParseURI(URL)

    ' Reference this control and its drawing-object
    Set RenderEngine.Control = Me
    Set RenderEngine.Destination = picBrowser

    ' Then render the document
    RenderEngine.DrawDocument Document

End Sub

Public Sub UseFont(Font As StdFont)

    ' Set all font settings
    Set picBrowser.Font = Font

End Sub

Private Sub picBrowser_Resize()

    ' Resize the picturebox
    picBrowser.Width = UserControl.ScaleWidth
    picBrowser.Height = UserControl.ScaleHeight

End Sub

Private Sub UserControl_Initialize()

    ' Create the parser
    Set XMLParser = CreateLSParser(0, "")
    
    ' Set the font
    Set DefaultFont = picBrowser.Font

End Sub
