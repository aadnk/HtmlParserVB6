VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "HTML-Parser"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   452
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picToolBox 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   6495
      TabIndex        =   1
      Top             =   3840
      Width           =   6495
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdParse 
         Caption         =   "&Parse"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.TextBox txtHTML 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmMain.frx":0000
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Copyright (C) 2004 Kristian. S.Stangeland

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public HTML As clsParser
Public Document As clsDocument
Public TimeCount As New clsCount

Private Sub cmdExit_Click()

    ' Unload application
    Unload Me

End Sub

Private Sub cmdParse_Click()

    Dim objInput As clsLSInput

    ' Create the input-class
    Set objInput = CreateLSInput()

    ' Set the data to use
    objInput.StringData = txtHTML

    ' Measure the time taken to parse the HTML-document
    TimeCount.StartTimer

    ' Load the document
    Set Document = HTML.Parse(objInput)
    
    ' Stop clock
    TimeCount.StopTimer
    
    ' Show the time taken in the form
    frmParseTree.Caption = "Parse tree (Time: " & TimeCount.Elasped & " ms" & ")"
    
    ' Create the parse tree
    frmParseTree.LoadTree Document
    
    ' Show the form
    frmParseTree.Show
    
End Sub

Private Sub Form_Resize()

    ' Resize text box
    txtHTML.Width = Me.ScaleWidth - 16
    txtHTML.Height = Me.ScaleHeight - 24 - picToolBox.Height

    ' Move command buttons
    picToolBox.Top = txtHTML.Top + txtHTML.Height + 9

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim Form As Form

    ' Unload other forms
    For Each Form In Forms
        Unload Form
    Next

End Sub

Private Sub Form_Initialize()
    
    ' Necessary for XP-style
    InitCommonControls
    
    ' Create the parser
    Set HTML = CreateLSParser(0, "")

End Sub
