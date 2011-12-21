VERSION 5.00
Begin VB.Form frmAttribute 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change attribute"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblValue 
      Caption         =   "&Value:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblName 
      Caption         =   "&Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Demo Project for HtmlParserVB6
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

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub cmdCancel_Click()

    ' Just hide form
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    ' Confirm that the user pressed OK
    Me.Tag = "OK"
    
    ' Hide form
    Me.Hide

End Sub

Private Sub Form_Initialize()
    
    ' Necessary for XP-style
    InitCommonControls

End Sub
