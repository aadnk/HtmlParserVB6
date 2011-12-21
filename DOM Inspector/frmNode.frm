VERSION 5.00
Begin VB.Form frmNode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit node"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtParent 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   1680
      Width           =   3735
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmNode.frx":0000
      Left            =   2160
      List            =   "frmNode.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label lblParent 
      Caption         =   "&Parent:"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblName 
      Caption         =   "&Name:"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblType 
      Caption         =   "&Type:"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmNode"
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

Private Sub cmbType_Change()

    ' Do different things
    Select Case cmbType.ListIndex + 1
    Case DOMComment
    
        ' A comment must have this name
        txtName.Text = "#comment"
        txtName.Enabled = False
        
    Case DOMText
    
        ' Same to this
        txtName.Text = "#text"
        txtName.Enabled = False
    
    Case Else
    
        ' Allow all names
        txtName.Enabled = True
    
    End Select

End Sub

Private Sub cmbType_Click()

    ' Update controls
    cmbType_Change

End Sub

Private Sub cmbType_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Update controls
    cmbType_Change

End Sub

Private Sub cmdCancel_Click()

    ' Just hide form
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Dim lngType As Long

    ' Get the type
    lngType = cmbType.ListIndex + 1
    
    ' Don't create invalid types
    If lngType = DOMAttribute Or lngType = DOMDocument Then
    
        ' Show error
        MsgBox "Cannot create node: The type is not allowed", vbCritical, "Error"
    
        ' Don't accept this
        Exit Sub
        
    End If

    ' Confirm that the user pressed OK
    Me.Tag = "OK"
    
    ' Hide form
    Me.Hide

End Sub

Private Sub Form_Initialize()
        
    ' Necessary for XP-style
    InitCommonControls

End Sub

