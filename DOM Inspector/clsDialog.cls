VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public ReferenceForm As Form

Public Function ShowDialog(Properties As PropertyBag, Optional Caption As String) As PropertyBag

    Dim txtBox As Object, aArray
    
    ' Initialize form
    ReferenceForm.Tag = ""
    
    ' If we have receved a caption, use it
    If Caption <> "" Then
        ReferenceForm.Caption = Caption
    End If
    
    ' Create a new PropertyBag
    Set ShowDialog = New PropertyBag
    
    ' Loop through all textboxes and comboboxes, and give them the passed property
    For Each txtBox In ReferenceForm.Controls
        If TypeOf txtBox Is TextBox Then
        
            ' Set the text of this textbox
            txtBox.Text = Properties.ReadProperty(txtBox.Name, "")
            
        ElseIf TypeOf txtBox Is ComboBox Then
    
            ' Set the index
            txtBox.ListIndex = Properties.ReadProperty(txtBox.Name, 0)

        End If
    Next
        
    ' Show form
    ReferenceForm.Show
    
    ' Wait for it to finish
    Do Until ReferenceForm.Tag <> "" Or Not ReferenceForm.Visible
        Sleep 10
        DoEvents
    Loop
    
    ' Write what the form "returned"
    ShowDialog.WriteProperty "Returned", UCase(ReferenceForm.Tag), ""
    
    ' The user has pressed OK
    If UCase(ReferenceForm.Tag) = "OK" Then
    
        ' Add all textboxes and comboboxes to the propertybag
        For Each txtBox In ReferenceForm.Controls
            If TypeOf txtBox Is TextBox Then
                ShowDialog.WriteProperty txtBox.Name, txtBox.Text, ""
            ElseIf TypeOf txtBox Is ComboBox Then
                ShowDialog.WriteProperty txtBox.Name, txtBox.ListIndex, ""
            End If
        Next
    
    End If

End Function



