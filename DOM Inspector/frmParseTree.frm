VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmParseTree 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Parse tree"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   753
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSelection 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   6840
      MousePointer    =   9  'Size W E
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   7
      Top             =   120
      Width           =   120
   End
   Begin VB.PictureBox picFooter 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6960
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   4
      Top             =   6480
      Width           =   4215
      Begin VB.CommandButton cmdUseChanges 
         Caption         =   "&Use changes"
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   2280
         TabIndex        =   5
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox picLeftFrame 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   6960
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtNodeValue 
         Height          =   3615
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2640
         Width           =   4215
      End
      Begin ComctlLib.ListView lstAttributes 
         Height          =   2535
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Value"
            Object.Width           =   3810
         EndProperty
      End
   End
   Begin ComctlLib.TreeView htmlTree 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11033
      _Version        =   327682
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Menu mnuItem 
      Caption         =   "Item"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit item"
      End
      Begin VB.Menu mnuRemoveItem 
         Caption         =   "&Remove item"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddItem 
         Caption         =   "&Add new"
      End
   End
   Begin VB.Menu mnuNode 
      Caption         =   "Node"
      Visible         =   0   'False
      Begin VB.Menu mnuAddNode 
         Caption         =   "Add new"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditNode 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeleteNode 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyXML 
         Caption         =   "&Copy XML"
      End
   End
End
Attribute VB_Name = "frmParseTree"
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

' API-calls necessary
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

' Used in getting the state of the left mouse-button
Private Const VK_LBUTTON = &H1

' The type used in getting the cursor position
Private Type POINTAPI
    X As Long
    Y As Long
End Type

' The selected nodes
Dim SelNode As clsElement
Dim SelItem As ListItem

Public Sub LoadTree(Document As clsDocument)

    ' Remove all
    htmlTree.Nodes.Clear

    ' Then add the document
    AddChildren Nothing, Document.DocumentElement

End Sub

Public Sub AddChildren(ParentNode As Node, Element As clsElement)

    Dim Tell As Long, Node As Node, Child As Object
    
    ' Add this element
    If ParentNode Is Nothing Then
    
        ' Add as root
        Set Node = htmlTree.Nodes.Add(, , Element.NodeIndex & "." & Element.NodeName, Element.NodeName)
        
        ' Open the node
        Node.Expanded = True
    
    Else
        ' Add as children
        Set Node = htmlTree.Nodes.Add(ParentNode.Key, tvwChild, Element.NodeIndex & "." & Element.NodeName, Element.NodeName)
    End If

    ' Does this element has child nodes?
    If Element.HasChildNodes Then

        ' Then add all children
        For Each Child In Element.ChildNodes
        
            ' Add the children
            AddChildren Node, Child
            
        Next

    End If
    
End Sub

Private Sub cmdCancel_Click()

    ' Just unload form
    Unload Me

End Sub

Private Sub cmdUseChanges_Click()

    ' We'll also convert the changes into HTML and set the textbox
    frmMain.txtHTML = frmMain.Document.DocumentElement.InnerHTML

    ' Hide this form
    Unload Me

End Sub

Private Sub Form_Resize()

    ' Set sizes
    lstAttributes.Width = picLeftFrame.Width
    txtNodeValue.Width = picLeftFrame.Width

    ' Resize treeview
    htmlTree.Width = Me.ScaleWidth - lstAttributes.Width - 24
    htmlTree.Height = Me.ScaleHeight - 24 - picFooter.Height

    ' Then move the frame and footer
    picLeftFrame.Left = htmlTree.Left + htmlTree.Width + 8
    picFooter.Left = picLeftFrame.Left

    ' And resize textbox and frame
    txtNodeValue.Height = Me.ScaleHeight - lstAttributes.Height - 30 - picFooter.Height
    picLeftFrame.Height = htmlTree.Height
    
    ' Move footer
    picFooter.Top = picLeftFrame.Top + picLeftFrame.Height + 8

    ' Move and resize resizer
    picSelection.Left = htmlTree.Left + htmlTree.Width
    picSelection.Height = htmlTree.Height

End Sub

Private Sub Form_Initialize()
    
    ' Necessary for XP-style
    InitCommonControls

End Sub

Private Sub htmlTree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim TreeNode As Node
 
    ' Only continue if this is a right click
    If Button = 2 Then

        ' Get the node
        Set TreeNode = htmlTree.HitTest(X, Y)

        ' Deleting when no node is selected should not be possible
        mnuDeleteNode.Enabled = Not (TreeNode Is Nothing)

        ' Don't try to select when no node is found
        If Not TreeNode Is Nothing Then
               
            ' Select the node
            TreeNode.Selected = True
        
            ' Initialize node-element
            htmlTree_NodeClick TreeNode
            
        End If
    
        ' Show the popup-menu
        Me.PopupMenu mnuNode
    
    End If

End Sub

Private Sub htmlTree_NodeClick(ByVal Node As ComctlLib.Node)

    Dim aArray As Variant, Element As clsElement
    
    ' Get this element's node index
    aArray = Split(Node.Key, ".")
    
    ' Only proceed if the index is valid
    If Val(aArray(0)) >= 0 Then
        
        ' Get the element
        Set Element = frmMain.Document.GetElementByIndex(Val(aArray(0)))
        
        ' Remember this node
        Set SelNode = Element
        
        ' Set the textbox
        txtNodeValue.Text = Element.NodeValue
        
        ' Update the list box
        UpdateListView Element
    
    End If

End Sub

Public Sub UpdateListView(Element As clsElement)

    Dim vAttribute As clsAttr
    
    ' Clear the list of attributes
    lstAttributes.ListItems.Clear
    
    ' Go through all, if there are any attributes at all
    For Each vAttribute In Element.Attributes
    
        ' Add the attribute
        lstAttributes.ListItems.Add(, , vAttribute.NodeName).SubItems(1) = vAttribute.NodeValue
    
    Next
        
End Sub

Private Sub lstAttributes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Must be a right-click.
    If Button = 2 Then
    
        ' Get the element at this location
        Set SelItem = lstAttributes.HitTest(X, Y)
        
        ' Don't try to select a item that dosen't exist
        If Not SelItem Is Nothing Then
            ' Select this element
            SelItem.Selected = True
        End If

        ' "Edit" and "Remove" can only be valid when a item is selected
        mnuEdit.Enabled = (Not SelItem Is Nothing)
        mnuRemoveItem.Enabled = mnuEdit.Enabled
    
        ' Show the menu
        Me.PopupMenu mnuItem
    
    End If

End Sub

Private Sub mnuAddItem_Click()

    Dim Dialog As New clsDialog, PropertyGet As PropertyBag

    ' Firstly, create a window for the dialog
    Set Dialog.ReferenceForm = New frmAttribute
    
    ' Then show the dialog and get the result
    Set PropertyGet = Dialog.ShowDialog(New PropertyBag, "Create attribute")
    
    ' Remove form
    Set Dialog.ReferenceForm = Nothing
    
    ' Only continue of the user pressed OK
    If PropertyGet.ReadProperty("Returned", "") = "OK" Then
    
        ' Create the attribute
        SelNode.SetAttribute PropertyGet.ReadProperty("txtName", ""), PropertyGet.ReadProperty("txtValue", "")
    
        ' Update the list
        UpdateListView SelNode
    
    End If
    
End Sub

Private Sub mnuAddNode_Click()
    
    Dim objParent As clsElement
    
    ' Are a node selected
    If Not SelNode Is Nothing Then
        ' Get the current node
        Set objParent = SelNode
    End If
    
    ' Firstly, create a node
    Set SelNode = frmMain.Document.CreateElement("", True)
    
    ' Add it as a child to the last selected node
    If Not objParent Is Nothing Then
        ' Add the child
        objParent.AppendChild SelNode
    End If
    
    ' Then set its properties
    mnuEditNode_Click

End Sub

Private Sub mnuCopyXML_Click()

    ' Clear clipboard
    Clipboard.Clear
    
    ' Copy the XML to the clipboard
    Clipboard.SetText SelNode.OuterHTML

End Sub

Private Sub mnuDeleteNode_Click()

    ' Delete the selected node
    SelNode.ParentNode.RemoveChild SelNode

    ' Remove the node from the treeview
    htmlTree.Nodes.Remove htmlTree.SelectedItem.Index

End Sub

Private Sub mnuEdit_Click()

    Dim CurrAttribute As Object, Dialog As New clsDialog
    Dim PropertySend As New PropertyBag, PropertyGet As PropertyBag

    ' Get the current attribute
    Set CurrAttribute = SelNode.Attributes.GetNamedItem(SelItem.Text)
    
    ' Only continue if it really exist
    If Not CurrAttribute Is Nothing Then
    
        ' Create a window for the dialog
        Set Dialog.ReferenceForm = New frmAttribute
        
        ' Set properties to the dialog
        PropertySend.WriteProperty "txtName", CurrAttribute.NodeName
        PropertySend.WriteProperty "txtValue", CurrAttribute.NodeValue
        
        ' Show dialog and get the result
        Set PropertyGet = Dialog.ShowDialog(PropertySend, "Change attribute")
        
        ' Remove form
        Set Dialog.ReferenceForm = Nothing
        
        ' Only continue of the user pressed OK
        If PropertyGet.ReadProperty("Returned", "") = "OK" Then
        
            ' Change node
            CurrAttribute.NodeName = PropertyGet.ReadProperty("txtName", "")
            CurrAttribute.NodeValue = PropertyGet.ReadProperty("txtValue", "")
        
            ' Save the changes
            SelNode.SetAttributeNode CurrAttribute
        
            ' The list box must of course now be updated
            UpdateListView SelNode
    
        End If
    
    End If


End Sub

Private Sub mnuEditNode_Click()

    Dim CurrAttribute As Object, Dialog As New clsDialog
    Dim PropertySend As New PropertyBag, PropertyGet As PropertyBag
    Dim lngParent As Long, lngNewParent As Long, ParentNode As clsElement

    ' Only continue if a node is in fact selected
    If Not SelNode Is Nothing Then
    
        ' Create a window for the dialog
        Set Dialog.ReferenceForm = New frmNode
    
        ' Get parent
        Set ParentNode = SelNode.ParentNode
    
        ' Find it's parrent index
        If ParentNode Is Nothing Then
            lngParent = -1
        Else
            lngParent = ParentNode.NodeIndex
        End If
    
        ' Initialize properties to send
        PropertySend.WriteProperty "txtName", SelNode.NodeName, ""
        PropertySend.WriteProperty "cmbType", SelNode.NodeType - 1, 0
        PropertySend.WriteProperty "txtParent", lngParent
    
        ' Show dialog
        Set PropertyGet = Dialog.ShowDialog(PropertySend, "Edit node")
        
        ' Remove form
        Set Dialog.ReferenceForm = Nothing
        
        ' Only continue if we've got an OK
        If PropertyGet.ReadProperty("Returned", "") = "OK" Then
        
            ' Change node
            SelNode.NodeName = PropertyGet.ReadProperty("txtName", "")
            SelNode.NodeType = PropertyGet.ReadProperty("cmbType", 0) + 1
            
            ' Get the parent
            lngNewParent = Val(PropertyGet.ReadProperty("txtParent", 0))
            
            ' Set parent if this has changed
            If lngParent <> lngNewParent Then
            
                ' Get the parent node if this node
                Set ParentNode = SelNode.ParentNode
            
                ' Then, remove this node from its parent
                If Not ParentNode Is Nothing Then
                    SelNode.ParentNode.RemoveChild SelNode
                End If
                
                ' And finally add it to the right element
                If lngNewParent >= 0 Then
                
                    ' Add the node
                    SelNode.OwnerDocument.GetElementByIndex(lngNewParent).AppendChild SelNode
                
                End If
                
            End If
            
            ' Update list
            LoadTree SelNode.OwnerDocument
        
        End If
    
    End If

End Sub

Private Sub mnuRemoveItem_Click()

    Dim CurrAttribute As Object

    ' Get the current attribute
    Set CurrAttribute = SelNode.Attributes.GetNamedItem(SelItem.Text)

    ' Only continue if this attribute exist
    If Not CurrAttribute Is Nothing Then
    
        ' Delete the attribute
        SelNode.RemoveAttributeNode CurrAttribute
        
        ' Update the list
        UpdateListView SelNode

    End If

End Sub

Private Sub picSelection_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim pointCursor As POINTAPI
    
    ' Show the selection of the picture box by changing its color
    picSelection.BackColor = vbBlue
    
    ' Move it whilst the cursor's position change till the user drops the button
    Do Until GetAsyncKeyState(VK_LBUTTON) = 0
    
        ' Get the position of the cursor
        GetCursorPos pointCursor
        
        ' Then move the selection
        picSelection.Left = pointCursor.X - (Me.Left / Screen.TwipsPerPixelX) - X
    
        ' Update events
        DoEvents
        
        ' Don't use 100% CPU
        Sleep 10
    
    Loop
    
    ' Set the size of the attribute control
    picLeftFrame.Width = Me.ScaleWidth - picSelection.Left - 16
    
    ' Reset color of the selection
    picSelection.BackColor = vbButtonFace
    
    ' Update and resize controls
    Form_Resize
    
End Sub

Private Sub txtNodeValue_Change()

    ' See if the node exists
    If Not SelNode Is Nothing Then

        ' Change the node's value
        SelNode.NodeValue = txtNodeValue.Text
    
    End If
    
End Sub
