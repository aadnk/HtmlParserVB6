VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

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

' The array containing all elements (points to the array in the main-class)
Private Elements() As HTMLElement

' The index and parent of this node
Public NodeIndex As Long
Attribute NodeIndex.VB_VarDescription = "The unique index of this node."

' The main object
Private pMain As clsDocument

Public Function AppendChild(newChild As clsElement) As clsElement
Attribute AppendChild.VB_Description = "Adds the node newChild to the end of the list of children of this node. If the newChild is already in the tree, it is first removed."

    ' Add the child to this node
    pMain.AppendChild NodeIndex, newChild.NodeIndex

End Function

Public Function InsertBefore(newChild As clsElement, refChild As clsElement) As clsElement
Attribute InsertBefore.VB_Description = "Inserts the node newChild before the existing child node refChild."

    ' Add the child to this node
    pMain.AppendChild NodeIndex, newChild.NodeIndex, refChild.NodeIndex
    
    ' Return the node being inserted
    Set InsertBefore = newChild

End Function

Public Sub RemoveAllChildren()
Attribute RemoveAllChildren.VB_Description = "Removes all children of this element."

    ' Remove all children of this node
    pMain.RemoveChildren NodeIndex

End Sub

Public Function RemoveChild(oldChild As clsElement) As clsElement
Attribute RemoveChild.VB_Description = "Removes the child node indicated by oldChild from the list of children, and returns it."

    Dim lngFind As Long, Tell As Long
    
    ' Try to locate the child
    lngFind = FindChild(oldChild.NodeIndex)

    ' Only proceed if we found the child
    If lngFind >= 0 Then
    
        ' Remove the node, firstly by moving all elements above it downwards
        For Tell = lngFind To Elements(NodeIndex).ChildCount - 2
            Elements(NodeIndex).Children(Tell) = Elements(NodeIndex).Children(Tell + 1)
        Next
            
        ' After whitch we need to decrease the child count
        Elements(NodeIndex).ChildCount = Elements(NodeIndex).ChildCount - 1
        
        If Elements(NodeIndex).ChildCount > 0 Then
        
            ' Then resize the array
            ReDim Preserve Elements(NodeIndex).Children(Elements(NodeIndex).ChildCount - 1)

        Else
        
            ' We have no children, so just erase the array
            Erase Elements(NodeIndex).Children
        
        End If

        ' The removed node has now NO parent
        Elements(oldChild.NodeIndex).Parent = -1
        
        ' Return the node replaced
        Set RemoveChild = oldChild

    End If

End Function

Public Function ReplaceChild(newChild As clsElement, oldChild As clsElement) As clsElement
Attribute ReplaceChild.VB_Description = "Replaces the child node oldChild with newChild in the list of children, and returns the oldChild node."

    Dim lngFind As Long
    
    ' Try to locate the child
    lngFind = FindChild(oldChild.NodeIndex)

    ' Only proceed if we found the child
    If lngFind >= 0 Then
    
        ' Replace the node
        Elements(NodeIndex).Children(lngFind) = newChild.NodeIndex

        ' Return the node replaced
        Set ReplaceChild = oldChild

    End If

End Function

Public Property Get ChildNodes() As clsNodeList
Attribute ChildNodes.VB_Description = "A NodeList that contains all children of this node. If there are no children, this is a NodeList containing no nodes."
Attribute ChildNodes.VB_UserMemId = 0

    ' Return the parent node
    Set ChildNodes = pMain.CreateNodeList(Elements(NodeIndex).Children)

End Property

Public Property Get ParentNode() As clsElement
Attribute ParentNode.VB_Description = "The parent of this node."

    ' Return the parent node
    Set ParentNode = pMain.GetElementByIndex(Elements(NodeIndex).Parent)

End Property

Public Property Get PreviousSibling() As clsElement
Attribute PreviousSibling.VB_Description = "The node immediately preceding this node."

    ' Return the previous sibling
    Set PreviousSibling = GetParentChild(1, 0, -1)
    
End Property

Public Property Get NextSibling() As clsElement
Attribute NextSibling.VB_Description = "The node immediately following this node."

    ' Return the next sibling
    Set NextSibling = GetParentChild(0, -1, 1)
    
End Property

Public Property Get FirstChild() As clsElement
Attribute FirstChild.VB_Description = "The first child of this node. If there is no such node, this returns null."
 
    ' Only go further if there ARE children to get
    If Elements(NodeIndex).ChildCount > 0 Then
        ' Return the first child
        Set FirstChild = pMain.GetElementByIndex(Elements(NodeIndex).Children(0))
    End If

End Property

Public Property Get LastChild() As clsElement
Attribute LastChild.VB_Description = "The last child of this node. If there is no such node, this returns null."

    ' Only go further if there ARE children to get
    If Elements(NodeIndex).ChildCount > 0 Then
        ' Return the last child
        Set LastChild = pMain.GetElementByIndex(Elements(NodeIndex).Children(Elements(NodeIndex).ChildCount))
    End If

End Property

Public Function IsSameNode(otherNode As clsElement) As Boolean
Attribute IsSameNode.VB_Description = "Returns whether this node is the same node as the given one."

    ' See if this node is pointing to the same element
    IsSameNode = CBool(otherNode.NodeIndex = NodeIndex)
    
End Function

Public Property Get HasChildNodes() As Boolean
Attribute HasChildNodes.VB_Description = "Returns whether this node has any children."

    ' Return wether or not this node has children
    HasChildNodes = (Elements(NodeIndex).ChildCount > 0)

End Property

Public Property Get HasAttributes() As Boolean
Attribute HasAttributes.VB_Description = "Returns whether this node has any attributes."

    ' Return wether or not this node has atrributes
    HasAttributes = (Elements(NodeIndex).PropertyCount > 0)

End Property

Public Property Get NodeType() As Long
Attribute NodeType.VB_Description = "The type of this node."
    
    ' Return the node type
    NodeType = Elements(NodeIndex).NodeType

End Property

Public Property Let NodeType(ByVal lngNew As Long)
    
    ' Set the node type of this element
    Elements(NodeIndex).NodeType = lngNew

End Property

Public Property Get NodeName() As String
Attribute NodeName.VB_Description = "The name of this node, depending on its type."

    ' Return the node name
    NodeName = Elements(NodeIndex).TagName

End Property

Public Property Let NodeName(ByVal sNew As String)

    ' Set the node name
    Elements(NodeIndex).TagName = sNew
    
    ' Inform that the name has been changed
    pMain.InvokeHandlers NodeIndex, Node_Renamed, Me, Nothing

End Property

Public Property Get NodeValue() As String
Attribute NodeValue.VB_Description = "The value of this node, depending on its type."

    ' Return the value of this element
    NodeValue = Elements(NodeIndex).Value

End Property

Public Property Let NodeValue(ByVal sNew As String)

    ' Return the value of this element
    Elements(NodeIndex).Value = sNew

End Property

Public Property Get OwnerDocument() As clsDocument
Attribute OwnerDocument.VB_Description = "The Document object associated with this node."

    ' Return the document-class of this node
    Set OwnerDocument = pMain

End Property

Public Property Get Attributes() As clsNamedNodeMap
Attribute Attributes.VB_Description = "A NamedNodeMap containing the attributes of this node or null otherwise."
    
    Dim cList As New Collection, Tell As Long
    
    ' Create a list of attributes
    Set Attributes = New clsNamedNodeMap
    
    ' Add all attributes
    For Tell = 0 To Elements(NodeIndex).PropertyCount - 1
        ' Add the attribute
        cList.Add pMain.AttributeFromIndex(Me, NodeIndex, Tell), Elements(NodeIndex).Properties(Tell).Name
    Next
    
    ' Then add the elements to the list
    Attributes.Initialize cList
    
End Property

Public Function CloneNode(Deep As Boolean) As clsElement
Attribute CloneNode.VB_Description = "Returns a duplicate of this node, i.e., serves as a generic copy constructor for nodes. The duplicate node has no parent."

    ' Copy the node
    Set CloneNode = pMain.GetElementByIndex(pMain.CloneNode(NodeIndex, -1, Deep))
    
    ' Inform about this
    pMain.InvokeHandlers NodeIndex, Node_Cloned, Me, CloneNode
    
End Function

Public Function CompareDocumentPosition(other As clsElement) As Long
Attribute CompareDocumentPosition.VB_Description = "Compares the reference node, i.e. the node on which this method is being called, with a node."

    Dim lngParent As Long
    
    ' See if the node contains the reference node
    If FindChild(other.NodeIndex) >= 0 Then
    
        CompareDocumentPosition = DocumentPosition_Contains Or DocumentPosition_Preceding
    
    ' Does the reference node contain the current node
    Else
    
        lngParent = Elements(NodeIndex).Parent
    
        ' We need to search upwards
        Do Until lngParent < 0
        
            If FindChild(other.NodeIndex, lngParent) >= 0 Then
            
                ' Yup, this element is preceding
                CompareDocumentPosition = DocumentPosition_Preceding
                
                ' We are finished
                Exit Function
            
            End If
            
            ' Get the parent of this node
            lngParent = Elements(lngParent).Parent
        
        Loop
    
        ' We will now check for followingness
        lngParent = FindInChildren(other.NodeIndex, NodeIndex)
         
        ' See if the search was successful
        If lngParent < 0 Then
        
            ' Nope, so we assume it isn't connected
            CompareDocumentPosition = DocumentPosition_Disconnected
        
        Else
        
            ' Is the parent element we found equal to this element
            If NodeIndex = lngParent Then
        
                CompareDocumentPosition = DocumentPosition_ContainedBy Or DocumentPosition_Following

            Else
            
                CompareDocumentPosition = DocumentPosition_ContainedBy
            
            End If
        
        End If
    
    End If
    
End Function

Property Get TextContent() As String
Attribute TextContent.VB_Description = "This attribute returns the text content of this node and its descendants."

    Dim Tell As Long, lngChildIndex As Long

    Select Case Elements(NodeIndex).NodeType
    Case DOMElement, DOMAttribute, DOMEntity, DOMEntityReference
        
        ' Combine the text content of all child nodes
        For Tell = 0 To Elements(NodeIndex).ChildCount - 1
        
            ' Get the index of this child
            lngChildIndex = Elements(NodeIndex).Children(Tell)
        
            If Elements(lngChildIndex).NodeType <> DOMComment And Elements(lngChildIndex).NodeType <> DOMProcessingInstruction Then
        
                ' Add the text content
                TextContent = TextContent & pMain.GetElementByIndex(lngChildIndex).TextContent
            
            End If
            
        Next
        
    Case DOMText, DOMCDataSection, DOMComment, DOMProcessingInstruction
    
        ' Return this elements value
        TextContent = Elements(NodeIndex).Value
    
    End Select

End Property

Property Let TextContent(ByVal sNewText As String)

    ' Simply remove all children, if any
    If Elements(NodeIndex).ChildCount > 0 Then
        RemoveAllChildren
    End If

    ' Add a new text node containing the text
    pMain.AppendElement "#text", -1, NodeIndex, DOMText, True, sNewText

End Property

Public Function IsEqualNode(arg As clsElement) As Boolean
Attribute IsEqualNode.VB_Description = "Tests whether two nodes are equal."

    If Elements(NodeIndex).NodeType = arg.NodeType Then
    
        If Elements(NodeIndex).TagName = arg.NodeName Then
        
            If Elements(NodeIndex).Value = arg.NodeValue Then
            
                If ChildNodes.IsEqual(arg.ChildNodes) Then
                  
                    If Attributes.IsEqual(arg.Attributes) Then
                  
                        ' Yes, this is a equal node
                        IsEqualNode = True
                    
                    End If
                  
                End If
            
            End If
            
        End If
        
    End If

End Function

Public Function SetUserData(Key As String, Data As Object, Handler As Object) As Object
Attribute SetUserData.VB_Description = "Associate an object to a key on this node."

    Dim lngIndex As Long

    ' See if this user data exists
    lngIndex = FindUserData(Key)
    
    If lngIndex < 0 Then
    
        ' Reallocate array
        ReDim Preserve Elements(NodeIndex).UserData(Elements(NodeIndex).UserDataCount)
        
        ' Set values
        Set Elements(NodeIndex).UserData(Elements(NodeIndex).UserDataCount).Data = Data
        Set Elements(NodeIndex).UserData(Elements(NodeIndex).UserDataCount).Handler = Handler
        Elements(NodeIndex).UserData(Elements(NodeIndex).UserDataCount).Key = Key
        
        ' Increse the count
        Elements(NodeIndex).UserDataCount = Elements(NodeIndex).UserDataCount + 1
    
    Else
    
        ' Return the data object of the latter object
        Set SetUserData = Elements(NodeIndex).UserData(lngIndex).Data
    
        ' Set values
        Set Elements(NodeIndex).UserData(lngIndex).Data = Data
        Set Elements(NodeIndex).UserData(lngIndex).Handler = Handler
        Elements(NodeIndex).UserData(lngIndex).Key = Key
    
    End If

End Function

Public Function GetUserData(Key As String) As Object
Attribute GetUserData.VB_Description = "Retrieves the object associated to a key on a this node."

    Dim lngIndex As Long
    
    ' Look after this user data
    lngIndex = FindUserData(Key)
    
    ' Is this in the array
    If lngIndex >= 0 Then
    
        ' Return the user data
        Set GetUserData = Elements(NodeIndex).UserData(lngIndex).Data
    
    End If

End Function

Public Function GetAttributeNode(Name As String) As clsAttr
Attribute GetAttributeNode.VB_Description = "Retrieves an attribute node by name."
    
    Dim lngIndex As Long
    
    ' Find the index
    lngIndex = FindAttribute(Name)
    
    ' Did we find the attribute
    If lngIndex >= 0 Then
    
        ' Return the attribute
        Set GetAttributeNode = pMain.AttributeFromIndex(Me, NodeIndex, lngIndex)
        
    End If

End Function

Public Function SetAttributeNode(newAttr As clsAttr) As clsAttr
Attribute SetAttributeNode.VB_Description = "Adds a new attribute node. If an attribute with that name (nodeName) is already present in the element, it is replaced by the new one."

    Dim lngIndex As Long

    ' Locate the attribute
    lngIndex = FindAttribute(newAttr.Name)
    
    ' Then, see if it's in the element
    If lngIndex >= 0 Then
    
        ' Replace it
        Elements(NodeIndex).Properties(lngIndex).Value = newAttr.Value
        Elements(NodeIndex).Properties(lngIndex).Name = newAttr.Name
        Elements(NodeIndex).Properties(lngIndex).ID = newAttr.IsID
    
        ' Return the attribute
        Set SetAttributeNode = pMain.AttributeFromIndex(Me, NodeIndex, lngIndex)
    
    Else

        ' Add the attribute
        lngIndex = pMain.AppendAttribute(NodeIndex, newAttr.Name, newAttr.Value, newAttr.IsID)
    
        ' Return the attribute
        Set SetAttributeNode = pMain.AttributeFromIndex(Me, NodeIndex, lngIndex)
    
    End If

End Function

Public Function RemoveAttributeNode(oldAttr As clsAttr) As clsAttr
Attribute RemoveAttributeNode.VB_Description = "Removes the specified attribute node"

    Dim lngIndex As Long, Tell As Long

    ' Locate the attribute
    lngIndex = FindAttribute(oldAttr.Name)
    
    ' Did we find the attribute
    If lngIndex >= 0 Then
    
        ' Move all attributes above this attribute downwards
        For Tell = lngIndex To Elements(NodeIndex).PropertyCount - 2
        
            ' Set the attribute above to this
            Let Elements(NodeIndex).Properties(Tell) = Elements(NodeIndex).Properties(Tell + 1)
        
        Next
    
        ' Then decrease count
        Elements(NodeIndex).PropertyCount = Elements(NodeIndex).PropertyCount - 1
        
        ' And resize array, or erase it
        If Elements(NodeIndex).PropertyCount <= 0 Then
            
            ' Erase array
            Erase Elements(NodeIndex).Properties
        
        Else
        
            ' Resize array
            ReDim Preserve Elements(NodeIndex).Properties(Elements(NodeIndex).PropertyCount - 1)
        
        End If
    
    End If

End Function

Public Sub SetIdAttribute(Name As String, IsID As Boolean)
Attribute SetIdAttribute.VB_Description = "This method declares the specified attribute to be a user-determined ID attribute."

    Dim lngIndex As Long
    
    ' Find the index
    lngIndex = FindAttribute(Name)
        
    If lngIndex >= 0 Then
    
        ' Set the ID
        Elements(NodeIndex).Properties(lngIndex).ID = IsID
        
    End If

End Sub
  
Public Sub SetIdAttributeNode(idAttr As clsAttr, IsID As Boolean)
Attribute SetIdAttributeNode.VB_Description = "This method declares the specified attribute to be a user-determined ID attribute."

    ' See if the attribute is in the node
    If NodeIndex = idAttr.NodeIndex Then
    
        ' Set the ID
        Elements(NodeIndex).Properties(idAttr.AttrIndex).ID = IsID
    
    End If

End Sub

Public Function GetElementsByTagName(Name As String) As clsNodeList
Attribute GetElementsByTagName.VB_Description = "eturns a NodeList of all descendant Elements  with a given tag name, in document order."

    ' Get all elements with this name below this element
    Set GetElementsByTagName = pMain.GetElementsByTagName(Name, NodeIndex)

End Function

Public Property Get HasAttribute(Name) As Boolean
Attribute HasAttribute.VB_Description = "Returns true when an attribute with a given name is specified on this element or has a default value, false  otherwise."

    Dim Tell As Long
    
    ' Go through all attributes
    For Tell = 0 To Elements(NodeIndex).PropertyCount - 1
    
        ' See if this attribute corresponds to what we're searching for
        If Elements(NodeIndex).Properties(Tell).Name = Name Then
        
            ' Return true
            HasAttribute = True
        
            ' We are finished
            Exit Property
        
        End If
    
    Next

End Property

Public Function GetAttribute(Name As String) As String
Attribute GetAttribute.VB_Description = "Retrieves an attribute value by name."

    Dim lngIndex As Long
    
    ' Get the index
    lngIndex = FindAttribute(Name)
    
    ' Only go further if we did find the attribute
    If lngIndex >= 0 Then
    
        ' Return the value
        GetAttribute = Elements(NodeIndex).Properties(lngIndex).Value
    
    End If

End Function

Public Sub SetAttribute(Name As String, Value As String)
Attribute SetAttribute.VB_Description = "Adds a new attribute. If an attribute with that name is already present in the element, its value is changed to be that of the value parameter."

    Dim lngIndex As Long
    
    ' Get the index
    lngIndex = FindAttribute(Name)
    
    ' Only go further if we did find the attribute
    If lngIndex >= 0 Then
    
        ' Set the value
        Elements(NodeIndex).Properties(lngIndex).Value = Value
    
    Else
    
        ' Add the attribute
        pMain.AppendAttribute NodeIndex, Name, Value
    
    End If
    
End Sub

Public Sub SwapChildren(nextChild As clsElement)
Attribute SwapChildren.VB_Description = "Swaps the two nodes given."

    Dim ElementTemp As HTMLElement
    
    ' We first need to save the content of the first element
    LSet ElementTemp = Elements(NodeIndex)
    
    ' Then, swap the elements
    LSet Elements(NodeIndex) = Elements(nextChild.NodeIndex)
    LSet Elements(nextChild.NodeIndex) = ElementTemp

End Sub

' Return the HTML-code for this object
Public Property Get OuterHTML() As String
Attribute OuterHTML.VB_Description = "The HTML-/XML-code representing this node and all its under-elements."
    
    Dim objString As New clsStringBuilder
    
    ' Get the inside-HTML together with the tags of this element
    pMain.GetElementHTML NodeIndex, objString

    ' Then, return this
    OuterHTML = objString.ToString

End Property

Public Property Get InnerHTML() As String
Attribute InnerHTML.VB_Description = "The HTML-/XML-code representation of all child nodes. "
    
    Dim objString As New clsStringBuilder, Tell As Long

    ' Only append inner data
    If Elements(NodeIndex).ChildCount = 0 Then
        
        ' Append value
        objString.Append Elements(NodeIndex).Value
    
    Else
    
        For Tell = 0 To Elements(NodeIndex).ChildCount - 1
            ' Add the childs inner HTML
            pMain.GetElementHTML Elements(NodeIndex).Children(Tell), objString
        Next
        
    End If

    ' Then, return this
    InnerHTML = objString.ToString

End Property

Friend Sub Initialize(lpArray() As HTMLElement, HTMLMain As clsDocument)

    ' This will make the local private array pointing towards the array referenced in the paramenter
    CopyMemory ByVal VarPtrArray(Elements), ByVal VarPtrArray(lpArray), 4
    
    ' Set the main HTML-class
    Set pMain = HTMLMain

End Sub

Private Sub Class_Terminate()

    Dim lngArray As Long
    
    ' Get the memory location of the array
    lngArray = VarPtrArray(Elements)
    
    ' Don't go further if there isn't anything to remove
    If lngArray <> 0 Then
        ' Remove reference to the array, thereby avoiding destroying it unintentionally
        ZeroMemory ByVal lngArray, 4
    End If
    
End Sub

Private Function FindAttribute(Name As String) As Long

    Dim Tell As Long
    
    ' Locate the attribute
    For Tell = 0 To Elements(NodeIndex).PropertyCount - 1
    
        ' See if this is the attribute
        If Elements(NodeIndex).Properties(Tell).Name = Name Then

            ' Return the index
            FindAttribute = Tell
            
            ' We are finished
            Exit Function

        End If
        
    Next
    
    ' Nope, no attribute with this name
    FindAttribute = -1

End Function

Friend Function FindChild(lngIndex As Long, Optional lngParentNode) As Long

    Dim Tell As Long
    
    If IsMissing(lngParentNode) Then
    
        ' Set it to the current index
        lngParentNode = NodeIndex
    
    Else
    
        ' If it's not missing, it MUST be valid
        If Not IsNumeric(lngParentNode) Then
        
            ' Nope, set to default
            lngParentNode = NodeIndex
        
        End If
    
    End If
    
    ' We must first find the old child in the children list
    For Tell = 0 To Elements(lngParentNode).ChildCount - 1
    
        ' See if the child and the node is equal
        If Elements(lngParentNode).Children(Tell) = lngIndex Then
            
            ' Return the position
            FindChild = Tell
        
            ' We are finished
            Exit Function
        
        End If
        
    Next

    ' We didn't find anything
    FindChild = -1

End Function

Private Function FindUserData(sKey As String) As Long

    Dim Tell As Long
    
    For Tell = 0 To Elements(NodeIndex).UserDataCount - 1
    
        ' See if this is equal to the given key
        If Elements(NodeIndex).UserData(Tell).Key = sKey Then
        
            ' Yup, we have found the user data
            FindUserData = Tell
            
            ' We are finished
            Exit Function
        
        End If
    
    Next
    
    ' Nope, we didn't find anything
    FindUserData = -1

End Function

Private Function FindInChildren(lngIndex As Long, lngParent As Long) As Long

    Dim Tell As Long
    
    ' Check all parents
    For Tell = 0 To Elements(lngParent).ChildCount - 1
    
        ' Is this children equal to what we are searching for?
        If Elements(lngParent).Children(Tell) = lngIndex Then
        
            ' Return the parent
            FindInChildren = lngParent
            
            ' We are finished
            Exit Function
        
        Else
        
            ' Nope, so search in its children
            FindInChildren = FindInChildren(lngIndex, Elements(lngParent).Children(Tell))
        
            ' If this was successful, return
            If FindInChildren >= 0 Then
                Exit Function
            End If

        End If
    
    Next

    ' We didn't find anything
    FindInChildren = -1

End Function

Private Function GetParentChild(lngStartAdd As Long, lngEndAdd As Long, lngChildFound As Long) As clsElement

    Dim Tell As Long, lngParent As Long
    
    ' Get the parent of this node
    lngParent = Elements(NodeIndex).Parent

    ' The parent must of course be valid
    If lngParent >= 0 Then

        ' Go trough all children of the parent node
        For Tell = lngStartAdd To Elements(lngParent).ChildCount - 1 + lngEndAdd
         
            ' Not yet implemented
            If Elements(lngParent).Children(Tell) = NodeIndex Then
            
                ' Return the child below this
                Set GetParentChild = pMain.GetElementByIndex(Elements(lngParent).Children(Tell + lngChildFound))
            
            End If
        
        Next
    
    End If

End Function

Public Property Get Target() As String
Attribute Target.VB_Description = "The content of the processing instruction."

    ' Used in processing instruction
    Target = Me.NodeValue

End Property

Public Property Get Data() As String
Attribute Data.VB_Description = "The character data of the node that implements this interface."

    ' Used in processing instruction
    Data = Me.NodeValue

End Property

Public Property Let Data(ByVal sValue As String)

    ' Used in processing instruction
    Me.NodeValue = sValue

End Property

Public Property Get Length() As Long
Attribute Length.VB_Description = "The number of 16-bit units  that are available through data and the substringData methods."

    ' The lenght of the data
    Length = Len(Me.NodeValue)

End Property

Public Function SubstringData(ByVal Offset As Long, ByVal Count As Long) As String
Attribute SubstringData.VB_Description = "Extracts a range of data from the node."

    ' Extracts a string from the data
    SubstringData = Mid(Me.NodeValue, Offset + 1, Count)

End Function

Public Sub AppendData(ByVal arg As String)
Attribute AppendData.VB_Description = "Append the string to the end of the character data of the node."

    ' Add the string
    Elements(NodeIndex).Value = Elements(NodeIndex).Value & arg

End Sub

Public Sub InsertData(ByVal Offset As Long, ByVal arg As String)
Attribute InsertData.VB_Description = "Insert a string at the specified 16-bit unit offset."

    ' Insert the text into a specified position in the string
    Elements(NodeIndex).Value = Mid(Elements(NodeIndex).Value, 1, Offset) & arg & Mid(Elements(NodeIndex).Value, Offset + 2)
 
End Sub

Public Sub DeleteData(ByVal Offset As Long, ByvalCount As Long)
Attribute DeleteData.VB_Description = "Remove a range of 16-bit units  from the node."

    ' Removes a part of the data
    Elements(NodeIndex).Value = Mid(Elements(NodeIndex).Value, 1, Offset) & Mid(Elements(NodeIndex).Value, Offset + 2)

End Sub
                                                                          
Public Sub ReplaceData(ByVal Offset As Long, ByVal Count As Long, ByVal arg As String)
Attribute ReplaceData.VB_Description = "Replace the characters starting at the specified 16-bit unit offset with the specified string."

    ' Replaces a specified portion of the string
    Elements(NodeIndex).Value = Mid(Elements(NodeIndex).Value, 1, Offset) & arg & Mid(Elements(NodeIndex).Value, Offset + 2 + Count)
 
End Sub

