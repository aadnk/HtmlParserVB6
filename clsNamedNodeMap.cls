VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNamedNodeMap"
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

' How many enum operations that is currently under operation on this collection
Private m_nCurrentEnumOperations As Long

' All nodes
Private Nodes As New Collection
Attribute Nodes.VB_VarUserMemId = 0

' The main object
Private pMain As clsDocument

Friend Sub Initialize(cList As Collection)

    ' Set the internal collection
    Set Nodes = cList

End Sub

Public Property Get Lenght() As Long
Attribute Lenght.VB_Description = "The number of nodes in this map."

    ' Return the lenght
    Lenght = Nodes.Count - 1

End Property

Public Function SetNamedItem(arg As Object) As Object
Attribute SetNamedItem.VB_Description = "Adds a node using its nodeName attribute."

    Dim objTemp As Object, lngParentNode As Long

    ' Search for this node in the node map
    For Each objTemp In Nodes
        
        ' If the node index correspond, these nodes MUST be equal
        If objTemp.NodeName = arg.NodeName Then
        
            ' We need to replace the node depending on its type
            If TypeOf arg Is clsAttr Then
            
                ' Just set the attribute node
                pMain.GetElementByIndex(objTemp.NodeIndex).SetAttributeNode arg
            
                ' Return the other node
                Set SetNamedItem = objTemp
            
            Else
            
                ' Get the index of the parent node
                lngParentNode = objTemp.ParentNode.NodeIndex
            
                ' Here we need to set the current node as child and make the other childless
                pMain.SetChildIndex objTemp.ParentNode.NodeIndex, objTemp.ParentNode.FindChild(objTemp.NodeIndex), arg.NodeIndex
            
                ' Then remove the parent
                pMain.SetParent objTemp.NodeIndex, -1
                
                ' And set the parent of the other node
                pMain.SetParent arg.NodeIndex, lngParentNode
            
                ' Return the other node
                Set SetNamedItem = objTemp
            
            End If
            
            ' In any case we are finished
            Exit Function

        End If
        
    Next
    
    ' Add the element depending on its type
    If TypeOf arg Is clsAttr Then
    
        ' Just set the attribute node
        pMain.GetElementByIndex(objTemp.NodeIndex).SetAttributeNode arg
    
    Else
    
        ' Simply add this node by using the first element
        Item(0).ParentNode.AppendChild arg
     
    End If

    ' And of course add the node to this map
    Nodes.Add arg
    
    ' Return the node
    Set SetNamedItem = arg

End Function

Public Function IsEqual(refNode As clsNamedNodeMap) As Boolean

    Dim Tell As Long

    ' First, check the lenght
    If Lenght = refNode.Lenght Then
    
        ' Then check indexes
        For Tell = 0 To Lenght
        
            If Not Nodes(Tell) Is refNode.Item(Tell) Then
            
                ' Nope, the node is not equal
                Exit Function
            
            End If
        
        Next
    
    End If
    
    ' Yes, it is equal
    IsEqual = True

End Function

Public Function GetNamedItem(Name As String) As Object
Attribute GetNamedItem.VB_Description = "Retrieves a node specified by name."

    Dim Node As clsAttr

    ' Search for the node
    For Each Node In Nodes
    
        If Node.NodeName = Name Then
        
            ' Return this node
            Set GetNamedItem = Node
        
            ' We are finished
            Exit Function
        
        End If
    
    Next

End Function

Public Function RemoveNamedItem(Name As String) As Object
Attribute RemoveNamedItem.VB_Description = "Removes a node specified by name."

    ' Firstly, get the item
    Set RemoveNamedItem = GetNamedItem(Name)
    
    ' Then remove it
    Nodes.Remove Name

End Function

Property Get Item(ByVal Index As Long) As Object
Attribute Item.VB_Description = "Returns the indexth item in the map."

    ' Ignore all errors
    On Error Resume Next

    ' Get the node
    Set Item = Nodes.Item(Index + 1)
    
    ' Clear ALL errors
    Err.Clear
    
End Property

Public Property Get NewEnum() As IEnumVARIANT
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Dim oEnumerator As clsEnumator
    
    ' Create a new enumator to use in the enumation-process
    Set oEnumerator = New clsEnumator
    
    ' Initialize enumator
    oEnumerator.InitializeEnumeration ObjPtr(Me), 0, 0
    
    ' Count this operation
    m_nCurrentEnumOperations = m_nCurrentEnumOperations + 1
    
    ' Return the enumerator object's IEnumVARIANTReDef interface
    Set NewEnum = oEnumerator
    
    ' Clean up
    Set oEnumerator = Nothing
  
End Property

Public Sub DecrementEnumCounter()
Attribute DecrementEnumCounter.VB_MemberFlags = "40"

    ' We're finish with one operation
    m_nCurrentEnumOperations = m_nCurrentEnumOperations - 1

End Sub

Private Sub Class_Initialize()
  
    ' Add this collection to the list
    LookupList.AddPointerToLookupList ObjPtr(Me)
  
End Sub

Private Sub Class_Terminate()

    ' Remove this collection from the list
    LookupList.RemovePointerFromLookupList ObjPtr(Me)
  
End Sub
