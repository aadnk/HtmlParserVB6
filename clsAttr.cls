VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAttr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

' Type info
'Property Get schemaTypeInfo() As clsTypeInfo
'End Property

' The index and parent of this node
Public NodeIndex As Long
Attribute NodeIndex.VB_VarDescription = "A unique index of the node."
Public AttrIndex As Long
Attribute AttrIndex.VB_VarDescription = "The index of this attribute."
Public OwnerElement As clsElement
Attribute OwnerElement.VB_VarDescription = "The Element node this attribute is attached to."
Public Name As String
Attribute Name.VB_VarDescription = "Returns the name of this attribute."
Public Specified As Boolean
Attribute Specified.VB_VarDescription = "True if this attribute was explicitly given a value in the instance document, false otherwise."
Public Value As String
Attribute Value.VB_VarDescription = "The value of the attribute; on setting, this creates a Text node with the unparsed contents of the string."
Public IsID As Boolean
Attribute IsID.VB_VarDescription = "Returns whether this attribute is known to be of type ID or not."

' The main object
Private pMain As clsDocument

Property Get NodeName() As String
Attribute NodeName.VB_Description = "The name of this node, depending on its type."

    ' Return the name
    NodeName = Name

End Property

Property Let NodeName(ByVal sNewValue As String)

    ' Set the name
    Name = sNewValue

End Property

Property Get NodeValue() As String

    ' Return the value
    NodeValue = Value
    
End Property

Property Let NodeValue(ByVal sNewValue As String)

    ' Set the name
    Value = sNewValue

End Property
