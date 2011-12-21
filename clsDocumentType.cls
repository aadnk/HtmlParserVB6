VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumentType"
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

' Simple public variables holding all the information
Public Name As String
Attribute Name.VB_VarDescription = "The name of DTD; i.e., the name immediately following the DOCTYPE keyword."
Public Entities As New clsNamedNodeMap
Attribute Entities.VB_VarDescription = "A NamedNodeMap containing the general entities, both external and internal, declared in the DTD."
Public Notations As New clsNamedNodeMap
Attribute Notations.VB_VarDescription = "A NamedNodeMap containing the notations declared in the DTD. Duplicates are discarded. Every node in this map also implements the Notation interface."
Public PublicId As String
Attribute PublicId.VB_VarDescription = "The public identifier of the external subset."
Public SystemId As String
Attribute SystemId.VB_VarDescription = "The system identifier of the external subset. This may be an absolute URI or not."
Public InternalSubset As String
Attribute InternalSubset.VB_VarDescription = "The internal subset as a string, or null if there is none. This is does not contain the delimiting square brackets."
