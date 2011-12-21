VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImplementationLS"
Attribute VB_GlobalNameSpace = True
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

' Operation types
Public Enum OperationTypes
    Node_Cloned = 1
    Node_Imported = 2
    Node_Deleted = 3
    Node_Renamed = 4
    Node_Adopted = 5
End Enum

' Constants for whatToShow
Public Enum FilterShow
    Show_All = &HFFFFFFFF
    Show_Element = &H1
    Show_Attribute = &H2
    Show_Text = &H4
    Show_CData_Section = &H8
    Show_Entity_Reference = &H10
    Show_Entity = &H20
    Show_Processing_Instruction = &H40
    Show_Comment = &H80
    Show_Document = &H100
    Show_Document_Type = &H200
    Show_Document_Fragment = &H400
    Show_Notation = &H800
End Enum

' The mode constants
Public Enum ModeConst
    Mode_Synchronous = 1
    Mode_Asynchronous = 2
End Enum

' Constants returned by startElement and acceptNode
Public Enum Filter
    Filter_Accept = 1
    Filter_Reject = 2
    Filter_Skip = 3
    Filter_Interrupt = 4
End Enum

' Node position
Public Enum DocumentPosition
    DocumentPosition_Disconnected = &H1
    DocumentPosition_Preceding = &H2
    DocumentPosition_Following = &H4
    DocumentPosition_Contains = &H8
    DocumentPosition_ContainedBy = &H10
    DocumentPosition_ImplementationSpecific = &H20
End Enum

' The different node types
Public Enum NodeType
    DOMElement = 1
    DOMAttribute = 2
    DOMText = 3
    DOMCDataSection = 4
    DOMEntityReference = 5
    DOMEntity = 6
    DOMProcessingInstruction = 7
    DOMComment = 8
    DOMDocument = 9
    DOMDocumentType = 10
    DOMNotation = 12
End Enum

Public Enum ActionTypes
    Action_AppendAsChildren = 1
    Action_ReplaceChildren = 2
    Action_InsertBefore = 3
    Action_InsertAfter = 4
    Action_Replace = 5
End Enum

Public Function CreateLSParser(Mode As Long, SchemaType As String) As clsParser
Attribute CreateLSParser.VB_Description = "Create a new LSParser."

    ' Create the parser
    Set CreateLSParser = New clsParser

End Function

Public Function CreateLSSerializer() As clsLSSerializer
Attribute CreateLSSerializer.VB_Description = "Create a new LSSerializer object."

    ' Create serializer
    Set CreateLSSerializer = New clsLSSerializer

End Function

Public Function CreateLSInput() As clsLSInput
Attribute CreateLSInput.VB_Description = "Create a new empty input source object."

    ' Create the new input
    Set CreateLSInput = New clsLSInput

End Function

Public Function CreateLSOutput() As clsLSOutput
Attribute CreateLSOutput.VB_Description = "Create a new empty output destination object."

    ' Create output
    Set CreateLSOutput = New clsLSOutput

End Function
