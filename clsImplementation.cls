VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImplementation"
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

Public Function CreateDocumentType(QualifiedName As String, PublicId As String, SystemId As String) As clsDocumentType
Attribute CreateDocumentType.VB_Description = "Creates an empty DocumentType node."

    ' Create the class
    Set CreateDocumentType = New clsDocumentType
    
    ' Then set the properties
    CreateDocumentType.Name = QualifiedName
    CreateDocumentType.PublicId = PublicId
    CreateDocumentType.SystemId = SystemId

End Function

Public Function CreateDocument(NamespaceURI As String, QualifiedName As String, DocType As clsDocumentType) As clsDocument
Attribute CreateDocument.VB_Description = "Creates a DOM Document object of the specified type with its document element."

    ' Create the document class
    Set CreateDocument = New clsDocument
    
    ' Set the doctype
    CreateDocument.DocType = DocType
    CreateDocument.AppendElement "#document", 0, -1, DOMDocument, True, ""

    ' Reference implementation
    Set CreateDocument.Implementation = Me

End Function

Public Function GetFeature(Feature As String, Version As String) As Object
Attribute GetFeature.VB_Description = " This method returns a specialized object which implements the specialized APIs of the specified feature and version."
End Function

Property Get HasFeature(Feature As String, Version As String) As Boolean
Attribute HasFeature.VB_Description = "Test if the DOM implementation implements a specific feature and version."
End Property
