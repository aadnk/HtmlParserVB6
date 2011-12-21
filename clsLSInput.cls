VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLSInput"
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

'Public LSReader As CharacterStream
Public StringData As String
Attribute StringData.VB_VarDescription = "String data to parse."
Public SystemId As String
Attribute SystemId.VB_VarDescription = "The system identifier, a URI reference, for this input source."
Public PublicId As String
Attribute PublicId.VB_VarDescription = "The public identifier for this input source."
Public BaseURI As String
Attribute BaseURI.VB_VarDescription = "The base URI to be used for resolving a relative systemId  to an absolute URI."
Public Encoding As String
Attribute Encoding.VB_VarDescription = "The character encoding, if known."
Public CertifiedText As Boolean
Attribute CertifiedText.VB_VarDescription = " If set to true, assume that the input is certified when parsing."

Property Get LSInputStream() As Byte()

    ' Return the string data
    LSInputStream = StringData
    
End Property
