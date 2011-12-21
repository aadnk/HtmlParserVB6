VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLSSerializer"
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
 
'Public DomConfig As DOMConfiguration,
Public NewLine As String
'Public LSSerializerFilter As Filter

' This function saves the document
Public Function WriteToURI(NodeArg As Object, Uri As String, Optional bAppend As Boolean) As Long
Attribute WriteToURI.VB_Description = " A convenience method that acts as if LSSerializer.write was called with a LSOutput with no encoding specified."

    Dim Free As Long
    
    If Not bAppend Then
    
        ' Delete file if it exists
        If Dir(Uri) <> "" Then
        
            ' Delete the file
            Kill Uri
        
        End If
    
    End If
    
    ' Create the directory
    SHCreateDirectoryEx ByVal 0&, Left(Uri, InStrRev(Uri, "\")), ByVal 0&
    
    ' Get a free handle
    Free = FreeFile
    
    ' Add the text to the file
    Open Uri For Binary As #Free
    
        ' Add text
        Put #Free, , WriteToString(NodeArg)
    
    Close #Free

End Function

Public Function WriteToOutput(NodeArg As Object, objOutput As clsLSOutput) As Long
Attribute WriteToOutput.VB_Description = "Serialize the specified node as described above in the general description of the LSSerializer interface. The output is written to the supplied LSOutput."

    ' Does the class exist
    If Not objOutput Is Nothing Then
    
        ' Then simply set the data of the output
        objOutput.ByteStream = WriteToString(NodeArg)
        
    End If

End Function

Public Function WriteToString(NodeArg As Object) As String
Attribute WriteToString.VB_Description = "Serialize the specified node as described above in the general description of the LSSerializer interface. The output is written to a DOMString that is returned to the caller."

    If TypeOf NodeArg Is clsElement Then
    
        ' Get the XML-string of this element and return it
        WriteToString = NodeArg.InnerHTML
    
    ElseIf TypeOf NodeArg Is clsDocument Then
    
        ' As the above, only here access the HTML from the first element of the document
        WriteToString = NodeArg.DocumentElement.InnerHTML
    
    End If

End Function

Private Sub Class_Initialize()

    ' Default values
    NewLine = vbCrLf

End Sub
