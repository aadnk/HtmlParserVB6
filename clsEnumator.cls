VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnumator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

' Here is where the magic happens.  the IEnumVARIANT interface has been
' redefined in the 'ienumv.tlb' type library to make it VB friendly.
Implements IEnumVReDef.IEnumVARIANTReDef

' Enumerator counter for IEnumVARIANT implementation
Private m_nCurrentEnumRef As Long

' Change this if you want something other than a one based collection
Private m_nBaseIndex As Long

Private m_lpObjCollection As Long

Private Sub Class_Initialize()

    Dim oIev As IEnumVARIANTReDef
    Set oIev = Me
    
    ' Replace the vtable entry for the Next function on this interface
    ReplaceVtableEntry ObjPtr(oIev), 4, AddressOf IEnumVARIANT_Next
    
    Set oIev = Nothing

End Sub

Private Sub Class_Terminate()

    ' decrement the enum counter in the collection object
    If LookupList.IsCollectionPointerValid(m_lpObjCollection) > (-1) Then
        LookupList.ResolveCollectionPointer(m_lpObjCollection).DecrementEnumCounter
    End If

End Sub

Private Sub IEnumVARIANTReDef_Clone(lppIEnum As IEnumVReDef.IEnumVARIANTReDef)

  Dim oEnumerator As clsEnumator
  
  ' Create new enumator
  Set oEnumerator = New clsEnumator
  
  oEnumerator.InitializeEnumeration m_lpObjCollection, m_nCurrentEnumRef, m_nBaseIndex
  
  Set lppIEnum = oEnumerator
  Set oEnumerator = Nothing

End Sub

Private Sub IEnumVARIANTReDef_GetItems(ByVal cElements As Long, aVariants As Variant, ByVal lpcElementsFetched As Long, lRetVal As Long)

    On Error Resume Next

    If LookupList.IsCollectionPointerValid(m_lpObjCollection) > (-1) Then
    
        ' Then get the current item
        Set aVariants = LookupList.ResolveCollectionPointer(m_lpObjCollection).Item(m_nCurrentEnumRef)
        
        ' And finally, see if we did invoke an error
        If Not aVariants Is Nothing Then
        
            ' increment the module level counter
            m_nCurrentEnumRef = m_nCurrentEnumRef + 1
            
            ' Everything went well
            lRetVal = S_OK
        
        Else
        
            ' Nope, terminate the enumation (this is most likely due to reaching the end of the array)
            lRetVal = S_FALSE
        
        End If
        
    Else
        ' Something has gone terrible wrong
        lRetVal = S_FALSE
    End If

End Sub

Private Sub IEnumVARIANTReDef_Next(ByVal cElements As Long, aVariants As Variant, ByVal lpcElementsFetched As Long)
  ' This function is delegated to the IEnumVARIANT_Next function in our MIEnumVARIANT bas module.
  ' Right before the actual enumeration begins, we alter the vtable of this object so that it
  ' points to the address of the IEnumVARIANT_Next function in the bas module
End Sub

Private Sub IEnumVARIANTReDef_Reset()

    ' Reset the pointer to the beginning of the collection
    m_nCurrentEnumRef = m_nBaseIndex
    
End Sub

Private Sub IEnumVARIANTReDef_Skip(ByVal cElements As Long)

    ' Increment the enum counter the specified number of elements to skip
    m_nCurrentEnumRef = m_nCurrentEnumRef + cElements
  
End Sub

Friend Function InitializeEnumeration(ByVal lpObjCollection&, ByVal nCurrentRef&, ByVal nLbound&) As Boolean
    
    m_lpObjCollection = lpObjCollection
    m_nBaseIndex = nLbound
    m_nCurrentEnumRef = nCurrentRef
  
End Function


