VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLookupList"
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

' Array to hold pointers to all instances of the collection class. This allows us to verify
' the existence of the collection object before making a callback from an enumerator object.
Private m_alpCollections() As Long

Public Sub AddPointerToLookupList(ByVal lpObject&)

    ' Pass this task on to the FindPointer routine
    FindPointer lpObject, True
    
End Sub

' Removes a pointer from the array. This function is called from the terminate event of the collection class
Public Sub RemovePointerFromLookupList(ByVal lpObject&)

    Dim nItem As Long, nUbound As Long
    
    ' first find the location of the pointer in the array
    nItem = LookupList.IsCollectionPointerValid(lpObject)
    
    ' if we found it, remove it and shift the rest of the items....
    If nItem > (-1) Then
    
        nUbound = UBound(m_alpCollections)
        
        If nItem < nUbound Then
            ' grab all of the items below the moved one and shift them up
            CopyMemory ByVal VarPtr(m_alpCollections(nItem)), ByVal VarPtr(m_alpCollections(nItem + 1)), (nUbound - nItem) * 4&
        End If
        
        If nUbound Then
            ' if there are other items in the array, preserve them
            ReDim Preserve m_alpCollections(nUbound - 1) As Long
        Else
            ' if this is the last item in the array, just redim the array and make sure the value is zero
            ReDim m_alpCollections(0) As Long
            m_alpCollections(0) = 0
        End If
        
    End If
    
End Sub
    
Public Function IsCollectionPointerValid(ByVal lpObject&) As Long

    ' Checks whether or not the object still exists in the array.  for all intents
    ' and purposes, if the object isn't in the array, it no longer exists.
    IsCollectionPointerValid = FindPointer(lpObject)

End Function

' Function to provide fast lookups for pointers in the array
Private Function FindPointer(ByVal lpObject&, Optional ByVal bAddIfNotFound As Boolean) As Long
    
    Static bInitialized As Boolean
    
    Dim i As Long, nLow As Long, nHigh As Long, nUbound As Long
    
    ' Make sure the array is initialized
    If bInitialized = False Then
        If bAddIfNotFound Then
            GoTo AddFirsItem
        Else
            FindPointer = (-1)
        End If
    End If
    
    nHigh = UBound(m_alpCollections)
    
    ' Loop through the array looking for the object pointer.
    ' the array is in numerical order so we can do fast lookups
    Do
        ' Divide and conquer!  Each time we loop, devide the difference between the
        ' last items checked and search between the two indexes.  This is MUCH faster
        ' than looping through the entire list when dealing with a sorted array.
        i = nLow + ((nHigh - nLow) / 2)
        
        ' See how sKey relates to the current index....
        Select Case m_alpCollections(i)
          Case Is = lpObject
          
            FindPointer = i
            Exit Do
            
          Case Is > lpObject: nHigh = i - 1
          Case Is < lpObject: nLow = i + 1
        End Select
        
        ' If the low search bound has become greater than the high search bound, the
        ' item does not exist in the array.  if the bAddIfNotFound flag is set, a new
        ' item is being added.  otherwise, just return the not found value.
        If nLow > nHigh Then
        
            If bAddIfNotFound Then
              
AddFirsItem:
                ' check to see whether or not this item is initialized
                If Not bInitialized Then
                    bInitialized = True
                    ReDim m_alpCollections(0) As Long
                Else
                      
                    If m_alpCollections(0) <> 0 Then
                        ReDim Preserve m_alpCollections(UBound(m_alpCollections) + 1) As Long
                    End If
                  
                    nUbound = UBound(m_alpCollections)
                    
                    ' see whether we should add this item above or below the item at index 'i'
                    Select Case m_alpCollections(i)
                    Case Is < lpObject: i = i + 1
                    Case Is > lpObject: i = i '<- included for self documentation
                    End Select
                    
                    If i > nUbound Then
                        i = nUbound
                    ElseIf i < nUbound Then
                      ' Grab all of the items above the moved one and shift them down
                      CopyMemory ByVal VarPtr(m_alpCollections(i + 1)), ByVal VarPtr(m_alpCollections(i)), (nUbound - i) * 4&
                    End If
                    
                End If
              
                ' place the new pointer into the position that used to be held by the target
                m_alpCollections(i) = lpObject
                  
            End If ' bAddIfNotFound
            
            ' return value of KEY_NOT_FOUND tells the caller no match was found
            FindPointer = (-1)
            
            Exit Do
            
        End If ' nLow > nHigh
        
    Loop

End Function

Public Function ResolveCollectionPointer(ByVal lpObj&) As Object

    ' Used by all of the enumerator and key objects to resolve
    ' unreferenced pointers back to the parent collection class
    
    Dim oSC As Object
    
    CopyMemory oSC, lpObj, 4
    Set ResolveCollectionPointer = oSC
    CopyMemory oSC, 0&, 4

End Function

