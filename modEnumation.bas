Attribute VB_Name = "modEnumation"
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

' Used to make it possible replacing vtable-entries
Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress&, ByVal dwSize&, ByVal flNewProtect&, lpflOldProtect&) As Long

' Constants used in the following code
Public Const S_OK As Long = &H0&
Public Const S_FALSE As Long = &H1&
Public Const KEY_NOT_FOUND As Long = (-&HEFFFFFFF)
Public Const ERROR_ALREADY_EXISTS As Long = &H800000B7
Public Const PAGE_EXECUTE_READWRITE As Long = &H40&

' The list of all collections
Public LookupList As New clsLookupList

Public Function IEnumVARIANT_Next(ByVal this As IEnumVReDef.IEnumVARIANTReDef, ByVal cElements As Long, avObjects As Variant, ByVal nNumFetched As Long) As Long

    On Error Resume Next
    
    Dim i&, lpVariantArray&, lRet&, nFetched&, nDummy&
    Dim vTmp As Variant, vEmpty As Variant

    ' Get the address of the first variant in array
    lpVariantArray = VarPtr(avObjects)
    
    ' Iterate through each requested variant
    For i = 1 To cElements
    
        ' Call the method that was added to the IEnumVARIANT interface in the typelib.
        ' nDummy is a space filler since the two params it occupies are not used in this implementation.
        ' lRet is the return value from the GetItems call passed byref
        this.GetItems nDummy, vTmp, nDummy, lRet
        
        ' If failure or nothing fetched, we're done
        If (Err) Or (lRet = 1) Then
            Exit For
        End If
        
        ' Copy variant to current array position
        CopyMemory ByVal lpVariantArray, vTmp, 16&
        ' Empty work variant without destroying its object or string
        CopyMemory vTmp, vEmpty, 16&
        ' Count the variant and point to the next one
        nFetched = nFetched + 1
        
        lpVariantArray = lpVariantArray + 16
    Next
    
    ' If error caused termination, undo what we did
    If Err.Number Then
        
        ' Iterate back, emptying the invalid fetched variants
        For i = i To 1 Step -1
            ' Copy variant to current array position
            CopyMemory vTmp, ByVal lpVariantArray, 16
            ' Empty work variant, destroying any object or string
            vTmp = Empty
            ' Empty array variant without destroying any object or string
            CopyMemory ByVal lpVariantArray, vEmpty, 16
            ' Point to previous array element
            lpVariantArray = lpVariantArray - 16
        Next
        
        ' Convert error to COM format
        IEnumVARIANT_Next = MapErr(Err)
        ' Return 0 as the number fetched after error
        If nNumFetched Then CopyMemory ByVal nNumFetched, ByVal 0&, 4&
      
    Else
    
        ' If nothing fetched, break out of enumeration
        If nFetched = 0 Then
            IEnumVARIANT_Next = S_FALSE '<-- the value of S_FALSE is &H1&.  Confusing, eh'?
        End If
        
        ' Copy the actual number fetched to the pointer to fetched count
        If nNumFetched Then
            CopyMemory ByVal nNumFetched, nFetched, 4&
        End If
        
    End If
  
End Function

' Put the function address (callback) directly into the object v-table
Public Function ReplaceVtableEntry(ByVal lpObj As Long, ByVal nEntry As Integer, ByVal lpFunc As Long) As Long
             
    Dim lpFuncOld&, lpVTableHead&, lpFuncTmp&, nOldProtect&
    
    ' Object pointer contains a pointer to v-table--copy it to temporary
    CopyMemory lpVTableHead, ByVal lpObj, 4&
    
    ' Calculate pointer to specified entry
    lpFuncTmp = lpVTableHead + (nEntry - 1) * 4
    
    ' Save address of previous method for return
    CopyMemory lpFuncOld, ByVal lpFuncTmp, 4&
    
    ' Ignore if they're already the same
    If lpFuncOld <> lpFunc Then
    
        ' Need to change page protection to write to code
        VirtualProtect lpFuncTmp, 4&, PAGE_EXECUTE_READWRITE, nOldProtect
        
        ' Write the new function address into the v-table
        CopyMemory ByVal lpFuncTmp, lpFunc, 4&
        
        ' Restore the previous page protection
        VirtualProtect lpFuncTmp, 4&, nOldProtect, nOldProtect
        
    End If
    
    ReplaceVtableEntry = lpFuncOld
    
End Function

Private Function MapErr(ByVal ErrNumber As Long) As Long
  
    If ErrNumber Then
        If (ErrNumber And &H80000000) Or (ErrNumber = 1) Then
            'Error HRESULT already set
            MapErr = ErrNumber
        Else
            'Map back to a basic error number
            MapErr = &H800A0000 Or ErrNumber
        End If
    End If
  
End Function
