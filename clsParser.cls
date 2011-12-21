VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParser"
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

' A array that points to the corresponding array in the document
Private Elements() As HTMLElement

' All entities
Private Entities() As Entity

' A lookup table used in the process of assembling the document back together
Private eLookupTable() As Integer

' Whether or not we should ignore parsing
Private bAbort As Boolean

' Public variables
Public Busy As Boolean
Attribute Busy.VB_VarDescription = "True if the LSParser is currently busy loading a document, otherwise False."
'Public DomConfig As DOMConfiguration
Public Filter As Object
Attribute Filter.VB_VarDescription = "When a filter is provided, the implementation will call out to the filter as it is constructing the DOM tree structure."
Public Async As Boolean
Attribute Async.VB_VarDescription = "True if the LSParser is asynchronous, False if it is synchronous. "

Public Function Parse(objInput As clsLSInput) As clsDocument
Attribute Parse.VB_Description = "Parse an XML document from a resource identified by a LSInput."

    Dim Document As clsDocument, Implementation As New clsImplementation

    ' We are busy parsing
    Busy = True

    ' Firstly, create the document
    Set Document = Implementation.CreateDocument("", "", Nothing)
    
    ' Parse the document
    ParseText Document, objInput.StringData, 1, 0
    
    ' Return the document
    Set Parse = Document
    
    ' We are finished
    Busy = False

End Function

Friend Sub InitializeArray(intChar() As Integer, SafeArray As SAFEARRAY1D, sRefString As String)

    ' Firstly, point the integer array to the string
    With SafeArray
        .cDims = 1
        .Bounds(0).lLbound = 1
        .Bounds(0).cElements = Len(sRefString)
        .pvData = StrPtr(sRefString)
    End With
    
    ' Set the array
    CopyMemory ByVal VarPtrArray(intChar), VarPtr(SafeArray), 4

End Sub

Friend Sub ParseText(Document As clsDocument, sText As String, lngIndex As Long, lngParent As Long)

    Dim SafeArray As SAFEARRAY1D, intChar() As Integer, lngExit As Long

    ' Don't abort
    bAbort = False
    
    ' Make this array reference to the string
    InitializeArray intChar, SafeArray, sText
    
    ' Reference the arrays
    Document.Initialize Elements, Entities, eLookupTable, Me
    
    ' Start searching for tags
    HTMLTagProc lngIndex, lngParent, intChar, lngExit, Document
    
    ' Since tags without end tags cannot have children, those must change parent
    FixElements Document
    
    ' Refresh document
    Document.RefreshInfo
    
    ' Clean up
    ZeroMemory ByVal VarPtrArray(intChar), 4
    ZeroMemory ByVal VarPtrArray(Elements), 4

End Sub

Public Function ParseURI(Uri As String) As clsDocument
Attribute ParseURI.VB_Description = "Parse an XML document from a location identified by a URI reference."

    Dim sBuffer As String, Free As Long, objInput As New clsLSInput
    
    ' Get a free file
    Free = FreeFile
    
    ' Simply load all data in this file
    Open Uri For Binary As #Free

        ' Allocate buffer
        sBuffer = SysAllocStringByteLen(0, LOF(Free))
        
        ' Get data
        Get #Free, , sBuffer

    Close #Free
    
    ' Set the buffer
    objInput.BaseURI = Uri
    objInput.StringData = sBuffer
    
    ' Clear local buffer
    sBuffer = ""
    
    ' Parse the HTML by passing the code further on
    Set ParseURI = Parse(objInput)

End Function

Public Function ParseWithContext(objInput As clsLSInput, ContextArg As Object, Action As Long) As clsDocument
Attribute ParseWithContext.VB_Description = "Parse an XML fragment from a resource identified by a LSInput and insert the content into an existing document at the position specified with the context and action arguments."

    Dim aTemp() As Long, Temp As Long, lngIndex As Long, lngSearch As Long, Tell As Long, Document As clsDocument

    ' We are busy parsing
    Busy = True

    If TypeOf ContextArg Is clsDocument Then
    
        Select Case Action
        Case Action_ReplaceChildren

            ' In reality, this is like creating a new document
            Set ContextArg = Parse(objInput)

            ' Then return the created object
            Set ParseWithContext = ContextArg

        End Select
    
    ElseIf TypeOf ContextArg Is clsElement Then
    
        Select Case Action
        Case Action_AppendAsChildren, Action_ReplaceChildren
        
            If Action = Action_ReplaceChildren Then
            
                ' Remove all children
                ContextArg.RemoveAllChildren
            
            End If
        
            ' Add all parsed elements as children
            ParseText ContextArg.OwnerDocument, objInput.StringData, 1, ContextArg.NodeIndex
        
        Case Action_Replace
        
            ' Remove all children of the parent
            ContextArg.ParentNode.RemoveAllChildren
            
            ' Add new nodes
            ParseText ContextArg.OwnerDocument, objInput.StringData, 1, ContextArg.ParentNode.NodeIndex
        
        Case Action_InsertAfter
        
            ' Add all parsed elements after this element
            ParseText ContextArg.OwnerDocument, objInput.StringData, 1, ContextArg.ParentNode.NodeIndex

        Case Action_InsertBefore
        
            Set Document = ContextArg.OwnerDocument
        
            ' Reference the array
            Document.Initialize Elements, Entities, eLookupTable, Me
            
            ' Get the index of the parent node
            lngIndex = ContextArg.ParentNode.NodeIndex
            
            ' Index to search for
            lngSearch = ContextArg.NodeIndex
            
            ' Find the index in the children array
            For Tell = 0 To Elements(lngIndex).ChildCount - 1
            
                ' See if the index corresponds to what we're searching for
                If Elements(lngIndex).Children(Tell) = lngSearch Then
                
                    ' Add all elements above to the temp array
                    ReDim aTemp(Elements(lngIndex).ChildCount - 1 - Tell)
                
                    ' Move to temp array
                    For Temp = LBound(aTemp) To UBound(aTemp)
                
                        ' Set element
                        aTemp(Temp) = Elements(lngIndex).Children(Tell + Temp)
                    
                    Next
                    
                    ' Then remove those elements
                    ReDim Preserve Elements(lngIndex).Children(Tell - 1)
                    
                    ' And set the count
                    Elements(lngIndex).ChildCount = Tell
                    
                    ' Append elements
                    ParseText ContextArg.OwnerDocument, objInput.StringData, 1, ContextArg.ParentNode.NodeIndex

                    ' And add all in the temp array
                    For Temp = LBound(aTemp) To UBound(aTemp)
                        ' Add the remaining children
                        ContextArg.OwnerDocument.AppendChild lngIndex, aTemp(Temp)
                    Next
                
                    ' We are finished
                    Exit For
                
                End If
            
            Next
            
            ' Clean up
            ZeroMemory ByVal VarPtrArray(Elements), 4
    
        End Select
        
    End If
    
    ' We are finished
    Busy = False

End Function

Public Sub Abort()
Attribute Abort.VB_Description = "Abort the loading of the document that is currently being loaded by the LSParser."

    ' This function will be used to abort the download
    bAbort = True

End Sub

' This function will simply fix all the elements so that none has children that cannot have such
Private Sub FixElements(Document As clsDocument)

    Dim Tell As Long, Temp As Long, lngParent As Long

    For Tell = 0 To Document.ElementCount
    
        ' This must only be done with tags without an end tag
        If Not Elements(Tell).EndTag Then
        
            ' Check if there are any children
            If Elements(Tell).ChildCount > 0 Then
            
                ' Get the parent of this tag
                lngParent = Elements(Tell).Parent
            
                ' Move all children to its parent
                For Temp = 0 To Elements(Tell).ChildCount - 1
            
                    ' Add the children to the parent
                    Document.AppendChild lngParent, Elements(Tell).Children(Temp)
            
                Next
                
                ' Then, remove all children at this tag
                Document.RemoveChildren Tell
            
            End If
            
            ' Remove it if it is set to be removed
            If Elements(Tell).Reject Then
            
                ' Remove it and its children
                RemoveChildren Document, Tell
            
            End If
        
        End If
    
    Next

End Sub

Private Sub RemoveChildren(Document As clsDocument, lngIndex As Long)

    Dim Tell As Long
    
    For Tell = 0 To Elements(lngIndex).ChildCount - 1
        ' Remove children
        RemoveChildren Document, Elements(lngIndex).Children(Tell)
    Next
    
    ' Then remove this element
    Document.RemoveElement lngIndex

End Sub

Private Function FindValue(intStart As Long, lngEnd As Long, lngValue As Integer, intChar() As Integer) As Long

    Dim Tell As Long, Temp As Long, intCurrChar As Integer
    
    ' Start the loop that will search for the char
    For Tell = intStart To lngEnd
    
        ' The current character
        intCurrChar = intChar(Tell)
    
        If intCurrChar = QuotationMark Or intCurrChar = Apostrophe Then
        
            ' Find the next quotation mark or apostrophe
            For Temp = Tell + 1 To lngEnd
            
                ' See if this character is equal to the character we're searching for
                If intChar(Temp) = intCurrChar Then
                
                    ' We are finish searching
                    Exit For
                    
                End If
            
            Next
            
            ' Set the new position
            Tell = Temp
        
        Else
    
            If intCurrChar = lngValue Then
            
                ' Return the index
                FindValue = Tell
                
                ' Nothing more to do now
                Exit Function
            
            End If
    
        End If
    
    Next

End Function

Private Function HTMLTagProc(ByVal lngIndex As Long, ByVal lngParent As Long, intChar() As Integer, lngExit As Long, Document As clsDocument) As Long

    Dim Tell As Long, lngNextSpace As Long, lngNextGreater As Long, lngEqualSign As Long, bFound As Boolean
    Dim lngPropIndex As Long, lngSearch As Long, Temp As Long, lngAmout As Long, intNextChar As Integer
    Dim lngBegin As Long, bExitLoop As Boolean, sEndTag As String, lngPointer As Long, sName As String
    Dim lngType As NodeType, aPropName As Variant, ElementCount As Long, bCDATASection As Boolean

    ' The begin-index is at first the same as the start index
    lngBegin = lngIndex

    ' We ought to find all tags, begin searching from the index
    For Tell = lngIndex To UBound(intChar)

        ' Is this the begining of a new tag?
        If intChar(Tell) = LessThan Then
        
            ' Has the abort-procedure been invoked?
            If bAbort Then
                ' Simply just exit the function
                Exit Function
            End If
                    
            ' Get the char after the current
            intNextChar = intChar(Tell + 1)
        
            ' See if this is an end tag
            If intNextChar = Slash Then
                
                ' Get the next greater-than sign
                lngNextGreater = FindValue(Tell + 1, UBound(intChar), GreaterThan, intChar)
                
                ' Then, use that information to extract the name of the end tag
                sEndTag = ProcessSegment(intChar, Tell + 2, lngNextGreater, True)
                
                ' The first element to be tested is of course the parent
                lngPointer = lngParent
                
                ' Initialize this variable
                lngAmout = 0
                
                ' Assume that nothing was found
                bFound = False
                
                Do Until lngPointer < 0
                
                    ' See if the two strings is equal
                    If LenB(Elements(lngPointer).TagName) = LenB(sEndTag) Then
                        If Elements(lngPointer).TagName = sEndTag Then
                            
                            ' We've found the correct beginner tag
                            bFound = True
                            
                            ' No more searching needed
                            Exit Do
                        
                        End If
                    End If
                
                    ' Get the pointer of the next element
                    lngPointer = Elements(lngPointer).Parent
                
                    ' We have dived one more time into the hierarchy
                    lngAmout = lngAmout + 1
                
                Loop
                
                ' The pointer must be valid for us to take this end tag into account
                If lngPointer >= 0 Then
                
                    ' Only add a text-element if this element is not empty
                    If lngBegin <> Tell Then
                
                        ' Add the text as a text-tag
                        ElementCount = Document.AppendElement("#text", lngBegin, lngParent, DOMText, True, ProcessSegment(intChar, lngBegin, Tell))
                    
                        ' See if we should ignore this element
                        FilterElement Document, ElementCount, True
            
                    End If
                    
                    ' This is only needed when we've actually found a tag
                    If bFound Then
            
                        ' The tag has been closed
                        Elements(lngPointer).EndTag = True
                    
                        ' How many tag-searching prosedures that is to be stopped
                        lngExit = lngAmout
                    
                        ' Return the position of the first character after this end tag
                        HTMLTagProc = lngNextGreater
                    
                        ' We haven't got anything more to do
                        Exit Function
                    
                    End If
                
                End If
            
            Else

                ' See if this might be a CDATA-section
                bCDATASection = intChar(Tell + 2) = LeftSquareBracket

                ' Add the text element, if necessary
                If Tell - lngBegin > 0 Then
                
                    ' Add this text element
                    ElementCount = Document.AppendElement("#text", lngBegin, lngParent, DOMText, True, ProcessSegment(intChar, lngBegin, Tell))
                    
                    ' See if we should ignore this element
                    FilterElement Document, ElementCount, True
                    
                End If

                If intNextChar = ExclamationMark And (intChar(Tell + 2) = MinusSign And intChar(Tell + 3) = MinusSign Or bCDATASection) Then  ' Is this a comment tag or CDATA-tag?
            
                    ' Try to locate the next end-char
                    Temp = LocateSegment(intChar, IIf(bCDATASection, "]]>", "-->"), Tell + IIf(bCDATASection, 9, 4), UBound(intChar))

                    ' Create different tags depending on what we've found
                    If bCDATASection Then
                        ElementCount = Document.AppendElement("#cdata-section", lngBegin, lngParent, DOMCDataSection, True, GetString(intChar, Tell + 9, Temp))
                    Else
                        ' Create the comment tag
                        ElementCount = Document.AppendElement("#comment", lngBegin, lngParent, DOMComment, True, GetString(intChar, Tell + 4, Temp))
                    End If
                    
                    ' See if we should ignore this element
                    FilterElement Document, ElementCount, True
                    
                    ' Set the new begin position
                    lngBegin = Temp + 3
                    
                    ' Set the tell
                    Tell = Temp + 2
            
                Else
                
                    ' This could also a doctype-tag
                    If intNextChar = ExclamationMark Then
                    
                        ' Is this a doctype?
                        If LocateSegment(intChar, "DOCTYPE", Tell + 2, Tell + 9) = Tell + 2 Then
                    
                            ' Yup, it's a doctype
                            lngType = DOMDocumentType
                    
                        Else
                        
                            ' Nope, this must then be an entity
                            lngType = DOMEntity
                        
                        End If
                    
                        ' Next, we need to find out where we should read the doctype's name
                        lngNextSpace = FindValue(Tell + 1, UBound(intChar), SpaceChar, intChar)
                        
                        ' Go pass the doctype-declaration
                        Tell = lngNextSpace
                        
                    ElseIf intNextChar = QuestionMark Then
                    
                        ' This is a processing instruction
                        lngType = DOMProcessingInstruction
                        
                        ' We need to move one step further (because of the question mark)
                        Tell = Tell + 1
                    
                    Else
                    
                        ' If not, this is a regular element
                        lngType = DOMElement
                    
                    End If
            
                    ' Find the next space and next greater than sign
                    lngNextGreater = FindValue(Tell + 1, UBound(intChar), GreaterThan, intChar)
                    lngNextSpace = FindValue(Tell + 1, lngNextGreater - 1, SpaceChar, intChar)
                    
                    ' See whats smallest
                    If lngNextSpace < lngNextGreater And lngNextSpace <> 0 Then
                        Temp = lngNextSpace
                    Else
                        Temp = lngNextGreater
                    End If
                          
                    ' Add the element
                    ElementCount = Document.AppendElement(ProcessSegment(intChar, Tell + 1, Temp, True), Tell, lngParent, lngType, False)
                    
                    ' Add all properties of this tag, if there are some
                    If intChar(Temp) = SpaceChar Then
                    
                        ' Start searching after the space character
                        lngSearch = Temp + 1
                    
                        ' Add properties
                        Do Until lngSearch >= lngNextGreater
                            
                            ' Find the next space
                            lngNextSpace = FindValue(lngSearch, lngNextGreater - 1, SpaceChar, intChar)
                        
                            ' See if this space isn't follwed by a space. If so, loop until we have found a space that dosen't
                            Do
                                If intChar(lngNextSpace + 1) = SpaceChar Then
                                    lngNextSpace = lngNextSpace + 1
                                Else
                                    Exit Do
                                End If
                            Loop
                            
                            ' Continue if found
                            If lngNextSpace <= 0 Then
                            
                                ' Exit loop after adding this property
                                bExitLoop = True
                                
                                ' Set position
                                lngNextSpace = lngNextGreater
                            
                            End If
                            
                            ' Increse the count
                            If intChar(lngSearch) = Slash Or intChar(lngSearch) = QuestionMark Then
                                
                                ' We're finished
                                bExitLoop = True
                                
                            Else
                            
                                ' Find the equal sign
                                lngEqualSign = FindValue(lngSearch, lngNextGreater - 1, EqualSign, intChar)
                            
                                ' Get the index to the propery that is to be created
                                lngPropIndex = Elements(ElementCount).PropertyCount
                            
                                ' Create a new property
                                ReDim Preserve Elements(ElementCount).Properties(lngPropIndex)
                            
                                ' We need to save the property differently when it only exist as a variable
                                If lngEqualSign <= 0 Then
                                    
                                    If intChar(lngSearch) = QuotationMark Or intChar(lngSearch) = Apostrophe Then
                                        ' Save the name, without the qoutation mark or aprostrophe
                                        Elements(ElementCount).Properties(lngPropIndex).Name = ProcessSegment(intChar, lngSearch + 1, lngNextSpace - 1, True)
                                    Else
                                        ' Save the property name
                                        Elements(ElementCount).Properties(lngPropIndex).Name = ProcessSegment(intChar, lngSearch, lngNextSpace, True)
                                    End If
                            
                                Else
                                
                                    ' Here we need to save both the property AND the value
                                    If intChar(lngSearch) = QuotationMark Or intChar(lngSearch) = Apostrophe Then
                                        ' Save the name, without the qoutation mark or aprostrophe
                                        Elements(ElementCount).Properties(lngPropIndex).Name = ProcessSegment(intChar, lngSearch + 1, lngEqualSign - 1, True)
                                    Else
                                        ' Save the property name
                                        Elements(ElementCount).Properties(lngPropIndex).Name = ProcessSegment(intChar, lngSearch, lngEqualSign, True)
                                    End If
                                    
                                    If intChar(lngEqualSign + 1) = QuotationMark Or intChar(lngEqualSign + 1) = Apostrophe Then
                                        ' Save the value, without the qoutation mark or aprostrophe
                                        Elements(ElementCount).Properties(lngPropIndex).Value = ProcessSegment(intChar, lngEqualSign + 2, lngNextSpace - 1, True)
                                    Else
                                        ' Save the property value
                                        Elements(ElementCount).Properties(lngPropIndex).Value = ProcessSegment(intChar, lngEqualSign + 1, lngNextSpace, True)
                                    End If
                                    
                                End If
                            
                                ' Is this a ID-value
                                Elements(ElementCount).Properties(lngPropIndex).ID = CBool(LCase(Elements(ElementCount).Properties(lngPropIndex).Name) = "id" _
                                 Or LCase(Elements(ElementCount).Properties(lngPropIndex).Name) = "name")
                            
                                ' Increse count
                                Elements(ElementCount).PropertyCount = lngPropIndex + 1
                            
                            End If
                                                            
                            If bExitLoop Then
                                ' If not, exit do
                                Exit Do
                            End If
                        
                            ' Set to next search
                            lngSearch = lngNextSpace + 1
                        Loop
                                    
                    End If
    
                    ' See if we should ignore this element
                    If Not FilterElement(Document, ElementCount, False) Then
    
                    ' Is this element already closed
                        If Elements(ElementCount).EndTag Then
                        
                            ' We'll just move further
                            Tell = lngNextGreater + 1
                            lngBegin = lngNextGreater + 1
                            
                        Else
                        
                            ' Continue looking for new tags, this time inside this element
                            Tell = HTMLTagProc(lngNextGreater + 1, ElementCount, intChar, lngExit, Document)
                        
                            ' See if we should ignore this element
                            FilterElement Document, ElementCount, True
                            
                            ' If we are finished, exit prosedure
                            If Tell <= 0 Then
                                Exit Function
                            End If
            
                            ' Set the begin index to this position
                            lngBegin = Tell + 1
            
                            ' If some tags are to remove there search, do it
                            If lngExit > 0 Then
                            
                                ' We have stopped searching in one more tag
                                lngExit = lngExit - 1
                                
                                ' Just return the position we have ended on
                                HTMLTagProc = Tell
                                
                                ' Nothing more to do
                                Exit Function
                                
                            End If
                        
                        End If
                
                    End If
                
                End If
            
            End If
            
        End If
            
    Next
            
End Function

Friend Function FilterElement(Document As clsDocument, lngIndex As Long, Complete As Boolean) As Boolean

    ' Skip errors
    On Error Resume Next
    Dim lngResult As Filter, lngShow As FilterShow, Element As clsElement
    
    ' See if a filter is registered
    If Not Filter Is Nothing Then
    
        ' Get the element
        Set Element = Document.GetElementByIndex(lngIndex)
        
        ' Get the show constant
        lngShow = GetShowConstant(Element.NodeType)
        
        ' See if the filter requires us to call the functions
        If (Filter.WhatToShow And lngShow) = lngShow Then
        
            ' See what to call
            If Complete Then
            
                ' The element is finish - call acceptNode
                lngResult = Filter.AcceptNode()
    
            Else
            
                ' The element is not finish, so we'll call this function
                lngResult = Filter.StartElement(Document.GetElementByIndex(lngIndex))
                
            End If
            
            ' Do the appropriate course of action
            Select Case lngResult
            Case Filter_Interrupt
            
                ' Stop the processing of the document
                Abort
            
            Case Filter_Reject
            
                ' Reject the node and everything in it, including its children.
                Elements(lngIndex).Reject = True
                
                ' The element has been filtered
                FilterElement = True
                
            Case Filter_Skip
            
                ' Skip this element, simply by removing it
                Document.RemoveElement lngIndex
            
                ' Same here
                FilterElement = True
            
            End Select
        
        End If
        
    End If

End Function

Friend Function GetShowConstant(lngNodeType As NodeType) As FilterShow

    ' Obtain the show constant based on the node type
    Select Case lngNodeType
    Case DOMElement = 1: GetShowConstant = Show_Element
    Case DOMAttribute = 2: GetShowConstant = Show_Attribute
    Case DOMText = 3: GetShowConstant = Show_Text
    Case DOMCDataSection = 4: GetShowConstant = Show_CData_Section
    Case DOMEntityReference = 5: GetShowConstant = Show_Entity_Reference
    Case DOMEntity = 6: GetShowConstant = Show_Entity
    Case DOMProcessingInstruction = 7: GetShowConstant = Show_Processing_Instruction
    Case DOMComment = 8: GetShowConstant = Show_Comment
    Case DOMDocument = 9: GetShowConstant = Show_Document
    Case DOMDocumentType = 10: GetShowConstant = Show_Document_Type
    Case DOMNotation = 12: GetShowConstant = Show_Notation
    End Select

End Function

Private Sub LoadEntities(sFile As String)

    Dim sBuffer As String, Free As Long, Tell As Long, aLines, aEntity
    Dim lngMinValue As Long, lngMaxValue As Long
    
    ' Get a free file handle
    Free = FreeFile
    
    ' Load the file
    Open sFile For Binary As #Free
    
        ' Allocate buffer
        sBuffer = SysAllocStringByteLen(0, LOF(Free))
    
        ' Load data
        Get #Free, , sBuffer
    
    Close #Free
    
    ' Get all lines
    aLines = Split(sBuffer, vbNewLine)
    
    ' Then, make all the entities
    For Tell = LBound(aLines) To UBound(aLines)
    
        ' Firstly, split the entity
        aEntity = Split(aLines(Tell), ";")
        
        ' Only go further if this is a real line
        If UBound(aEntity) >= 1 Then
        
            ' Reallocate the array and lookup table
            ReDim Preserve Entities(Tell)
            
            ' Firstly, reallocate the arrays
            ReDim Entities(Tell).lngName(1 To Len(aEntity(0)))
            ReDim Entities(Tell).lngText(1 To 1)
            
            ' Then set the array
            CopyMemory Entities(Tell).lngName(1), ByVal StrPtr(aEntity(0)), LenB(aEntity(0))
            
            ' And the text it should be replaced with
            Entities(Tell).lngText(1) = aEntity(1)
            
            ' If we haven't checked two lines yet, just assume that the first entity is the biggest and smallest
            If Tell = LBound(aLines) Then
            
                ' Set the values
                lngMinValue = aEntity(1)
                lngMaxValue = aEntity(1)
            
            Else
            
                ' See if this is the least value
                If aEntity(1) < lngMinValue Then
        
                    ' Simply set the new value
                    lngMinValue = aEntity(1)
        
                End If
        
                ' See if this is the biggest value
                If aEntity(1) > lngMaxValue Then
                
                    ' As above
                    lngMaxValue = aEntity(1)
                
                End If
            
            End If
    
        End If
    
    Next
    
    ' Create the lookup table
    ReDim eLookupTable(lngMinValue To lngMaxValue)

    ' We need to set the default value of all elements within this array
    For Tell = lngMinValue To lngMaxValue
        eLookupTable(Tell) = -1
    Next

    ' Then add all entities to this table
    For Tell = LBound(Entities) To UBound(Entities)

        ' Set this table cell
        eLookupTable(Entities(Tell).lngText(1)) = Tell

    Next

End Sub

' This function compares a segment with another to find an occurrence (using VarPtrArray at small strings like this actually makes things worse)
Private Function LocateSegment(intChar() As Integer, sCompare As String, ByVal lngStart As Long, ByVal lngEnd As Long) As Long

    Dim Tell As Long, Temp As Long, intCompare() As Integer, lngLenght As Long, bCorresponding As Boolean
    
    ' Get the lenght of the compare-string
    lngLenght = Len(sCompare)
    
    ' Then, allocate the array copy
    ReDim intCompare(1 To lngLenght)
    
    ' And copy data
    CopyMemory intCompare(1), ByVal StrPtr(sCompare), lngLenght * 2

    ' Don't search pass the arrays
    If UBound(intChar) < lngEnd Then
        ' Set the new position to be precisely adequate
        lngEnd = UBound(intChar)
    End If
    
    ' The searching loop must also not bother looking for match when there's even not enough space left
    lngEnd = lngEnd - ((lngEnd - lngStart + 1) Mod lngLenght)
    
    ' Loop until we've found what we're looking for
    Do While lngStart <= lngEnd
    
        ' The start position of the other string
        Temp = 1
    
        ' Assume firstly that this position correspond - the next loop will determine if this is true or not
        bCorresponding = True
    
        ' Search through the array, beginning at the specified position
        For Tell = lngStart To lngStart + lngLenght - 1
        
            ' See if it's corresponding
            If intChar(Tell) <> intCompare(Temp) Then
            
                ' They didn't correspond
                bCorresponding = False

                ' If not, we need to search at the next position
                Exit For
        
            End If
            
            ' Go to next character
            Temp = Temp + 1
        
        Next
        
        ' See if it at this position actually correspond
        If bCorresponding Then
        
            ' Return position
            LocateSegment = lngStart
            
            ' We're finish with all
            Exit Function
        
        End If
        
        ' Try next character
        lngStart = lngStart + 1
    
    Loop
    
End Function

Private Function ProcessSegment(intChar() As Integer, lngStart As Long, lngEnd As Long, Optional IgnoreEndChar As Boolean) As String

    Dim intData() As Integer, Tell As Long, lngEndPos As Long, intNumber As Integer
    Dim Temp As Long, lngSemicolon As Long, Search As Long, currExponent As Long, lngTimes As Long
    Dim lngChar As Long, lngAdd As Long, bFound As Boolean, intOriginal As Integer, bHex As Boolean
    
    If lngEnd - lngStart < 1 Then
        ' Nothing to processs
        Exit Function
    End If
    
    ' Reallocate buffer copy
    ReDim intData(1 To (lngEnd - lngStart))

    ' Copy data from real array
    CopyMemory intData(1), ByVal (VarPtr(intChar(1)) + ((lngStart - 1) * 2)), (lngEnd - lngStart) * 2

    ' We begin at position one
    Temp = 1

    ' Next, do the parsing
    For Tell = 1 To UBound(intData)
    
        ' Is this a end-character
        If intData(Tell) = 13 Then
    
            ' Should we ignore end characters?
            If IgnoreEndChar Then
    
                If intData(Tell + 1) = 10 Then
                
                    ' Skip these two characters
                    Tell = Tell + 1
                
                Else
                
                    ' The character should still not be added to the return
                
                End If
            
            Else
            
                ' Process as normal
                GoTo Normal
            
            End If
        
        ElseIf intData(Tell) = Ampersand Then
        
            ' Find the end of this enity
            For lngSemicolon = Tell + 2 To UBound(intData)
            
                ' If the character is a semicolon, this is the end
                If intData(lngSemicolon) = Semicolon Then
                    
                    ' We are finished
                    Exit For
                    
                End If
                
                ' If this is a ampersand or space, something is terribly wrong
                If intData(lngSemicolon) = Ampersand Or intData(lngSemicolon) = SpaceChar Then
                    ' Just proceed processing as normal
                    GoTo Normal
                End If
            
            Next
        
            ' The position of the semicolon must not of course exceed the string to parse
            If lngSemicolon > UBound(intData) Then
                GoTo Normal
            End If
        
            ' Do what we need to do
            If intData(Tell + 1) = NumberSign Then
            
                ' See if this is a hex-value
                If intData(Tell + 2) = SmallX Then
                    lngTimes = 16
                    lngEndPos = Tell + 3
                    bHex = True
                Else ' This is a regular decimal number
                    lngTimes = 10
                    lngEndPos = Tell + 2
                    bHex = False
                End If
            
                ' Set the exponent
                currExponent = 1
            
                ' Save the original character
                intOriginal = intData(Temp)
            
                ' Firstly, clear the character
                intData(Temp) = 0
            
                For Search = lngSemicolon - 1 To lngEndPos Step -1
                
                    ' Get the number
                    intNumber = intData(Search)
                
                    ' Do extra processing with hex-numbers
                    If bHex Then
                    
                        If intNumber >= SmallA And intNumber <= SmallF Then
                            intNumber = intNumber - SmallA + IntZero + 10
                        End If
                        
                        If intNumber >= LargeA And intNumber <= LargeF Then
                            intNumber = intNumber - LargeA + IntZero + 10
                        End If
                
                    End If
                    
                    ' Get the real number
                    intNumber = intNumber - IntZero
                
                    ' See if it's valid
                    If intNumber >= 0 And intNumber <= lngTimes Then
                
                        ' Set the value
                        intData(Temp) = intData(Temp) + (intNumber * currExponent)
        
                        ' Increse exponent
                        currExponent = currExponent * lngTimes
                    
                    Else
                    
                        ' Set character back to original
                        intData(Temp) = intOriginal
                    
                        ' Process as normal
                        GoTo Normal
                    
                    End If
        
                Next
        
                ' The character has been added
                Temp = Temp + 1
                
                ' Set the new position
                Tell = lngSemicolon
            
            Else
            
                ' This is a name of which entity we're parsing. Try to locate it
                For Search = LBound(Entities) To UBound(Entities)
            
                    ' Assume this is a match
                    bFound = True
            
                    For lngChar = 1 To UBound(Entities(Search).lngName)
                    
                        If Entities(Search).lngName(lngChar) <> intData(Tell + lngChar) Then
                                                    
                            ' End the search
                            bFound = False
                            
                            ' We can now defently assume this is NOT a match
                            Exit For
                        
                        End If
                        
                    Next
                    
                    ' If we have found the entity, exit for
                    If bFound Then
                    
                        ' Add this character
                        For lngAdd = 1 To UBound(Entities(Search).lngText)
                        
                            ' Set the character
                            intData(Temp) = Entities(Search).lngText(lngAdd)
                            
                            ' The character has been added
                            Temp = Temp + 1
                        
                        Next
                        
                        ' No need for more searching
                        Exit For
                
                    End If
            
                Next
                
                If Search <= UBound(Entities) Then
                    ' Set the new position
                    Tell = lngSemicolon
                Else
                    ' No entity with this name - process as normal
                    GoTo Normal
                End If
            
            End If
            
        Else

' The normal way
Normal:
        
            ' Set the value
            If Temp <> Tell Then
                intData(Temp) = intData(Tell)
            End If
        
            ' The character is to be added
            Temp = Temp + 1
        
        End If
        
    Next
    
    ' Return string
    ProcessSegment = GetString(intData, 1, Temp)

End Function

Friend Function GetString(intChar() As Integer, lngStart As Long, lngEnd As Long) As String

    ' Don't continue if we cannot create the string
    If lngEnd - lngStart > 0 Then
         
         ' Allocate the space holding the string
        GetString = SysAllocStringByteLen(0, lngEnd - lngStart)
        
        ' Then set the string
        CopyMemory ByVal StrPtr(GetString), ByVal (VarPtr(intChar(1)) + ((lngStart - 1) * 2)), (lngEnd - lngStart) * 2
    
    End If

End Function

Private Sub Class_Initialize()

    ' Load all enities
    LoadEntities ValidPath(App.Path) & "Entities.dat"

End Sub

Private Function ValidPath(Path As String) As String

    ' Add a slash if it dosen't exist
    ValidPath = Path & IIf(Right(Path, 1) = "\", "", "\")

End Function
