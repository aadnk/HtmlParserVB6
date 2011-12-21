VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

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

' Private variables
Private Elements() As HTMLElement
Private StaticEntities() As Entity
Private AddedEntites() As Entity

' A lookup table used in the process of assembling the document back together
Private eLookupTable() As Integer

' The object that created this document
Private Parser As clsParser

' The element count
Public ElementCount As Long
Attribute ElementCount.VB_VarDescription = "The amout of elements this document contains."

' Other variables
Public InputEncoding As String
Attribute InputEncoding.VB_VarDescription = "An attribute specifying the encoding used for this document at the time of the parsing."
Public XMLEncoding As String
Attribute XMLEncoding.VB_VarDescription = "An attribute specifying, as part of the XML declaration, the encoding of this document."
Public XMLVersion As String
Attribute XMLVersion.VB_VarDescription = "An attribute specifying the version number of this document."
Public Implementation As Object
Attribute Implementation.VB_VarDescription = "The DOMImplementation object that handles this document."
Public DocumentURI As String
Attribute DocumentURI.VB_VarDescription = "The location of the document or null if undefined."

' The class holding information about the doctype
Private pDocType As New clsDocumentType

' Used in optmizing the allocation of array elements - only allocate different sections
Private ArraySize As Long

Private Property Get ElementTitle() As clsElement

    ' Ignore all errors
    On Error Resume Next

    ' This simple property will return the title element, if found
    Set ElementTitle = DocumentElement.GetElementsByTagName("html").Item(0). _
     GetElementsByTagName("head").Item(0).GetElementsByTagName("title").Item(0)
     
End Property

Public Property Get Title() As String

    On Error Resume Next
    
    ' Return the underlying text-node of this element
    Title = ElementTitle.InnerHTML
 
End Property

Public Property Let Title(ByVal sNewTitle As String)

    On Error Resume Next
    
    ' Set the text-node of the title-node
    ElementTitle.TextContent = sNewTitle
    
End Property

Public Property Get DocType() As clsDocumentType
Attribute DocType.VB_Description = "The Document Type Declaration associated with this document."

    ' Return the internal doctype
    Set DocType = pDocType

End Property

Public Property Let DocType(ByVal NewDocType As clsDocumentType)

    On Error Resume Next
    Dim Tell As Long, Temp As Long

    ' Firstly, set the doctype
    Set pDocType = NewDocType
    
    ' Then, look for the doctype
    For Tell = 0 To ElementCount
    
        ' See if this is a document type
        If Elements(Tell).NodeType = DOMDocumentType Then
        
            ' Set its property
            Elements(Tell).TagName = NewDocType.Name
            Elements(Tell).Properties(0).Name = NewDocType.PublicId
            Elements(Tell).Properties(1).Name = NewDocType.SystemId
            
            For Temp = 0 To NewDocType.Notations.Lenght
            
                ' Reallocate array
                ReDim Preserve Elements(Tell).Properties(Temp + 2)
                
                ' Set the data of this element
                Elements(Tell).Properties(Temp + 2).Name = NewDocType.Notations.Item(Temp)
                
            Next
            
            ' We are finished
            Exit Property
        
        End If
    
    Next

End Property

Public Function CreateNodeList(lpArray() As Long) As clsNodeList
Attribute CreateNodeList.VB_Description = "Creates a node list from the given array of indexes."

    ' Create the new array
    Set CreateNodeList = New clsNodeList
    
    ' Then, initialize it
    CreateNodeList.Initialize lpArray, Me

End Function

Public Function CreateProcessingInstruction(Target As String, Data As String) As clsElement
Attribute CreateProcessingInstruction.VB_Description = "Creates a ProcessingInstruction node given the specified name and data strings."

    ' Create a new text node
    Set CreateProcessingInstruction = GetElementByIndex(AppendElement(Target, 0, -1, DOMText, True, Data))

End Function

Public Function CreateElement(TagName As String, Optional EndTag As Boolean = True) As clsElement
Attribute CreateElement.VB_Description = "Creates an element of the type specified."

    ' Create a new text node
    Set CreateElement = GetElementByIndex(AppendElement(TagName, 0, -1, DOMElement, EndTag))

End Function

Public Function CreateComment(Data As String) As clsElement
Attribute CreateComment.VB_Description = "Creates a Comment node given the specified string."

    ' Create a new text node
    Set CreateComment = GetElementByIndex(AppendElement("#comment", 0, -1, DOMComment, True, Data))

End Function

Public Function CreateAttribute(Name As String) As clsAttr
Attribute CreateAttribute.VB_Description = "Creates an Attr of the given name."

    ' Create a new text node
    Set CreateAttribute = New clsAttr
    
    ' Set information
    CreateAttribute.Name = Name

End Function

Public Function CreateTextNode(sText As String) As clsElement
Attribute CreateTextNode.VB_Description = "Creates a Text node given the specified string."

    ' Create a new text node
    Set CreateTextNode = GetElementByIndex(AppendElement("#text", 0, -1, DOMText, True, sText))

End Function

Public Function CreateCDATASection(ByVal Data As String) As clsElement
Attribute CreateCDATASection.VB_Description = "Creates a CDATASection node whose value is the specified string."
    
    ' Simply create a new CDATA-node
    Set CreateCDATASection = GetElementByIndex(AppendElement("#cdata-section", 0, -1, DOMCDataSection, True, Data))

End Function

Public Function GetElementByIndex(ByVal lngIndex As Long) As clsElement
Attribute GetElementByIndex.VB_Description = "Returns the element given by the index."

    ' The index must of course be valid
    If lngIndex < 0 Or lngIndex > ElementCount Then
        Exit Function
    End If

    ' Create a node-class
    Set GetElementByIndex = New clsElement
    
    ' Initialize it
    GetElementByIndex.Initialize Elements, Me
    GetElementByIndex.NodeIndex = lngIndex

End Function

Public Function GetElementsByTagName(ByVal TagName As String, Optional lngStart As Long) As clsNodeList
Attribute GetElementsByTagName.VB_Description = "Returns a NodeList of all the Elements in document order with a given tag name and are contained in the document."

    Dim lngElements() As Long, Tell As Long, cElements As New Collection
    
    ' Search in all elements below
    SearchInElements lngStart, cElements, TagName
    
    ' Then allocate the long list
    ReDim lngElements(cElements.Count - 1)
    
    ' And fill it
    For Tell = 1 To cElements.Count
    
        ' Set the element
        lngElements(Tell - 1) = cElements(Tell)
    
    Next
    
    ' Create the node list
    Set GetElementsByTagName = CreateNodeList(lngElements)

End Function

Friend Sub SearchInElements(searchFrom As Long, cElements As Collection, TagName As String)

    Dim Tell As Long, lngIndex As Long

    ' Check all sub-elements
    For Tell = 0 To Elements(searchFrom).ChildCount - 1
    
        ' Get the index of this element
        lngIndex = Elements(searchFrom).Children(Tell)
    
        ' Firstly, check its name
        If Elements(lngIndex).TagName = TagName Then
        
            ' Yup, add it
            cElements.Add lngIndex
        
        End If
        
        ' Then, search in its sub-elements
        SearchInElements lngIndex, cElements, TagName
    
    Next
    
End Sub

Public Function GetElementByID(ByVal lngID As Long) As clsElement
Attribute GetElementByID.VB_Description = "Returns the Element that has an ID attribute with the given value. If no such element exists, this returns null. "

    Dim Tell As Long, Temp As Long

    ' Try to locate the element
    For Tell = 0 To ElementCount
    
        ' Find the ID-property
        For Temp = 0 To Elements(Tell).PropertyCount - 1
    
            If Elements(Tell).Properties(Temp).ID Then
        
                ' Then, if the ID corresponds to this element, return it
                If Elements(Tell).Properties(Temp).Value = lngID Then
                
                    ' Return object
                    Set GetElementByID = GetElementByIndex(Tell)
                    
                    ' We are finished
                    Exit Function
                
                End If
        
            End If
            
        Next
    
    Next

End Function

Public Property Get DocumentElement() As clsElement
Attribute DocumentElement.VB_Description = "This is a convenience  attribute that allows direct access to the child node that is the document element of the document."
Attribute DocumentElement.VB_UserMemId = 0

    ' The first node is always the document node
    Set DocumentElement = GetElementByIndex(0)

End Property

Public Function RenameNode(Node As Object, strName As String) As Object
Attribute RenameNode.VB_Description = "Rename an existing node of type DOMElement or DOMAttribute."

    If TypeOf Node Is clsElement Then
    
        ' This is a node, rename based on the index
        Elements(Node.NodeIndex).TagName = strName
        
        ' Return the node
        Set RenameNode = GetElementByIndex(Node.NodeIndex)
        
        ' Inform about the change
        InvokeHandlers Node.NodeIndex, Node_Renamed, Node, RenameNode
    
    ElseIf TypeOf Node Is clsAttr Then

        ' This is a attribute, rename based on the index
        Elements(Node.NodeIndex).Properties(Node.AttrIndex).Name = strName
        
        ' Return the attribute
        Set RenameNode = AttributeFromIndex(GetElementByIndex(Node.NodeIndex), Node.NodeIndex, Node.AttrIndex)
        
        ' Inform about the change
        InvokeHandlers Node.NodeIndex, Node_Renamed, Node, RenameNode

    End If

End Function

Friend Function CloneNode(ByVal lngIndex As Long, ByVal lngParent As Long, ByVal bCopyChildren As Boolean) As Long

    Dim Tell As Long

    ' Firstly, reallocate a new element
    AppendElement "", 0, 0, 0, False
    
    ' Then copy it
    LSet Elements(ElementCount) = Elements(lngIndex)
    
    ' Set the parent
    If lngParent < 0 Then
        
        ' We cannot add this element to any children list
        Elements(ElementCount).Parent = lngParent
        
    Else
    
        ' Add it to its parent
        AppendChild ElementCount, lngParent
    
    End If
    
    ' The element added
    CloneNode = ElementCount
    
    ' Add all children if specified
    If bCopyChildren Then
    
        ' Go through all children
        For Tell = 0 To Elements(lngIndex).ChildCount - 1
    
            ' Copy it to this node
            CloneNode Elements(lngIndex).Children(Tell), CloneNode, bCopyChildren
    
        Next
    
    End If
    
    ' Increse the element count
    ElementCount = ElementCount + 1

End Function

Friend Sub Initialize(lpElements() As HTMLElement, lpEntities() As Entity, lpLookupTable() As Integer, ParserClass As clsParser)

    ' Copy the reference of the arrays
    CopyMemory ByVal VarPtrArray(lpElements), ByVal VarPtrArray(Elements), 4

    ' Make a reference to the static entities and the lookup table
    CopyMemory ByVal VarPtrArray(StaticEntities), ByVal VarPtrArray(lpEntities), 4
    CopyMemory ByVal VarPtrArray(eLookupTable), ByVal VarPtrArray(lpLookupTable), 4
    
    ' And get the reference to the calling class
    Set Parser = ParserClass

End Sub

Friend Sub RemoveChildren(lngIndex As Long)

    Dim Tell As Long

    ' Firstly, go through children, removing their reference to this node
    For Tell = 0 To Elements(lngIndex).ChildCount - 1
    
        ' Remove the parent reference
        Elements(Elements(lngIndex).Children(Tell)).Parent = -1
    
    Next

    ' Secoundly, free array
    Erase Elements(lngIndex).Children
    
    ' Then, reset count
    Elements(lngIndex).ChildCount = 0

End Sub

Friend Sub ConvertString(sText As String, clsString As clsStringBuilder, Optional bAddQuotation As Boolean)

    On Error Resume Next
    Dim sData As String, sResult As String, intChar() As Integer, currChar As Integer
    Dim SafeArray As SAFEARRAY1D, Tell As Long, lngLast As Long, bAddedAprostophe As Boolean
    Dim bAprostophe As Boolean, bQuote As Boolean
        
    ' Let the array point to the string
    With SafeArray
        .cDims = 1
        .Bounds(0).lLbound = 1
        .Bounds(0).cElements = Len(sText)
        .pvData = StrPtr(sText)
    End With
    
    ' Use copy memory for this purpose
    CopyMemory ByVal VarPtrArray(intChar), VarPtr(SafeArray), 4

    If bAddQuotation Then
        
        ' We need to find out what to use
        bQuote = (InStr(1, sText, """") <> 0)
        bAprostophe = (InStr(1, sText, "'") <> 0)
        
        If bQuote And Not bAprostophe Then
            ' Use aprostophes as quotation tags
            clsString.Append "'"
            bQuote = False
            bAddedAprostophe = True
        Else
            ' If not, use quotation tags
            clsString.Append """"
            bAprostophe = False
        End If
        
    End If

    ' Let it begin at the first position
    lngLast = 1

    ' Go trough the entire array
    For Tell = 1 To UBound(intChar)
    
        ' Get the current char
        currChar = intChar(Tell)
    
        If currChar = LessThan Or currChar = GreaterThan Or currChar = Ampersand Or currChar = QuotationMark Or currChar = Apostrophe Or currChar > 127 Then
        
            ' Add the text before
            clsString.Append Mid(sText, lngLast, Tell - lngLast)
            
            ' To add a entity, we first need to check whether or not its already defined
            If eLookupTable(currChar) < 0 Then
            
                ' Add this special enitity
                clsString.Append "&#" & currChar & ";"
            
            Else
            
                ' Yup, we need to use the entity name
                clsString.Append "&" & Parser.GetString(StaticEntities(eLookupTable(currChar)).lngName, 1, UBound(StaticEntities(eLookupTable(currChar)).lngName) + 1) & ";"

            End If

            ' The last position is now this plus one
            lngLast = Tell + 1

        End If
    
    Next
    
    ' See if we need to add a string
    If lngLast < Len(sText) Then
        ' Add the string
        clsString.Append Mid(sText, lngLast)
    End If
    
    If bAddQuotation Then
        If bAddedAprostophe Then
            ' Use aprostophes as quotation tags
            clsString.Append "'"
        Else
            ' If not, use quotation tags
            clsString.Append """"
        End If
    End If
    
    ' Clear up
    ZeroMemory ByVal VarPtrArray(intChar), 4
    
End Sub

Friend Function IsValidName(sName As String) As Boolean

    Dim SafeArray As SAFEARRAY1D, Tell As Long, intChar() As Integer, currChar As Integer

    ' Let the array point to the string
    With SafeArray
        .cDims = 1
        .Bounds(0).lLbound = 1
        .Bounds(0).cElements = Len(sName)
        .pvData = StrPtr(sName)
    End With
    
    ' Use copy memory for this purpose
    CopyMemory ByVal VarPtrArray(intChar), VarPtr(SafeArray), 4

    ' Go trough all characters of this string
    For Tell = LBound(intChar) To UBound(intChar)

        ' Get the current character
        currChar = intChar(Tell)

        ' See if it is not a normal character
        If Not ((currChar >= SmallA And currChar <= SmallZ) Or (currChar >= LargeA And currChar <= LargeZ)) Then
        
            ' This is not a valid name
            Exit For
        
        End If

    Next

    ' See if we haven't ended
    If Tell > UBound(intChar) Then
    
        ' If not, this is a valid name
        IsValidName = True
        
    End If

    ' Clear up
    ZeroMemory ByVal VarPtrArray(intChar), 4

End Function

Friend Sub GetElementHTML(lngIndex As Long, clsString As clsStringBuilder)

    Dim Tell As Long, sTagName As String, lngNodeType As Long

    ' Get the tag name
    sTagName = Elements(lngIndex).TagName
    lngNodeType = Elements(lngIndex).NodeType

    Select Case lngNodeType
    Case DOMText, DOMComment, DOMDocument, DOMCDataSection ' If this is just a text alike tag, simply return its value
    
        ' The comment tag require tags
        If lngNodeType = DOMComment Then
            ' Add the start tag
            clsString.Append "<!--"
        ElseIf lngNodeType = DOMCDataSection Then
            clsString.Append "<![CDATA["
        End If
    
        ' Only append inner data
        If Elements(lngIndex).ChildCount = 0 Then
            
            ' Append value
            ConvertString Elements(lngIndex).Value, clsString
        
        Else
        
            For Tell = 0 To Elements(lngIndex).ChildCount - 1
                ' Add the childs inner HTML
                GetElementHTML Elements(lngIndex).Children(Tell), clsString
            Next
            
        End If
        
        ' The comment tag alo require end tags
        If lngNodeType = DOMComment Then
            clsString.Append "-->"
        ElseIf lngNodeType = DOMCDataSection Then
            clsString.Append "]]>"
        End If

    Case DOMElement, DOMDocument, DOMDocumentType, DOMProcessingInstruction
    
        ' Write the first tag
        Select Case lngNodeType
        Case DOMElement
            ' Just add the begin tag for an normal element
            clsString.Append "<" & sTagName
            
        Case DOMDocumentType
            ' Add the document-type tag
            clsString.Append "<!DOCTYPE " & sTagName
            
        Case DOMProcessingInstruction
            ' A processing instruction always begin with a question mark
            clsString.Append "<?" & sTagName
            
        End Select
        
        ' Then write each property
        For Tell = 0 To Elements(lngIndex).PropertyCount - 1
    
            ' Add the space
            clsString.Append " "
    
            ' Append the property name
            ConvertString Elements(lngIndex).Properties(Tell).Name, clsString, Not IsValidName(Elements(lngIndex).Properties(Tell).Name)
            
            If Len(Elements(lngIndex).Properties(Tell).Value) Then
                ' Add the equal sign
                clsString.Append "="
                
                ' Add the value
                ConvertString Elements(lngIndex).Properties(Tell).Value, clsString, True
            End If
            
        Next
        
        ' Insert the end tag
        If Elements(lngIndex).ChildCount = 0 Then
        
            ' With no children, add a "special" end tag
            If lngNodeType = DOMDocumentType Then
                ' A DOCTYPE dosen't need a slash at the end
                clsString.Append ">"
            ElseIf lngNodeType = DOMProcessingInstruction Then
                ' A processing instruction also ends with a question mark
                clsString.Append "?>"
            Else
                ' This is just a normal element
                clsString.Append " />"
            End If

        Else
        
            ' First, end the start tag
            clsString.Append ">"
        
            For Tell = 0 To Elements(lngIndex).ChildCount - 1
            
                ' Add the childs inner HTML
                GetElementHTML Elements(lngIndex).Children(Tell), clsString
            
            Next
            
            ' Add the end-tag if this is an element
            If lngNodeType = DOMElement Then
                clsString.Append "</" & sTagName & ">"
            End If
        
        End If

    Case DOMEntityReference
    
        ' Simply add the enitiy code
        clsString.Append "&" & sTagName & ";"

    End Select

End Sub

Friend Function AppendElement(strName As String, lngPosition As Long, lngParent As Long, lngType As NodeType, EndTag As Boolean, Optional Value As String) As Long

    ' Increse tag count
    ElementCount = ElementCount + 1

    ' Do we need to reallocate array?
    If ElementCount > ArraySize Then

        ' Increse array size
        ArraySize = (2 * ArraySize) + 50
        
        ' Reallocate elements
        ReDim Preserve Elements(ArraySize)
    
    End If
    
    ' Set the properies
    Elements(ElementCount).EndTag = EndTag
    Elements(ElementCount).TagName = strName
    Elements(ElementCount).Position = lngPosition
    Elements(ElementCount).Parent = lngParent
    Elements(ElementCount).Value = Value
    Elements(ElementCount).NodeType = lngType
    
    ' Add the tag to the parent
    AppendChild lngParent, ElementCount
    
    ' Return the index
    AppendElement = ElementCount

End Function

Friend Sub AppendChild(lngIndex As Long, lngChildrenIndex As Long, Optional lngChildStart As Long = -1)

    Dim newIndex As Long, newSize As Long, Tell As Long
    
    ' Don't continue if the indexes isn't valid
    If lngIndex < 0 Or lngIndex > ElementCount Or lngChildrenIndex < 0 Or lngChildrenIndex > ElementCount Then
        Exit Sub
    End If
    
    ' The new size of the children array
    newSize = Elements(lngIndex).ChildCount
    
    ' The array element to be created
    If lngChildStart < 0 Then
        newIndex = Elements(lngIndex).ChildCount
    Else
        newIndex = lngChildStart
    End If

    ' Reallocate the array
    ReDim Preserve Elements(lngIndex).Children(newSize)
    
    ' Move them upwards
    For Tell = newSize To newIndex + 1 Step -1
        Elements(lngIndex).Children(Tell) = Elements(lngIndex).Children(Tell - 1)
    Next
    
    ' Then, set the value
    Elements(lngIndex).Children(newIndex) = lngChildrenIndex
    
    ' Of course, we need to reference this childern index to its parent
    Elements(lngChildrenIndex).Parent = lngIndex
    
    ' And finally increse the children count
    Elements(lngIndex).ChildCount = newSize + 1
    
End Sub

Friend Sub SetParent(lngNodeIndex As Long, lngParent As Long)

    ' Set this element's parent
    Elements(lngNodeIndex).Parent = lngParent

End Sub

Friend Sub SetChildIndex(lngNodeIndex As Long, lngChildIndex As Long, lngValue As Long)

    ' This is just a 'hack' to get access to the children array of a certain element
    Elements(lngNodeIndex).Children(lngChildIndex) = lngValue

End Sub

Friend Sub RemoveElement(NodeIndex As Long)

    Dim Tell As Long
    
    If NodeIndex < 0 Then
        ' Firstly, we know that ElementCount is always one bigger than the actual array.
        ' We also need to take into account that one element is to be removed.
        NodeIndex = ElementCount - 2
    End If
    
    ' Move elements if necessary
    For Tell = NodeIndex To ElementCount - 3
    
        ' Copy the data of this element
        LSet Elements(Tell) = Elements(Tell + 1)
    
    Next

    ' Decrese tag count
    ElementCount = ElementCount - 1
    
    ' The class is now not really valid
    NodeIndex = -1

End Sub

Friend Function AppendAttribute(NodeIndex As Long, Optional Name As String, Optional Value, Optional IsID As Boolean) As Long

    ' Return the index of this attribute
    AppendAttribute = Elements(NodeIndex).PropertyCount

    ' Reallocate array
    ReDim Preserve Elements(NodeIndex).Properties(AppendAttribute)
    
    ' Set properties
    Elements(NodeIndex).Properties(AppendAttribute).Name = Name
    Elements(NodeIndex).Properties(AppendAttribute).Value = Value
    Elements(NodeIndex).Properties(AppendAttribute).ID = IsID
    
    ' Increse attribute count
    Elements(NodeIndex).PropertyCount = AppendAttribute + 1

End Function

Friend Function InvokeHandlers(NodeIndex As Long, lngOperation As OperationTypes, Node As Object, newly As Object)

    On Error Resume Next
    Dim Tell As Long
    
    ' Go through all user data elements
    For Tell = 0 To Elements(NodeIndex).UserDataCount - 1
    
        ' Call the handle prosedure
        Elements(NodeIndex).UserData(Tell).Handler.Handle lngOperation, _
         Elements(NodeIndex).UserData(Tell).Key, Elements(NodeIndex).UserData(Tell).Data, Node, newly
        
    Next

End Function

Friend Function AttributeFromIndex(OwnerElement As clsElement, NodeIndex As Long, AttrIndex As Long) As clsAttr

    ' Then, return the attribute
    Set AttributeFromIndex = New clsAttr
    
    ' Set properties
    Set AttributeFromIndex.OwnerElement = OwnerElement
    AttributeFromIndex.AttrIndex = AttrIndex
    AttributeFromIndex.NodeIndex = NodeIndex
    AttributeFromIndex.Value = Elements(NodeIndex).Properties(AttrIndex).Value
    AttributeFromIndex.Name = Elements(NodeIndex).Properties(AttrIndex).Name
    AttributeFromIndex.IsID = Elements(NodeIndex).Properties(AttrIndex).ID

End Function

' Will update different information regarding this document based on different tags
Friend Sub RefreshInfo()

    On Error Resume Next
    Dim Tell As Long, cList As Collection, Temp As Long
    
    ' Firstly, find the doctype element
    For Tell = 0 To ElementCount
    
        ' See if this is a doctype
        If Elements(Tell).NodeType = DOMDocumentType Then
        
            ' Create a fresh doctype
            Set pDocType = New clsDocumentType
        
            ' Then set the properties
            pDocType.Name = Elements(Tell).TagName
            pDocType.PublicId = Elements(Tell).Properties(0).Name
            pDocType.SystemId = Elements(Tell).Properties(1).Name
            
            ' Create the listr
            Set cList = New Collection
            
            ' Add all elements after PublicID and SystemId
            For Temp = 2 To Elements(Tell).PropertyCount - 1
                ' Add the element
                cList.Add Elements(Tell).Properties(Temp).Name
            Next
            
            ' Then, set it to doctype
            pDocType.Notations.Initialize cList
            
            ' We haven't got more to do witht the doctype
            Exit For
        
        End If
    
    Next
    
End Sub

Private Sub Class_Initialize()

    ' The amout of elements
    ElementCount = -1
    ArraySize = -1

End Sub

Private Sub Class_Terminate()

    ' Remove references of these arrays
    ZeroMemory ByVal VarPtrArray(StaticEntities), 4
    ZeroMemory ByVal VarPtrArray(eLookupTable), 4

End Sub
