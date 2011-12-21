VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "Example"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4740
   ScaleHeight     =   4725
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin pXmlExample.Downloader Downloader1 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   4080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
   End
   Begin VB.ListBox lstCustomers 
      Height          =   3180
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label lblCustomers 
      Caption         =   "Customers:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private currentURL As String
Private parser As clsParser

Private Sub Form_Load()

    currentURL = "http://comphenix.net/dev/dtd/data/company.xml"
    Set parser = CreateLSParser(0, "")

End Sub

Private Sub cmdUpdate_Click()

    ' Download the file asynchronously
    Downloader1.Start currentURL
    
End Sub

Private Sub Downloader1_Complete(URL As String, Data As String, Key As String)

    ' In case you want to handle multiple data sources
    If currentURL = URL Then
        ProcessData StrConv(Data, vbUnicode)
    End If

End Sub

Private Sub ProcessData(sData As String)

    Dim document As clsDocument
    Dim employeeNode As clsElement
    Dim nameNodes As clsNodeList
    
    ' Parse XML
    Set document = parser.ParseText(sData)
    
    lstCustomers.Clear
    
    For Each employeeNode In document.GetElementsByTagName("employee")
    
        Set nameNodes = employeeNode.GetElementsByTagName("name")
        
        If nameNodes.Lenght > 0 Then
            lstCustomers.AddItem nameNodes.Item(0).TextContent
        End If
    Next

End Sub

