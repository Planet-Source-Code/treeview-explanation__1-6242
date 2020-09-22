VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TreeView Demo"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   405
      Left            =   3375
      TabIndex        =   3
      Top             =   3240
      Width           =   2400
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   855
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":030A
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   1815
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0362
      Top             =   120
      Width           =   2655
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3270
      Top             =   3990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":045D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0779
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A95
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8281
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   3
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DisplayContacts(Names As String)
    Dim tempNode As Node
    Dim nCounter As Integer
    Dim sTemp As String

On Error GoTo ErrorHandler
    
    tv1.Nodes.Clear 'clear tv1's of any previous nodes
    
    'add root node
    Set tempNode = tv1.Nodes.Add(, , "R", Names, 2)
    
    'add child nodes
    Set tempNode = tv1.Nodes.Add("R", tvwChild, "W", "Work Contacts", 3)
    Set tempNode = tv1.Nodes.Add("R", tvwChild, "H", "Home Contacts", 3)
    Set tempNode = tv1.Nodes.Add("R", tvwChild, "M", "Miscillaneous contacts", 3)
    
    tempNode.EnsureVisible 'this makes sure that the 3 nodes
                            'created become instantly visible
    
    'open file with names in
    Open App.path & "\Contacts.txt" For Input As #1
    
    Dim NoOfNames As String
    Dim NoOfNames1 As String
    Dim NoOfNames2 As String
    Line Input #1, NoOfNames
    Line Input #1, NoOfNames1
    Line Input #1, NoOfNames2
    
    'Add work contacts
    For nCounter = 1 To NoOfNames
        Line Input #1, sTemp
        Set tempNode = tv1.Nodes.Add("W", tvwChild, "W" & nCounter, sTemp, 1)
    Next nCounter
   ' tempNode.EnsureVisible
    
    'Add Home contacts
    For nCounter = 1 To NoOfNames1
        Line Input #1, sTemp
        Set tempNode = tv1.Nodes.Add("H", tvwChild, "H" & nCounter, sTemp, 1)
    Next nCounter
    'tempNode.EnsureVisible

    'Add Misc contacts
    For nCounter = 1 To NoOfNames2
        Line Input #1, sTemp
        Set tempNode = tv1.Nodes.Add("M", tvwChild, "M" & nCounter, sTemp, 1)
    Next nCounter
    'tempNode.EnsureVisible
    
    'close file
    Close #1
    
    'set treeview control on top
    tv1.ZOrder 0

    Exit Sub 'this stops the program from running the error
            'handler bellow
ErrorHandler:
    MsgBox "An error has occured", vbCritical, "TreeView Demo"
    End
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    'call DisplayContacts and set root nodes text as andys's
    'contacts
    DisplayContacts "Andy's Contacts"
End Sub


Private Sub tv1_NodeClick(ByVal Node As MSComctlLib.Node)
'I used the image index for this if statment
'what it means is that if the nodes image is 1
'then run the msgbox with the nodes text (the person name
'in this case) if not do nothing
'this image index relates to imagelist1 index 1 is an icon
'(a yellow page with writing on it), index 2 is a set of books
'and index 3 is a set of books with one open
If Node.Image = 1 Then
    MsgBox "You clicked " & Node.Text, vbInformation, "Treeview Demo"
Else
    DoEvents
End If
End Sub
