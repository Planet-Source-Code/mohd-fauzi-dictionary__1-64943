VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   3795
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   10635
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   3975
      Begin MSComctlLib.ListView lstView 
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4471
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No."
            Object.Width           =   1129
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Word"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Search :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   5400
         Picture         =   "frmMain.frx":1762
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Delete"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Height          =   615
         Left            =   4800
         Picture         =   "frmMain.frx":1B24
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Update"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdNew 
         Height          =   615
         Left            =   3600
         Picture         =   "frmMain.frx":1CB0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "New"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox gintIdItem 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdSave 
         Height          =   615
         Left            =   4200
         Picture         =   "frmMain.frx":1ECF
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Save"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtMeaning 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox txtWord 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Meaning"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Word"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
    Dim strDelete As String
    
    strDelete = "Delete from EngToMalay Where Id = " & gintIdItem.Text & ""
    gAdoConn.Execute strDelete
    PopData (strTextSearch)
    txtWord.Text = ""
    txtMeaning.Text = ""
    
End Sub

Private Sub cmdNew_Click()
    txtWord.Text = ""
    txtMeaning.Text = ""
End Sub

Private Sub cmdSave_Click()
Dim strSQL As String
Dim rs As ADODB.Recordset

If txtWord.Text = "" Then
    MsgBox "Enter the word. ", vbExclamation, "Alert"
    Exit Sub
End If
If txtMeaning.Text = "" Then
    MsgBox "Enter the meaning of word.", vbExclamation, "Alert"
    Exit Sub
End If

strSQL = "Insert into EngToMalay(Istilah,IstilahDesc)Values('" & SQLSafe(txtWord.Text) & "','" & _
        SQLSafe(txtMeaning.Text) & "')"
gAdoConn.Execute strSQL

PopData (strTextSearch)
txtWord.Text = ""
txtMeaning.Text = ""



End Sub

Private Sub cmdUpdate_Click()
Dim strUpdate As String

    strUpdate = "Update EngToMalay Set Istilah = '" & SQLSafe(txtWord) & "'," & _
    "IstilahDesc = '" & SQLSafe(txtMeaning) & "' Where Id = " & gintIdItem & ""
    gAdoConn.Execute strUpdate
    PopData (strTextSearch)
    txtWord.Text = ""
    txtMeaning.Text = ""
    
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title
    InitConnection
    PopData (strTextSearch)
End Sub





Private Sub lstView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim intSelItem As Integer
    
    intSelItem = Item
    
    txtWord.Text = lstView.ListItems(intSelItem).ListSubItems(1).Text
    txtMeaning.Text = lstView.ListItems(intSelItem).ListSubItems(2).Text
    gintIdItem = lstView.ListItems(intSelItem).ListSubItems(3).Text
    
   

End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub InitConnection()
    Dim conDBString As String

    conDBString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Dictionary.mdb"
    
    Set gAdoConn = New ADODB.Connection
        gAdoConn.ConnectionString = conDBString
        gAdoConn.Open

End Sub

Private Sub PopData(strTextSearch As String)
    
    Dim lstX As ListItem
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    Dim intCounter As Integer
    If strTextSearch = "" Then
        strSQL = "select * from EngToMalay Order by Istilah ASC"
    Else
        strSQL = "Select * from EngToMalay Istilah " & _
        "where Istilah like '%" & strTextSearch & "%' order by Istilah asc"
        
    End If
    
    
    Set rs = New ADODB.Recordset
        rs.Open strSQL, gAdoConn, 3, 1
    
    lstView.ListItems.Clear
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            intCounter = 1
            While Not .EOF
            Set lstX = lstView.ListItems.Add(, , intCounter)
                lstX.ListSubItems.Add = Trim(!Istilah)
                lstX.ListSubItems.Add = Trim(!IstilahDesc)
                lstX.ListSubItems.Add = Trim(!Id)
            intCounter = intCounter + 1
            .MoveNext
            Wend
        End If
    End With
End Sub



Private Sub txtSearch_Change()
    PopData (txtSearch.Text)

End Sub


