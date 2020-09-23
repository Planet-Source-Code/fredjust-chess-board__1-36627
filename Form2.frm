VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Tree"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   7740
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0712
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1536
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":235A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":317E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":3890
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":3FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":46B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4DC6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   847
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Load"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Back"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Next"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Up"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Down"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Collapse"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Expand"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6720
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":54D8
            Key             =   "pb"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":55EA
            Key             =   "fb"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":56FC
            Key             =   "tn"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":580E
            Key             =   "dn"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":5920
            Key             =   "pn"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":5A32
            Key             =   "cn"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":5B44
            Key             =   "rn"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":5C56
            Key             =   "fn"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":5D68
            Key             =   "cb"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":5E7A
            Key             =   "rb"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":5F8C
            Key             =   "db"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":609E
            Key             =   "tb"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":61B0
            Key             =   "start"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4260
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   452
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList2"
      Appearance      =   1
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FredJust
'fred.just@free.fr
'Active Visual Basic
'http://fred.just.free.fr/


Option Explicit

Dim StopSearch  As Boolean
Dim LesParent As Collection

Private Sub cmdStop_Click()
    StopSearch = True
End Sub

Private Sub Form_Load()
    DeleteNode
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    tv.Visible = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tv.Move 30, 30 + tb.Height, ScaleWidth - 60, ScaleHeight - 60 - tb.Height
    'pb.Move 30, ScaleHeight \ 2, ScaleWidth - 60
    'lblMove.Move ScaleWidth \ 2, ScaleHeight \ 2 - pb.Height * 2
    'cmdStop.Move ScaleWidth \ 2, ScaleHeight \ 2 + pb.Height * 2
End Sub

Private Sub lblMove_Click()

End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            DeleteNode
        Case 2
            MDIForm1.LoadEch
        Case 3
            MDIForm1.SaveEch
        Case 5
            If tv.SelectedItem.Key <> "start" Then
                MoveToPos tv.SelectedItem.Parent
            End If
        Case 6
            MoveToPos tv.SelectedItem.Child
        Case 7
            MoveToPos tv.SelectedItem.Previous
        Case 8
            MoveToPos tv.SelectedItem.Next
        Case 10
            If tv.SelectedItem.Key <> "start" Then
                tv.Nodes.Remove (tv.SelectedItem.Index)
            End If
        Case 11
            tv.Visible = False
            Expandnode tv.SelectedItem, False
            tv.Visible = True
        Case 12
            tv.Visible = False
            Expandnode tv.SelectedItem, True
            tv.Visible = True
    End Select
End Sub


Private Sub Expandnode(Neu As Node, Ouvert As Boolean)
Dim Fils As Node
Dim i As Integer
'On Error Resume Next
    Neu.Expanded = Ouvert
    
    If Neu.Children <> 0 Then
        Set Fils = Neu.Child
        For i = 1 To Neu.Children
            Expandnode Fils, Ouvert
            Set Fils = Fils.Next
            DoEvents
        Next
    End If
    
End Sub

Private Sub DeleteNode()
    With tv
        .Visible = False
        .Nodes.Clear
        .Visible = True
        .Nodes.Add , , " ", "Start", "start"
        MoveToPos tv.Nodes(1)
    End With
    LastKey = " "
End Sub



Private Sub tv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MDIForm1.mnuPopup
    End If
    MoveToPos tv.SelectedItem
End Sub

Private Sub MoveToPos(Node)
    Dim i As Long
    Dim Tempo As String

    On Error Resume Next
    
    If Node.Key = "" Then Exit Sub
    If Err.Number <> 0 Then Exit Sub
    
    Node.Selected = True

    Form1.StartPos
    For i = 0 To 64
        Plateau(i) = ""
    Next

    For i = 9 To 16
        Plateau(i) = "pn"
    Next

    For i = 49 To 56
        Plateau(i) = "pb"
    Next

    Plateau(57) = "tb"
    Plateau(64) = "tb"

    Plateau(1) = "tn"
    Plateau(8) = "tn"

    Plateau(58) = "cb"
    Plateau(63) = "cb"

    Plateau(2) = "cn"
    Plateau(7) = "cn"

    Plateau(59) = "fb"
    Plateau(62) = "fb"

    Plateau(3) = "fn"
    Plateau(6) = "fn"

    Plateau(60) = "db"
    Plateau(4) = "dn"

    Plateau(61) = "rb"
    Plateau(5) = "rn"

    Moves = Split(Node.FullPath, "\")
    Tempo = Trim(Node.Key)

    For i = 1 To UBound(Moves)
        'MakeMove Moves(i)
        MakeMove Mid(Tempo, (i - 1) * 4 + 1, 4)
    Next

    Form1.Dessineplateau

    LastKey = Node.Key
    
    WhiteMove = ((Len(Node.Key) - 1) Mod 8) = 0
    BlackMove = Not WhiteMove

End Sub



'==================================================================================
'
'==================================================================================
Private Sub MakeMove(ByVal aMove As String)
    CaseFrom = Mid(aMove, 1, 2)
    CaseTo = Mid(aMove, 3, 2)
    
    If Plateau(Numero(CaseFrom)) = "rb" And CaseFrom = "e1" And CaseTo = "g1" Then
        Plateau(Numero("h1")) = ""
        Plateau(Numero("f1")) = "tb"
    End If
    
    If Plateau(Numero(CaseFrom)) = "rb" And CaseFrom = "e1" And CaseTo = "c1" Then
        Plateau(Numero("a1")) = ""
        Plateau(Numero("d1")) = "tb"
    End If
    
    If Plateau(Numero(CaseFrom)) = "rn" And CaseFrom = "e8" And CaseTo = "g8" Then
        Plateau(Numero("h8")) = ""
        Plateau(Numero("f8")) = "tn"
    End If
    
    If Plateau(Numero(CaseFrom)) = "rn" And CaseFrom = "e8" And CaseTo = "c8" Then
        Plateau(Numero("a8")) = ""
        Plateau(Numero("d8")) = "tn"
    End If
    
    Plateau(Numero(CaseTo)) = Plateau(Numero(CaseFrom))
    Plateau(Numero(CaseFrom)) = ""
End Sub


'==================================================================================
'
'==================================================================================
Public Function Numero(coord As String)
    Dim l As Integer, c As Integer

    l = Asc(Mid(coord, 1, 1)) - 96
    c = Mid(coord, 2, 1)

    Numero = (8 - c) * 8 + l

End Function



'==================================================================================
'
'==================================================================================
Public Sub OpenLarsenBook()
Dim ligne As String
Dim i As Long
Dim Coup As String

'Dim l As Long

On Error Resume Next
    Set ts = FSO.OpenTextFile(App.Path & "\LarsenVB.opn")
    
    tv.Visible = False
    tv.Nodes.Clear
    tv.Nodes.Add , , " ", "Start", "start"
'    pb.Max = 3140
   ' l = 0
    While Not ts.AtEndOfStream And Not StopSearch
        ligne = Trim(ts.ReadLine)
        LastKey = " "
        For i = 1 To Len(ligne) Step 4
            Coup = Mid(ligne, i, 4)
            
            'tv.Nodes.Add LastKey, tvwChild,  LastKey & Coup, Mid(Coup, 1, 2) & "-" & Mid(Coup, 3, 2)
            allMoves.Add Mid(Coup, 1, 2) & "-" & Mid(Coup, 3, 2), LastKey & Coup
            
            If Len(LastKey) = 1 Then tv.Nodes.Add " ", tvwChild, Coup, Mid(Coup, 1, 2) & "-" & Mid(Coup, 3, 2)
            LastKey = LastKey & Coup
        Next
        'lblMove.Caption = "Loading ... " & Mid(ligne, 1, 4)
  '      l = l + 1
 '       pb.Value = l
        DoEvents
    Wend
    ts.Close
    tv.Visible = True
    
End Sub




Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
Dim Neu As Node
Dim element

On Error Resume Next

    For Each element In LesParent
        tv.Nodes(element).ForeColor = 0
    Next

    Set LesParent = New Collection

    Set Neu = Node
    While Not Neu.Parent Is Nothing
        Set Neu = Neu.Parent
        Neu.ForeColor = RGB(0, 0, 255)
        LesParent.Add Neu.Index
    Wend
End Sub
