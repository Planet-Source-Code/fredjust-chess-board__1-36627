VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Ouvertures"
   ClientHeight    =   10515
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10800
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuPopup 
      Caption         =   "Tools"
      Begin VB.Menu mnuComent 
         Caption         =   "Add Comments"
      End
      Begin VB.Menu mnuBold 
         Caption         =   "Bold"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    mnuPopup.Visible = False
    GameSize = 1
    Load Form1
    Form1.Hide
    Form1.ResizeForm GameSize
    Form1.Show
    Form2.Show
    Form1.Move 0, 0
    Form2.Move Form1.Width, 0, MDIForm1.ScaleWidth - Form1.Width, MDIForm1.ScaleHeight
End Sub

Private Sub mnuBold_Click()
    With Form2.tv
        .SelectedItem.Bold = Not .SelectedItem.Bold
    End With

End Sub

Private Sub mnuComent_Click()
    With Form2.tv
        .StartLabelEdit
    End With
End Sub

Public Sub LoadEch()
    Dim tempo As String
    tempo = cFile.ShowOpen("Chess files (*.ech)|*.ech|", , App.Path)
    If tempo <> "" Then
        Form2.tv.Visible = False
        tvLoadFromFile Form2.tv, tempo, , True, True
        Form2.tv.Visible = True
    End If
    
End Sub

Public Sub SaveEch()
Dim tempo As String
    tempo = cFile.ShowSave("Chess files (*.ech)|*.ech|", , App.Path)
    If tempo <> "" Then
        Form2.tv.Visible = False
        tvSaveToFile Form2.tv, tempo, True, True
        Form2.tv.Visible = True
    End If
End Sub


Private Sub mnuDel_Click()

End Sub
