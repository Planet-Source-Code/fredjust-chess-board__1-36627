Attribute VB_Name = "mTreeViewInFile"
'fredjust
'contact@fredjust.com
'Active Visual Basic
'http://www.fredjust.com

Option Explicit

Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpString As String, _
        ByVal lpFileName As String) As Long

'===========================================================================
'   SAVE A TREEVIEW IN A INI FILE
'===========================================================================
Public Sub tvSaveToFile(TV As TreeView, ByVal FileName As String, _
    Optional SaveNodeStyle As Boolean = False, _
    Optional SaveNodeImage As Boolean = False)

    Dim Neu As node
    Dim Tempo As String
    Dim ImageIndex As Long

    On Error Resume Next

    Kill FileName

    WritePrivateProfileString "FileInfo", "Type", "TVW FILE", FileName
    WritePrivateProfileString "FileInfo", "Version", "1.0", FileName

    With TV
        WritePrivateProfileString "TreeView", "Appearance", .Appearance, FileName
        WritePrivateProfileString "TreeView", "BorderStyle", .BorderStyle, FileName
        WritePrivateProfileString "TreeView", "Checkboxes", BoolToStr(.Checkboxes), FileName
        WritePrivateProfileString "TreeView", "FullRowSelect", BoolToStr(.FullRowSelect), FileName
        WritePrivateProfileString "TreeView", "HideSelection", BoolToStr(.HideSelection), FileName
        WritePrivateProfileString "TreeView", "ImageList", .ImageList.Name, FileName
        WritePrivateProfileString "TreeView", "Indentation", .Indentation, FileName
        WritePrivateProfileString "TreeView", "LabelEdit", .LabelEdit, FileName
        WritePrivateProfileString "TreeView", "LineStyle", .LineStyle, FileName
        WritePrivateProfileString "TreeView", "PathSeparator", .PathSeparator, FileName
        WritePrivateProfileString "TreeView", "Sorted", BoolToStr(.Sorted), FileName
        WritePrivateProfileString "TreeView", "Style", .Style, FileName
        WritePrivateProfileString "TreeView", "ToolTipText", .ToolTipText, FileName

        WritePrivateProfileString "Nodes", "Count", CStr(.Nodes.Count), FileName
    End With

    For Each Neu In TV.Nodes

        With Neu
            ' 0 TEXT
            Tempo = .Text

            '1 PARENT.INDEX
            If Not .Parent Is Nothing Then
                Tempo = Tempo & "|" & .Parent.Index
            Else
                Tempo = Tempo & "|" & "ROOT"
            End If

            '2 KEY 3 TAG 4 SORTED
            Tempo = Tempo & "|" & .Key & "|" & .Tag & "|" & BoolToStr(.Sorted)
            'Tempo = Tempo & "|" & "" & "|" & .Tag & "|" & BoolToStr(.Sorted)

            WritePrivateProfileString "Nodes", "Node" & .Index, Tempo, FileName

            If SaveNodeImage Then
                Tempo = TV.ImageList.ListImages(.Image).Index
                Tempo = Tempo & "|"
                Tempo = Tempo & TV.ImageList.ListImages(.SelectedImage).Index
                Tempo = Tempo & "|"
                Tempo = Tempo & TV.ImageList.ListImages(.ExpandedImage).Index

                WritePrivateProfileString "Nodes", "Image" & .Index, Tempo, FileName
            End If

            If SaveNodeStyle Then
                Tempo = .ForeColor & "|" & .BackColor & "|" & BoolToStr(.Bold) & "|" & BoolToStr(.Checked) & "|" & BoolToStr(.Expanded)
                WritePrivateProfileString "Nodes", "Style" & .Index, Tempo, FileName
            End If

        End With

    Next

End Sub


'===========================================================================
'   SAVE A TREEVIEW IN A INI FILE
'===========================================================================
Public Function tvLoadFromFile(TV As TreeView, ByVal FileName As String, _
    Optional LoadTvStyle As Boolean = True, _
    Optional LoadNodeStyle As Boolean = False, _
    Optional LoadNodeImage As Boolean = False) As Long

    Dim Neu As node
    Dim Tempo As String
    Dim NodesCount As Long
    Dim Champs
    Dim i As Long

    On Error Resume Next

    If ReadIniFile(FileName, "FileInfo", "Type", "") <> "TVW FILE" Then
        tvLoadFromFile = -1
        Exit Function
    End If

    If LoadTvStyle Then

        With TV
            .Nodes.Clear
            .Appearance = ReadIniFile(FileName, "TreeView", "Appearance", .Appearance)
            .BorderStyle = ReadIniFile(FileName, "TreeView", "BorderStyle", .BorderStyle)
            .Checkboxes = ReadIniFile(FileName, "TreeView", "Checkboxes", BoolToStr(.Checkboxes))
            .FullRowSelect = ReadIniFile(FileName, "TreeView", "FullRowSelect", BoolToStr(.FullRowSelect))
            .HideSelection = ReadIniFile(FileName, "TreeView", "HideSelection", BoolToStr(.HideSelection))
            .Indentation = ReadIniFile(FileName, "TreeView", "Indentation", .Indentation)
            .LabelEdit = ReadIniFile(FileName, "TreeView", "LabelEdit", .LabelEdit)
            .LineStyle = ReadIniFile(FileName, "TreeView", "LineStyle", .LineStyle)
            .PathSeparator = ReadIniFile(FileName, "TreeView", "PathSeparator", .PathSeparator)
            .Sorted = ReadIniFile(FileName, "TreeView", "Sorted", BoolToStr(.Sorted))
            .Style = ReadIniFile(FileName, "TreeView", "Style", .Style)
            .ToolTipText = ReadIniFile(FileName, "TreeView", "ToolTipText", .ToolTipText)
        End With

    End If

    NodesCount = ReadIniFile(FileName, "Nodes", "Count")

    For i = 1 To NodesCount

        Tempo = ReadIniFile(FileName, "Nodes", "Node" & i)
        Champs = Split(Tempo, "|")

        If Champs(1) = "ROOT" Then
            Set Neu = TV.Nodes.Add(, , Champs(2), Champs(0))
        Else
            Set Neu = TV.Nodes.Add(CLng(Champs(1)), tvwChild, Champs(2), Champs(0))
        End If

        With Neu
            .Tag = Champs(3)
            .Sorted = CBool(Champs(4))

            If LoadNodeImage Then
                Tempo = ReadIniFile(FileName, "Nodes", "Image" & i)
                Champs = Split(Tempo, "|")
                .Image = CLng(Champs(0))
                .SelectedImage = CLng(Champs(1))
                .ExpandedImage = CLng(Champs(2))
            End If

            If LoadNodeStyle Then
                Tempo = ReadIniFile(FileName, "Nodes", "Style" & i)
                Champs = Split(Tempo, "|")
                .ForeColor = Champs(0)
                .BackColor = Champs(1)
                .Bold = Champs(2)
                .Checked = Champs(3)
                .Expanded = Champs(4)
            End If

        End With

    Next

    tvLoadFromFile = Err.Number
End Function




'===========================================================================
'
'===========================================================================
Private Function BoolToStr(blnValue As Boolean) As String
    If blnValue Then
        BoolToStr = "1"
    Else
        BoolToStr = "0"
    End If
End Function


'===============================================================================
'
'===============================================================================
Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, _
        ByVal strKey As String, Optional ByVal strDefault As String) As String

    Dim szBuffer As String
    Dim iLen As Integer

    szBuffer = String(1024, Chr(0))
    iLen = GetPrivateProfileString(strSection, strKey, strDefault, szBuffer, Len(szBuffer), strIniFile)
    ReadIniFile = Left$(szBuffer, iLen)

End Function


