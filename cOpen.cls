VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260



Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Private Type tCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type



Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As tCHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long
Private Declare Sub CopyMemoryStr Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "KERNEL32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private OFName As OPENFILENAME

Public MehWnd As Long

'===============================================================================
'
'===============================================================================
Public Function ShowColor() As Long
    Dim cc As tCHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long
    Dim CustomColors

    'set the structure size
    cc.lStructSize = Len(cc)
    'Set the owner
    cc.hWndOwner = MehWnd
    'set the application's instance
    cc.hInstance = App.hInstance
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    cc.flags = 0

    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
    
End Function

'===============================================================================
'
'===============================================================================
Public Function ShowOpen(Optional ByVal strFilter As String = "", _
        Optional ByVal strTitle As String = "Ouvrir", _
        Optional ByVal strIniPath As String) As String
        
    'Set the structure size
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hWndOwner = MehWnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Set the filet
    If strFilter = "" Then strFilter = "All Files (*.*)|*.*|"
    OFName.lpstrFilter = Replace(strFilter, "|", Chr(0))
    
    'Create a buffer
    OFName.lpstrFile = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    If strIniPath = "" Then strIniPath = App.Path & "\"
    OFName.lpstrInitialDir = strIniPath
    'Set the dialog title
    OFName.lpstrTitle = strTitle
    'no extra flags
    OFName.flags = 0
    
    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
    
End Function


'===============================================================================
'
'===============================================================================
Public Function ShowSave(Optional ByVal strFilter As String = "*.*", _
        Optional ByVal strTitle As String = "Sauvegarder", _
        Optional ByVal strIniPath As String) As String
        
Dim aExt  As String
Dim Ext

    'Set the structure size
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hWndOwner = MehWnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Set the filet
    If strFilter = "*.*" Then strFilter = "All Files (*.*)|*.*|"
    Ext = Split(strFilter, "|")
    OFName.lpstrFilter = Replace(strFilter, "|", Chr(0))
    'Create a buffer
    OFName.lpstrFile = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    If strIniPath = "" Then strIniPath = App.Path & "\"
    OFName.lpstrInitialDir = strIniPath
    'Set the dialog title
    OFName.lpstrTitle = strTitle
    'no extra flags
    OFName.flags = 0
    
    'Show the 'Save File'-dialog
    If GetSaveFileName(OFName) Then
        
        ShowSave = Trim$(OFName.lpstrFile)
        ShowSave = Replace(ShowSave, Chr(0), "")
        aExt = Ext(OFName.nFilterIndex * 2 - 1)
        aExt = Mid$(aExt, InStr(2, aExt, ".") + 1)
        If Right$(ShowSave, Len(aExt)) <> aExt Then
            ShowSave = ShowSave & "." & aExt
        End If
    Else
        ShowSave = ""
    End If
    
End Function

'===============================================================================
'
'===============================================================================
Public Function BrowseFolder(Optional strTitle As String = "Choose a folder") As String
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hWndOwner = MehWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat(strTitle, "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    
    BrowseFolder = sPath
    
End Function

