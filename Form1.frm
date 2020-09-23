VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Echiquier"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   10260
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   3480
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Image Image3 
         Height          =   720
         Index           =   13
         Left            =   5040
         Picture         =   "Form1.frx":0000
         Top             =   1320
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   12
         Left            =   4080
         Picture         =   "Form1.frx":1CCA
         Top             =   1320
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   11
         Left            =   3120
         Picture         =   "Form1.frx":3994
         Top             =   1320
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   10
         Left            =   2160
         Picture         =   "Form1.frx":565E
         Top             =   1320
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   9
         Left            =   1200
         Picture         =   "Form1.frx":7328
         Top             =   1320
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   8
         Left            =   240
         Picture         =   "Form1.frx":8FF2
         Top             =   1320
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   7
         Left            =   5040
         Picture         =   "Form1.frx":ACBC
         Top             =   360
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   6
         Left            =   4080
         Picture         =   "Form1.frx":C986
         Top             =   360
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   5
         Left            =   3120
         Picture         =   "Form1.frx":E650
         Top             =   360
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   4
         Left            =   2160
         Picture         =   "Form1.frx":1031A
         Top             =   360
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   3
         Left            =   1200
         Picture         =   "Form1.frx":11FE4
         Top             =   360
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   720
         Index           =   2
         Left            =   240
         Picture         =   "Form1.frx":13CAE
         Top             =   360
         Width           =   720
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   1
         Left            =   2040
         Picture         =   "Form1.frx":15978
         Top             =   2400
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00004080&
      ForeColor       =   &H00004080&
      Height          =   8415
      Left            =   0
      ScaleHeight     =   8415
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.Image Image2 
         Height          =   720
         Index           =   0
         Left            =   720
         Top             =   7080
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   3
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   0
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   4
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   0
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   5
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   0
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   6
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   0
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   7
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   0
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   8
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   0
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   9
         Left            =   0
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   12
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   13
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   14
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   15
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   16
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   17
         Left            =   0
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   18
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   19
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   20
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   21
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   22
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   23
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   24
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   25
         Left            =   0
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   26
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   27
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   28
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   29
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   31
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   32
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   33
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   34
         Left            =   840
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   35
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   36
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   37
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   38
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   39
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   41
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   42
         Left            =   840
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   43
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   44
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   45
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   46
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   47
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   48
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   49
         Left            =   0
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   50
         Left            =   840
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   51
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   52
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   53
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   54
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   55
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   56
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   57
         Left            =   0
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   58
         Left            =   840
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   59
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   60
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   61
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   62
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   63
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   64
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   5880
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":15ACA
         Height          =   720
         Index           =   1
         Left            =   7440
         Picture         =   "Form1.frx":17248
         Tag             =   "tn"
         Top             =   240
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":18F12
         Height          =   720
         Index           =   9
         Left            =   8400
         Picture         =   "Form1.frx":1A30C
         Tag             =   "pn"
         Top             =   360
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":1BFD6
         Height          =   720
         Index           =   2
         Left            =   7440
         Picture         =   "Form1.frx":1DA24
         Tag             =   "cn"
         Top             =   1080
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":1F6EE
         Height          =   720
         Index           =   4
         Left            =   7320
         Picture         =   "Form1.frx":20FD4
         Tag             =   "dn"
         Top             =   3000
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":22C9E
         Height          =   720
         Index           =   5
         Left            =   7320
         Picture         =   "Form1.frx":2446C
         Tag             =   "rn"
         Top             =   3960
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":26136
         Height          =   720
         Index           =   6
         Left            =   7440
         Picture         =   "Form1.frx":27A1C
         Tag             =   "fn"
         Top             =   4800
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":296E6
         Height          =   720
         Index           =   7
         Left            =   7440
         Picture         =   "Form1.frx":2B134
         Tag             =   "cn"
         Top             =   5640
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":2CDFE
         Height          =   720
         Index           =   10
         Left            =   8400
         Picture         =   "Form1.frx":2E1F8
         Tag             =   "pn"
         Top             =   1200
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":2FEC2
         Height          =   720
         Index           =   11
         Left            =   8400
         MousePointer    =   4  'Icon
         Picture         =   "Form1.frx":312BC
         Tag             =   "pn"
         Top             =   2160
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":32F86
         Height          =   720
         Index           =   12
         Left            =   8400
         Picture         =   "Form1.frx":34380
         Tag             =   "pn"
         Top             =   3240
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":3604A
         Height          =   720
         Index           =   14
         Left            =   8400
         Picture         =   "Form1.frx":37444
         Tag             =   "pn"
         Top             =   4920
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":3910E
         Height          =   720
         Index           =   15
         Left            =   8400
         Picture         =   "Form1.frx":3A508
         Tag             =   "pn"
         Top             =   5640
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":3C1D2
         Height          =   720
         Index           =   16
         Left            =   8400
         Picture         =   "Form1.frx":3D5CC
         Tag             =   "pn"
         Top             =   6480
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":3F296
         Height          =   720
         Index           =   8
         Left            =   7440
         Picture         =   "Form1.frx":40A14
         Tag             =   "tn"
         Top             =   6480
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":426DE
         Height          =   720
         Index           =   13
         Left            =   8400
         MousePointer    =   4  'Icon
         Picture         =   "Form1.frx":43AD8
         Tag             =   "pn"
         Top             =   4200
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":457A2
         Height          =   720
         Index           =   3
         Left            =   7440
         Picture         =   "Form1.frx":47088
         Tag             =   "fn"
         Top             =   2040
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":48D52
         Height          =   720
         Index           =   60
         Left            =   2520
         Picture         =   "Form1.frx":4A39C
         Tag             =   "db"
         Top             =   5880
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":4C066
         Height          =   720
         Index           =   64
         Left            =   5880
         Picture         =   "Form1.frx":4D6B0
         Tag             =   "tb"
         Top             =   5880
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":4F37A
         Height          =   720
         Index           =   63
         Left            =   5040
         Picture         =   "Form1.frx":509C4
         Tag             =   "cb"
         Top             =   5880
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":5268E
         Height          =   720
         Index           =   62
         Left            =   4200
         Picture         =   "Form1.frx":53CD8
         Tag             =   "fb"
         Top             =   5880
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":559A2
         Height          =   720
         Index           =   61
         Left            =   3360
         Picture         =   "Form1.frx":56FEC
         Tag             =   "rb"
         Top             =   5880
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":58CB6
         Height          =   720
         Index           =   59
         Left            =   1680
         Picture         =   "Form1.frx":5A300
         Tag             =   "fb"
         Top             =   5880
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":5BFCA
         Height          =   720
         Index           =   58
         Left            =   840
         Picture         =   "Form1.frx":5D614
         Tag             =   "cb"
         Top             =   5880
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":5F2DE
         Height          =   720
         Index           =   56
         Left            =   5880
         Picture         =   "Form1.frx":60928
         Tag             =   "pb"
         Top             =   5040
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":625F2
         Height          =   720
         Index           =   55
         Left            =   5040
         Picture         =   "Form1.frx":63C3C
         Tag             =   "pb"
         Top             =   5040
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":65906
         Height          =   720
         Index           =   54
         Left            =   4200
         Picture         =   "Form1.frx":66F50
         Tag             =   "pb"
         Top             =   5040
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":68C1A
         Height          =   720
         Index           =   53
         Left            =   3360
         Picture         =   "Form1.frx":6A264
         Tag             =   "pb"
         Top             =   5040
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":6BF2E
         Height          =   720
         Index           =   52
         Left            =   2520
         Picture         =   "Form1.frx":6D578
         Tag             =   "pb"
         Top             =   5040
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":6F242
         Height          =   720
         Index           =   51
         Left            =   1680
         Picture         =   "Form1.frx":7088C
         Tag             =   "pb"
         Top             =   5040
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":72556
         Height          =   855
         Index           =   48
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":73BA0
         Height          =   855
         Index           =   47
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":751EA
         Height          =   855
         Index           =   46
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":76834
         Height          =   855
         Index           =   45
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":77E7E
         Height          =   855
         Index           =   44
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":794C8
         Height          =   855
         Index           =   43
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":7AB12
         Height          =   855
         Index           =   42
         Left            =   840
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":7C15C
         Height          =   855
         Index           =   41
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":7D7A6
         Height          =   855
         Index           =   40
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":7EDF0
         Height          =   855
         Index           =   39
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":8043A
         Height          =   855
         Index           =   38
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":81A84
         Height          =   855
         Index           =   37
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":830CE
         Height          =   855
         Index           =   36
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":84718
         Height          =   855
         Index           =   35
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":85D62
         Height          =   855
         Index           =   34
         Left            =   840
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":873AC
         Height          =   855
         Index           =   33
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":889F6
         Height          =   855
         Index           =   18
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":8A040
         Height          =   855
         Index           =   19
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":8B68A
         Height          =   855
         Index           =   20
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":8CCD4
         Height          =   855
         Index           =   21
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":8E31E
         Height          =   855
         Index           =   22
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":8F968
         Height          =   855
         Index           =   23
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":90FB2
         Height          =   855
         Index           =   24
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":925FC
         Height          =   855
         Index           =   26
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":93F96
         Height          =   855
         Index           =   27
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":95930
         Height          =   855
         Index           =   28
         Left            =   2520
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":974A6
         Height          =   855
         Index           =   29
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":98E40
         Height          =   855
         Index           =   30
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":9A7DA
         Height          =   855
         Index           =   31
         Left            =   5040
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":9C174
         Height          =   855
         Index           =   32
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":9D83E
         Height          =   855
         Index           =   25
         Left            =   0
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":9EF08
         Height          =   855
         Index           =   17
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":A0552
         Height          =   720
         Index           =   57
         Left            =   0
         Picture         =   "Form1.frx":A1B9C
         Tag             =   "tb"
         Top             =   5880
         Width           =   720
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":A3866
         Height          =   720
         Index           =   49
         Left            =   0
         Picture         =   "Form1.frx":A4EB0
         Tag             =   "pb"
         Top             =   5040
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   2
         Left            =   840
         Top             =   0
         Width           =   855
      End
      Begin VB.Image Image1 
         DragIcon        =   "Form1.frx":A6B7A
         Height          =   720
         Index           =   50
         Left            =   840
         Picture         =   "Form1.frx":A81C4
         Tag             =   "pb"
         Top             =   5040
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   11
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   10
         Left            =   840
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Image imgBack 
         Height          =   2040
         Left            =   0
         Top             =   0
         Width           =   2760
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   30
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   855
         Index           =   40
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   855
      End
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":A9E8E
      Top             =   9240
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const SizeCase = 48
Const TwipsPixel = 15
Const BackTop = 42
Const BackLeft = 42


Dim A As Integer, B As Integer, Prvi As String, Drugi As String
Dim c As Integer, D As Integer
Dim xBritneySpears As Single, Ponovi                       'znak za ponavljanje partije
Dim yBritneySpears As Single, BojaPolja As Integer
Dim Gore As Integer, Levo As Integer                       ' ZA AMPASAN
Dim VratiLeft As Integer, VratiTop As Integer, Pojedena As Integer





'===============================================================================
'
'===============================================================================
Private Sub Form_Load()

    imgBack.ZOrder 1
    WhiteMove = True
    BlackMove = False

End Sub



'==================================================================================
'
'==================================================================================
Public Sub ResizeForm(coef As Double)
    Dim img As Integer
    Dim Col As Integer


    GameSize = coef

    For Col = 0 To 7
        For img = 1 To 8
            With Image2(img + 8 * Col)
                .Stretch = True
                .Width = SizeCase * TwipsPixel * coef
                .Height = SizeCase * TwipsPixel * coef
                .Top = (BackTop + Col * SizeCase) * TwipsPixel * coef
                .Left = (BackLeft + (img - 1) * SizeCase) * TwipsPixel * coef
                'met les cases en haut
                .ZOrder 0
            End With


            With Image1(img + 8 * Col)
                .Stretch = True
                .Width = SizeCase * TwipsPixel * coef
                .Height = SizeCase * TwipsPixel * coef

                .Top = (BackTop + Col * SizeCase) * TwipsPixel * coef
                .Left = (BackLeft + (img - 1) * SizeCase) * TwipsPixel * coef
            End With
        Next
    Next

    For img = 1 To 16
        Image1(img).ZOrder 0
    Next

    For img = 49 To 64
        Image1(img).ZOrder 0
    Next


    With imgBack
        Set .Picture = LoadPicture(App.Path & "\echiquierB.jpg")
        .Stretch = True
        .Width = .Width * coef
        .Height = .Height * coef
    End With


    With Picture5
        .Height = imgBack.Height
        .Width = imgBack.Width
    End With

    With Form1
        .Width = Picture5.Width + .Width - .ScaleWidth
        .Height = Picture5.Height + .Height - .ScaleHeight
    End With


    'Image1(42).Picture = Image3(2).Picture

End Sub



'===============================================================================
'
'===============================================================================
Private Sub Image1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

    Dim i As Integer

    Debug.Print "IMG 1 " & Index & " " & Bouge
    'premakne figuro
    Source.Move Image1(Index).Left + X - xBritneySpears, Image1(Index).Top + Y - yBritneySpears

    'Ko spustis figuro da se poravna v polje
    Image1(Bouge).Left = Image1(Index).Left
    Image1(Bouge).Top = Image1(Index).Top

    CaseTo = Coordonnee(Image1(Index).Top, Image1(Index).Left)


    'Pojedena figura bo izginila

    If Image1(Bouge).Left = Image1(Index).Left _
            And Image1(Bouge).Top = Image1(Index).Top And Image1(Index).Picture <> LoadPicture("") Then
        Signal = 1
        'Bouge pojedene figure
        BrojOdneseneFigure = Image1(Index).Index
        'zapisi pozicijo LEFT odnesene figure
        PozicijaLeft2 = Image1(Index).Left
        'zapisi pozicijo TOP odnesene figure
        PozicijaTop2 = Image1(Index).Top
        'umakni pojedeno figuro z deske
        Image1(Index).Left = 7800
        'prikazi spusceno figuro
        Image1(Bouge).Visible = True
        'skrij pojedeno figuro
        Image1(Index).Visible = False

        TurnMove
        Exit Sub
    End If

    Signal = 0

    'figura se bo pokazala tam kamor je spuscena
    Image1(Bouge).Visible = True
    Image1(Bouge).ZOrder
    
    If Bouge = 61 And CaseFrom = "e1" And CaseTo = "g1" Then
        Image1(64).Left = (BackLeft + 5 * SizeCase) * TwipsPixel * GameSize
    End If

    If Bouge = 61 And CaseFrom = "e1" And CaseTo = "c1" Then
        Image1(57).Left = (BackLeft + 3 * SizeCase) * TwipsPixel * GameSize
    End If

    If Bouge = 5 And CaseFrom = "e8" And CaseTo = "g8" Then
        Image1(8).Left = (BackLeft + 5 * SizeCase) * TwipsPixel * GameSize
    End If

    If Bouge = 5 And CaseFrom = "e8" And CaseTo = "c8" Then
        Image1(1).Left = (BackLeft + 3 * SizeCase) * TwipsPixel * GameSize
    End If
    
    

    TurnMove

End Sub

'===============================================================================
'
'===============================================================================
Private Sub Image1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

    'If Image1(Index).Left >= 600 And Image1(Index).Left <= 6480 _
     And Image1(Index).Top >= 480 And Image1(Index).Top <= 6360 Then

    Image1(Bouge).Visible = False
    Source.Move Image1(Index).Left + X, Image1(Index).Top + Y
    Image1(Bouge).DragIcon = Image1(Bouge).Picture
    'Else
    'Image1(Bouge).DragIcon = LoadPicture("")
    'End If

End Sub

'===============================================================================
'
'===============================================================================
Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'récupération de la piece bougée
    Bouge = Index



    'zapomni si zacetni LEFT zajete figure
    PozicijaLeft = Image1(Index).Left
    'zapomni si zacetni TOP  zajete figure
    PozicijaTop = Image1(Index).Top
    BrojFigure = Image1(Index).Index

    CaseFrom = Coordonnee(Image1(Index).Top, Image1(Index).Left)
    Piece = Image1(Index).Tag


    xBritneySpears = X
    yBritneySpears = Y


    If BlackMove Then
        If Image1(Index).Picture <> LoadPicture("") Then
            Image1(Index).Drag 1
        Else
            Exit Sub
        End If
    End If


    If WhiteMove Then
        If Image1(Index).Picture <> LoadPicture("") Then
            Image1(Index).Drag 1
        Else
            Exit Sub
        End If
    End If

End Sub


'===============================================================================
'
'===============================================================================
Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    '    For i = 1 To 64
    '        If Image1(i).Tag = "" Then Image1(i).Visible = False
    '    Next

    'prikazi iznad figure roko
    
    If WhiteMove Then
        'If Index > 48 Then
            If Image1(Index).Picture <> LoadPicture("") Then
                Image1(Index).MousePointer = 99
                Image1(Index).MouseIcon = Image3(0).Picture
            Else
                Image1(Index).MousePointer = 0
                Image1(Index).MouseIcon = LoadPicture("")
            End If
        'End If
    End If

    'prikazi iznad figure roko

    If BlackMove Then
        'If Index < 17 Then
            If Image1(Index).Picture <> LoadPicture("") Then
                Image1(Index).MousePointer = 99
                Image1(Index).MouseIcon = Image3(0).Picture
            Else
                Image1(Index).MousePointer = 0
                Image1(Index).MouseIcon = LoadPicture("")
            End If
        'End If
    End If

End Sub

'===============================================================================
'
'===============================================================================
Private Sub Image2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

    Dim i As Integer

    Debug.Print "IMG 2 " & Index & " " & Bouge

    Source.Move Image2(Index).Left + X - xBritneySpears, Image2(Index).Top + Y - yBritneySpears

    'poravnaj spusceno figuro v polje
    Image1(Bouge).Left = Image2(Index).Left
    Image1(Bouge).Top = Image2(Index).Top

    CaseTo = Coordonnee(Image2(Index).Top, Image2(Index).Left)

    Pomakni = Image1(Bouge).Index
    Signal = 0
    Image1(Bouge).Visible = True
    Pomakni = Image1(Bouge).Index
    
    If Bouge = 61 And CaseFrom = "e1" And CaseTo = "g1" Then
        Image1(64).Left = (BackLeft + 5 * SizeCase) * TwipsPixel * GameSize
    End If

    If Bouge = 61 And CaseFrom = "e1" And CaseTo = "c1" Then
        Image1(57).Left = (BackLeft + 3 * SizeCase) * TwipsPixel * GameSize
    End If

    If Bouge = 5 And CaseFrom = "e8" And CaseTo = "g8" Then
        Image1(8).Left = (BackLeft + 5 * SizeCase) * TwipsPixel * GameSize
    End If

    If Bouge = 5 And CaseFrom = "e8" And CaseTo = "c8" Then
        Image1(1).Left = (BackLeft + 3 * SizeCase) * TwipsPixel * GameSize
    End If

    If CaseFrom <> CaseTo Then
        TurnMove
    End If

End Sub


'==================================================================================
'
'==================================================================================
Private Sub TurnMove()
    Dim Neu As Node
    Dim KeyNode As String

    On Error Resume Next
    BlackMove = Not BlackMove
    WhiteMove = Not WhiteMove

    KeyNode = CaseFrom & CaseTo

    Set Neu = Form2.TV.Nodes.Add(LastKey, tvwChild, LastKey & KeyNode, CaseFrom & "-" & CaseTo, Piece)
    If Err.Number <> 0 Then
        Set Neu = Form2.TV.Nodes(LastKey & KeyNode)
    End If
    Neu.EnsureVisible
    Neu.Selected = True
    
    LastKey = LastKey & KeyNode
End Sub


'==================================================================================
'
'==================================================================================
Private Function Coordonnee(t As Integer, l As Integer) As String
    Dim colonne As Integer
    Dim ligne As Integer
    colonne = (l - BackLeft * GameSize * TwipsPixel) / SizeCase / GameSize / TwipsPixel
    ligne = (t - BackTop * GameSize * TwipsPixel) / SizeCase / GameSize / TwipsPixel
    Coordonnee = Chr(97 + colonne) & CStr(8 - ligne)
End Function

'===============================================================================
'
'===============================================================================
Public Sub StartPos()
    ResizeForm 1
End Sub



'==================================================================================
'
'==================================================================================
Public Sub Dessineplateau()
    Dim i As Long

    Picture5.Visible = False
    For i = 1 To 64
        Image1(i).Tag = Plateau(i)
        Image1(i).ZOrder
        Image2(i).ZOrder 1
        Image1(i).Visible = True
        Select Case Plateau(i)
            Case ""
                Image1(i).Picture = Nothing
                'Image1(i).ZOrder 1
            Case "pb"
                Image1(i).Picture = Image3(2).Picture
            Case "tb"
                Image1(i).Picture = Image3(5).Picture
            Case "cb"
                Image1(i).Picture = Image3(3).Picture
            Case "fb"
                Image1(i).Picture = Image3(4).Picture
            Case "db"
                Image1(i).Picture = Image3(6).Picture
            Case "rb"
                Image1(i).Picture = Image3(7).Picture

            Case "pn"
                Set Image1(i).Picture = Image3(8).Picture
            Case "tn"
                Set Image1(i).Picture = Image3(11).Picture
            Case "cn"
                Set Image1(i).Picture = Image3(9).Picture
            Case "fn"
                Set Image1(i).Picture = Image3(10).Picture
            Case "dn"
                Set Image1(i).Picture = Image3(12).Picture
            Case "rn"
                Set Image1(i).Picture = Image3(13).Picture
        End Select
    Next
    imgBack.ZOrder 1
    Picture5.Visible = True
End Sub

