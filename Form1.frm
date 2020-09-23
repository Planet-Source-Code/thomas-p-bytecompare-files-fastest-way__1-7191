VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ByteCompare"
   ClientHeight    =   3984
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   2976
   LinkTopic       =   "Form1"
   ScaleHeight     =   3984
   ScaleWidth      =   2976
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command3 
      Caption         =   "&Start Compare"
      Height          =   372
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   2532
   End
   Begin VB.ListBox List3 
      Height          =   1968
      Left            =   2280
      TabIndex        =   6
      Top             =   1800
      Width           =   492
   End
   Begin VB.ListBox List2 
      Height          =   1968
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   492
   End
   Begin VB.ListBox List1 
      Height          =   1968
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1332
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   252
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   372
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   252
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   372
   End
   Begin VB.TextBox txtPatched 
      Height          =   288
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2052
   End
   Begin VB.TextBox txtOriginal 
      Height          =   288
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "O"
      Height          =   192
      Left            =   1680
      TabIndex        =   9
      Top             =   1560
      Width           =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "P"
      Height          =   192
      Left            =   2280
      TabIndex        =   8
      Top             =   1560
      Width           =   108
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Offset"
      Height          =   192
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   408
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This SourceCode shows you how to compare 2 files the
'fastest possible way in VB.
'
'
'greets
'
'Tom (tom@evilemail.com)
'***********************

Private Sub Command1_Click()
txtOriginal.Text = Open_File(Me.hWnd)
End Sub

Private Sub Command2_Click()
txtPatched.Text = Open_File(Me.hWnd)
End Sub

Private Sub Command3_Click()
ByteCompare txtOriginal.Text, txtPatched.Text, List1, List2, List3
End Sub

