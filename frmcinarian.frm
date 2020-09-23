VERSION 5.00
Begin VB.Form frmcinarian 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " How to read Cinarian Time"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6945
   Icon            =   "frmcinarian.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblcinarian 
      BackStyle       =   0  'Transparent
      Caption         =   "AM/PM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblcinarian 
      BackStyle       =   0  'Transparent
      Caption         =   "4/nSeconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblcinarian 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3/nOnes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblcinarian 
      BackStyle       =   0  'Transparent
      Caption         =   "Hour/n1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblcinarian 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tens/n2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblcinarian 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmcinarian.frx":000C
      Height          =   3975
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmcinarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim temp As Label
    For Each temp In Me
        temp.Caption = Replace(temp.Caption, "/n", vbNewLine)
    Next
    drawCinarian Me, 1350, 2150, 1200, , , "1111222333444"
End Sub
