VERSION 5.00
Object = "*\AMagicLabel.vbp"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin MagicLabel.GradientMagicLabel GradientMagicLabel1 
      Height          =   1965
      Index           =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   2175
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   3466
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlignmentV      =   1
      AlignmentH      =   1
      ForeColor       =   255
      Filigree        =   "Object Locked - please register"
   End
   Begin MagicLabel.GradientMagicLabel GradientMagicLabel1 
      Align           =   1  'Align Top
      Height          =   1065
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   1879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Top aligned"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

