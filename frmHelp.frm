VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   6435
   ClientLeft      =   4785
   ClientTop       =   3570
   ClientWidth     =   9960
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6435
   ScaleWidth      =   9960
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   10186
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmHelp.frx":08CA
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    RichTextBox1.LoadFile App.Path & "\" & "Help.rtf"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload frmHelp
End Sub
