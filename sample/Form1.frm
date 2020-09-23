VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Hide
Dim rp As New Database.Report
rp.TitleBold = True
rp.TitleItalic = False
rp.TitleUnderline = True
rp.TitleColor = colRed
rp.TitleFont = fntCursive
rp.TitleSize = 16
rp.BackGround = colWhite
rp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db.mdb;Persist Security Info=False"
rp.SQLStatement = "select * from books"
rp.DisplayReport
End Sub
