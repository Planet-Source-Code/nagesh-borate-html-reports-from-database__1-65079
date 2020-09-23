VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum ColorValues
colBlue
colBlack
colRed
colGreen
colYellow
colCyan
colWhite
End Enum

Public Enum FontValues
fntSerif
fntSans_serif
fntCursive
fntFantasy
fntMonospace
End Enum


Public TitleText As String
Public TitleFont As FontValues
Public TitleSize As Integer
Public TitleColor As ColorValues
Public TitleBold As Boolean
Public TitleItalic As Boolean
Public TitleUnderline As Boolean
Public ConnectionString As String
Private objrs As New ADODB.Recordset
Public SQLStatement As String
Private fs As New FileSystemObject
Private ts As TextStream
Public BackGround As ColorValues

Private Sub Class_Initialize()
ConnectionString = ""
SQLStatement = ""
BackGround = colCyan


TitleText = "Report"
TitleBold = True
TitleItalic = False
TitleUnderline = True
TitleFont = fntFantasy
TitleColor = colBlue
TitleSize = 14



End Sub

Public Sub DisplayReport()
On Error GoTo erh

If Trim(ConnectionString) = "" Or Trim(SQLStatement) = "" Then
MsgBox "Error"
Exit Sub
End If

objrs.Open SQLStatement, ConnectionString, adOpenDynamic, adLockOptimistic
objrs.MoveFirst
Set ts = fs.CreateTextFile(App.Path & "\rep.html", True)

ts.WriteLine ("<html><head>")
ts.WriteLine ("<script language=""vbscript"" for=cmdprn event=onclick>")
ts.WriteLine ("window.print")
ts.WriteLine ("</script>")
ts.WriteLine ("</head>")
ts.Write ("<body bgcolor=")

Select Case BackGround
Case colBlue
ts.Write ("blue")
Case colBlack
ts.Write ("black")
Case colRed
ts.Write ("red")
Case colGreen
ts.Write ("green")
Case colYellow
ts.Write ("yellow")
Case colCyan
ts.Write ("cyan")
Case colWhite
ts.Write ("white")
End Select

ts.WriteLine ("> ")
ts.Write ("<center><h1>")

If TitleBold Then
ts.Write ("<b>")
End If

If TitleItalic Then
ts.Write ("<i>")
End If

If TitleUnderline Then
ts.Write ("<u>")
End If

Select Case TitleFont
Case fntSerif
hf = "serif"
Case fntSans_serif
hf = "sans-serif"
Case fntCursive
hf = "cursive"
Case fntFantasy
hf = "fantasy"
Case fntMonospace
hf = "monospace"
End Select

Select Case TitleColor
Case colBlue
hc = "blue"
Case colBlack
hc = "black"
Case colRed
hc = "red"
Case colGreen
hc = "green"
Case colYellow
hc = "yellow"
Case colCyan
hc = "cyan"
Case colWhite
hc = "white"
End Select


ts.Write ("<font face=" & hf & " size=" & TitleSize & " color=" & hc & " >" & TitleText & "</font>")

If TitleUnderline Then
ts.Write ("</u>")
End If

If TitleItalic Then
ts.Write ("</i>")
End If

If TitleBold Then
ts.Write ("</b>")
End If

ts.Write ("</h1></center>" & vbCrLf)
ts.WriteLine ("<table border=1 align=center><tr>")

For i = 0 To objrs.Fields.Count - 1
ts.WriteLine ("<th>" & objrs.Fields(i).Name)
Next

Do While Not objrs.EOF
ts.WriteLine ("<tr>")
For i = 0 To objrs.Fields.Count - 1
ts.WriteLine ("<td>" & objrs.Fields(i))
Next
objrs.MoveNext
Loop

ts.WriteLine ("</table>")
ts.WriteLine ("<br><br>")
ts.WriteLine ("<form name=frmprn>")
ts.WriteLine ("<right>")
ts.WriteLine ("<input type=button name=cmdprn value='Print'>")
ts.WriteLine ("</right>")
ts.WriteLine ("</form>")
ts.WriteLine ("</body></html>")
ts.Close
Set fs = Nothing

Load Form1
Form1.Caption = TitleText
Form1.wb1.Navigate App.Path & "\rep.html"
Form1.Show
Exit Sub
erh:
MsgBox "Error"
End Sub
