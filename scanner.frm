VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form scanner 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Web Page Info"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "scanner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton exit 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Height          =   195
      Left            =   600
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5880
      Width           =   75
   End
   Begin VB.TextBox websiteinfo 
      Height          =   285
      Left            =   5880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   3000
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2280
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton save_html 
      Caption         =   "&Save HTML"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton save_info 
      Caption         =   "S&ave Info"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton reset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin MSComctlLib.ListView PageList 
      Height          =   1695
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2214
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   8643
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton go_cmd 
      Caption         =   "&Get Info"
      Default         =   -1  'True
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox website_txt 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   5880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox websitehtml 
      Height          =   2205
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   6255
   End
   Begin VB.Label info1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "URL:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Page Information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Page HTML:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "scanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'easy way to add things to the list
Private Sub AddList(TheName As String, TheDesc As String)
Set lstAdd = PageList.ListItems.Add(, , TheName)
    lstAdd.SubItems(1) = TheDesc

websiteinfo = websiteinfo + TheName & " - " & TheDesc & vbCrLf

End Sub



Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
PageList.AllowColumnReorder = True

End Sub

Private Sub go_cmd_Click()
info1.Caption = "Connecting..."
go_cmd.Enabled = False

Call ResetStats
'opens the inet controll to get the page's html

websitehtml = Inet.OpenURL("" & website_txt & "")
info1.Caption = "Processing HTML"
go_cmd.Enabled = True

PB = PB + 15 'updates the progress bar showing that we got the html


'checks the html
Call CheckHTML

End Sub


Function ResetStats()
PB = 0 'start setting the progress bar
PageList.ListItems.Clear
websiteinfo = ""
websitehtml = ""
End Function


'this function checks the html for
'the hardcoded charateristics
'hopfully i can make it so it loads the
'characteristics from a data file and
'checks it from that
Function CheckHTML()


websitehtml.Text = UCase(websitehtml)


'set the text report up
websiteinfo = "-- Page Scan Report --" & vbCrLf & vbCrLf
websiteinfo = websiteinfo + "URL Scanned: " & website_txt.Text & vbCrLf & vbCrLf
websiteinfo = websiteinfo + "Items Found:" & vbCrLf


'checks for javascript
PB = PB + 5
If InStr(1, websitehtml, "<SCRIPT") Then
    AddList "Scripting", "The page contains some scripting"
End If

'check for 404 error
PB = PB + 5
If InStr(1, websitehtml, "HTTP 404") Or InStr(1, websitehtml, "404 ERROR") Then
    AddList "404 Error", "The page brought up a 404 error"
End If


'start checks for popups
PB = PB + 5
If InStr(1, websitehtml, "WINDOW.OPEN") Then
    AddList "Popups", "The page contains popups"
End If

'checks for page redirect
PB = PB + 10
If InStr(1, websitehtml, "<META HTTP-EQUIV=""REFRESH""") Or InStr(1, websitehtml, "DOCUMENT.LOCATION.HREF") Then
    AddList "Redirect", "The page contains a redirect"
End If


'checks for bookmarking
PB = PB + 2.5
If InStr(1, websitehtml, "WINDOW.EXTERNAL.ADDFAVORITE") Then
    AddList "Bookmark", "The page contains JavaScript to bookmark the page"
End If

'checks for setting homepage .setHomePage
PB = PB + 2.5
If InStr(1, websitehtml, ".SETHOMEPAGE") Then
    AddList "Homepage", "The page contains JavaScript to set the Homepage"
End If

'checks for java applet
PB = PB + 5
If InStr(1, websitehtml, "<APPLET") Then
    AddList "Java", "The page contains a Java Applet"
End If


'checks for embedded object
PB = PB + 5
If InStr(1, websitehtml, "<EMBED") Then
    AddList "Embed", "The page contains a Embedded Object"
End If


'checks for frames
PB = PB + 5
If InStr(1, websitehtml, "<FRAMESET") Then
    AddList "Frames", "The page contains Frames"
End If


'checks for a <object
PB = PB + 5
If InStr(1, websitehtml, "<OBJECT") Then
    AddList "Object", "The page contains a ActiveX Control"
End If

'checks for cookies
PB = PB + 5
If InStr(1, websitehtml, "DOCUMENT.COOKIE") Then
    AddList "Cookie", "The page Creates a Cookie"
End If


'checks for a designer controll
PB = PB + 5
If InStr(1, websitehtml, "<!--METADATA TYPE=""DESIGNERCONTROL""") Then
    AddList "DesignerControll", "The page contains a Designer Control"
End If

'checks for a navigator.userAgent
PB = PB + 5
If InStr(1, websitehtml, "NAVIGATOR.USERAGENT") Then
    AddList "UserAgent", "The page contains code to get the User Agent"
End If

'checks for frontpage
PB = PB + 2.5
If InStr(1, websitehtml, "<META NAME=""GENERATOR"" CONTENT=""MICROSOFT FRONTPAGE") Then
    AddList "Frontpage", "The page was made in frontpage"
End If

'checks for javascript alert
PB = PB + 2.5
If InStr(1, websitehtml, "ALERT(""") Then
    AddList "Msgbox", "The page has JavaScript Msgbox's"
End If


'check for a form
PB = PB + 5
If InStr(1, websitehtml, "<FORM") Then
    AddList "Form", "The page Contains a form"
End If

'checks for style sheets
PB = PB + 2.5
If InStr(1, websitehtml, "<STYLE") Or InStr(1, websitehtml, "<LINK REL=""stylesheet""") Then
    AddList "Style Sheets", "The page uses style sheets"
End If


'checks for bess blocking it
'<!-- $ID: BLOCK.HTML
PB = PB + 2.5
If InStr(1, websitehtml, "<!-- $ID: BLOCK.HTML") Then
    AddList "Blocked", "The page has been blocked by bess"
End If




MsgBox "Found " & PageList.ListItems.Count & " Characteristics"
'resets the Progress Bar
PB = 0
info1.Caption = "Done"
End Function
Private Sub reset_Click()
'reset everything and set focus

Call ResetStats
website_txt = ""
website_txt.SetFocus

End Sub

Private Sub save_html_Click()
Call SaveHTML

End Sub

Private Sub save_info_Click()
Call SaveInfo
End Sub

'push enter it pushes the button
Private Sub website_txt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
go_cmd_Click
End If
End Sub

Function SaveInfo()
'open up the dialog box
CommonDialog.Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*"
CommonDialog.ShowSave
'save the file
If CommonDialog.FileName = "" Then
Else
Open CommonDialog.FileName For Append As 1
Print #1, websiteinfo
Close 1
MsgBox "Saved Successfuly"
End If
End Function

Function SaveHTML()
'open up the dialog box
CommonDialog.Filter = "Text File (*.txt)|*.txt|HTML (*.htm)|*.htm|All Files (*.*)|*.*"
CommonDialog.ShowSave
'save the file
If CommonDialog.FileName = "" Then
Else
Open CommonDialog.FileName For Append As 1
Print #1, websitehtml
Close 1
MsgBox "Saved Successfuly"
End If
End Function
