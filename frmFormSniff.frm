VERSION 5.00
Begin VB.Form frmSniff 
   Caption         =   "Form Sniffer"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReplay 
      Caption         =   "&Replay"
      Enabled         =   0   'False
      Height          =   330
      Left            =   2280
      TabIndex        =   8
      Top             =   4050
      Width           =   975
   End
   Begin VB.TextBox txtReplayURL 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4440
      Width           =   4815
   End
   Begin VB.HScrollBar scrRequests 
      Height          =   255
      Left            =   120
      Max             =   0
      TabIndex        =   5
      Top             =   720
      Width           =   4815
   End
   Begin VB.ComboBox cboWindows 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Select Browser Window"
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox txtPostData 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2880
      Width           =   4815
   End
   Begin VB.TextBox txtURL 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label lblReplayURL 
      Caption         =   "Replay to replay request:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label lblPostData 
      Caption         =   "PostData:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label lblURL 
      Caption         =   "URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmSniff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:
'  klemens.schmid@gmx.de, http://www.watchtheweb.de
'Description
'  Captures the URL and the post values just before IE
'  fires it off to the site. The advantage of this method
'  is that it works in any case regardless of frames.
'  It also assembles a URL which can be used to replay
'  the request it it contains post data.
'Prerequisite
'  You are using IE4 or IE5
'  Your VB project includes references to 'Microsoft Internet Controls'
'Known issues
'  Unfortunately this method doesn't allow to inspect the cookies of
'  a request. This would require a read http sniffer.
'Created
'  24 Jan 2001
'Changes
'  02 Jul 2001: Added assembly of replay URL

Option Explicit

Private Declare Function ShellExecute Lib _
   "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

Private Type RequestData
   URL As String
   PostData As String
End Type

Dim aRequest() As RequestData             'keeps all requests
Dim moBrowser As clsBrowser               'catches WebBrowser's events
Dim cntRequests As Integer                'number of sniffed requests

Private Sub cboWindows_Click()
'Monitor the events of the chosen browser window

Dim oSWs As ShellWindows
Dim oIE As SHDocVw.InternetExplorer
Dim p%, i%
Dim bFound As Boolean
Dim intSelected As Integer
Dim rc As Integer

intSelected = cboWindows.ListIndex
'fetch the selected window
Set oSWs = New ShellWindows
Set oIE = oSWs.Item(cboWindows.ItemData(intSelected))
Set moBrowser = New clsBrowser
moBrowser.Init oIE
MsgBox "Sniffing is enabled. Please continue browsing."
End Sub

Private Sub cmdReplay_Click()
Call ShellExecute(0, vbNullString, txtReplayURL.Text, vbNullString, "c:\", 1)
End Sub

Private Sub Form_Load()
Dim oSWs As ShellWindows
Dim oIE As SHDocVw.InternetExplorer
Dim oDoc As Object
Dim p%, i%
Dim bFound As Boolean

ReDim aRequest(10)
Set oSWs = New ShellWindows
'loop thru the open browser windows and take the first
For i = 0 To oSWs.Count - 1
   Set oIE = oSWs.Item(i)
   Set oDoc = oIE.Document
   If TypeName(oDoc) = "HTMLDocument" Then
      cboWindows.AddItem oDoc.Title
      cboWindows.ItemData(cboWindows.ListCount - 1) = i
      bFound = True
   End If
Next
If Not bFound Then
   cboWindows.Clear
   cboWindows.Text = "No browser window found"
Else
   cboWindows.Text = "Please choose a browser window"
End If
   
End Sub

Public Sub AddRequestData(ByVal URL$, ByVal PostData$)
'add a new request
If cntRequests > UBound(aRequest) Then
   ReDim Preserve aRequest(cntRequests + 5)
End If
aRequest(cntRequests).URL = URL
aRequest(cntRequests).PostData = PostData
'set the scrollbar
scrRequests.Max = cntRequests
scrRequests.Value = cntRequests
'point to next free element
cntRequests = cntRequests + 1
'dead sure:
Call scrRequests_Change
End Sub

Private Sub Form_Resize()
'make the form smart
Dim intWidth%
intWidth = ScaleWidth - 2 * cboWindows.Left
cboWindows.Width = intWidth
scrRequests.Width = intWidth
txtPostData.Width = intWidth
txtURL.Width = intWidth
txtReplayURL.Width = intWidth
txtReplayURL.Height = ScaleHeight - txtReplayURL.Top - cboWindows.Top
End Sub

Private Sub scrRequests_Change()
txtURL = aRequest(scrRequests.Value).URL
txtPostData = aRequest(scrRequests.Value).PostData
If Len(txtPostData.Text) > 0 Then
   'assemble replay URL
   'this URL points to my public ASP page
   txtReplayURL = "http://www.watchtheweb.de/get2post.asp?g2p_action=" & aRequest(scrRequests.Value).URL & "&" & aRequest(scrRequests.Value).PostData
Else
   txtReplayURL = "Replay doesn't work for this request because it doesn't contain postdata"
End If
cmdReplay.Enabled = (Len(txtReplayURL.Text) > 0)
End Sub
