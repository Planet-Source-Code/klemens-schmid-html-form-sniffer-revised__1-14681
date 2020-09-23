VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class catches the events of the monitored WebBrowser

Option Explicit

Dim WithEvents moIE As WebBrowser
Attribute moIE.VB_VarHelpID = -1

Public Sub Init(Browser As WebBrowser)
'assign the event sink
Set moIE = Browser
End Sub

Private Sub moIE_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
'ignore scripts calls
If Not LCase(URL) Like "http://*" Then Exit Sub
'show URL and postdata
frmSniff.AddRequestData URL, StrConv(PostData, vbUnicode)
End Sub
