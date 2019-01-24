VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProxy 
   Caption         =   "Proxy Switcher v1.0"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6825
   OleObjectBlob   =   "frmProxy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function internetsetoption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hinternet As Long, ByVal dwoption As Long, ByRef lpbuffer As Any, ByVal dwbufferlength As Long) As Long
Private ws As New WshShell

Private Sub CommandButton2_Click()
If CheckBox1.Value = True Then Call Reset
Unload Me
End Sub

Private Sub ListBox1_Click()
If ListBox1.List(ListBox1.ListIndex, 0) = "Default" Then
    Call Reset
Else
    Call setProxy(ListBox1.List(ListBox1.ListIndex, 1), ListBox1.List(ListBox1.ListIndex, 2))
End If
End Sub
 
Private Sub setProxy(ByVal ip As String, port As String)
If ip = "" Or port = "" Then Exit Sub
ws.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"  'enable proxy
ws.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", Join(Array(ip, port), ":"), "REG_SZ" '写入的代理服务器地址
ws.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride", "<local>", "REG_SZ" 'bypass local adress
ws.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigURL", "http://iepac.utc.com/iepac/tproxies.pac", "REG_SZ"
ws.RegDelete "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigURL"   'disable auto config script
Call internetsetoption(0, 39, 0, 0)
End Sub
Private Sub Reset()
ws.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"  'disable proxy
ws.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigURL", "http://iepac.utc.com/iepac/tproxies.pac", "REG_SZ" ' enable auto config script
Call internetsetoption(0, 39, 0, 0)
End Sub


Private Sub UserForm_Initialize()
Dim ar(1 To 8, 0 To 2) As Variant
'ar(0, 0) = "Region"
'ar(0, 1) = "Proxy"
'ar(0, 2) = "Port"

ar(1, 0) = "CN"
ar(1, 1) = "172.28.41.32"
ar(1, 2) = "8080"

ar(2, 0) = "SG"
ar(2, 1) = "sgcorpproxy-vip1.utc.com"
ar(2, 2) = "8080"

ar(3, 0) = "SG"
ar(3, 1) = "sgcorpproxy-vip2.utc.com"
ar(3, 2) = "8080"

ar(4, 0) = "FR"
ar(4, 1) = "frcorpproxyva.utc.com"
ar(4, 2) = "8089"

ar(5, 0) = "US"
ar(5, 1) = "comproxy.utc.com"
ar(5, 2) = "8080"

ar(6, 0) = "UK"
ar(6, 1) = "ukcorpproxyva.utc.com"
ar(6, 2) = "8089"

'ar(7, 0) = "goAgent"
'ar(7, 1) = "127.0.0.1"
'ar(7, 2) = "8087"

ar(8, 0) = "Default"
ar(8, 1) = "http://iepac.utc.com/iepac/tproxies.pac"
ar(8, 2) = ""

ListBox1.List = ar
ListBox1.Value = ""

Dim p As String
On Error GoTo testproxy
p = ws.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\AutoConfigURL")
If p = "http://iepac.utc.com/iepac/tproxies.pac" Then
ListBox1.Value = p
End If
Exit Sub
testproxy:
On Error GoTo unknown
If ws.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable") = 1 Then
    p = ws.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer")
    ListBox1.Value = Split(p, ":")(0)
End If
Exit Sub
unknown:
'do nothing
Exit Sub
End Sub

Private Sub UserForm_Terminate()
If CheckBox1.Value = True Then Call Reset
Set ws = Nothing
End Sub
