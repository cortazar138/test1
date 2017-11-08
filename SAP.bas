Attribute VB_Name = "SAP"
Sub Connect(session As Variant, Optional n As Integer = 0)
'session: name of session you want to create
'n: index of sap windows you want to open
If Not IsObject(App) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set App = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = App.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(n + 0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject App, "on"
End If


End Sub


Sub Save(session, Optional info)
'session: name of session you want to save
'info: name of variable you want to store the text from the sbar
session.FindById("wnd[0]").SendVKEy (11)
info = session.FindById("wnd[0]/sbar").Text

End Sub


Sub OA_Line(session As Variant, line As Integer)

session.FindById("wnd[0]/usr/txtRM06E-EBELP").Text = line
session.FindById("wnd[0]").SendVKEy 0
session.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/txtRM06E-EVRTP[0,0]").SetFocus
session.FindById("wnd[0]/usr/tblSAPMM06ETC_0220/txtRM06E-EVRTP[0,0]").caretPosition = 2

End Sub

Sub PO_Address(session As Variant, adr)
'searches for the mobile address of the PO returnes it as adr variable
Dim Ruchom As Variant
Ruchoma = Array("19", "15", "10", "20", "17")

On Error Resume Next

For Each element In Ruchoma
    session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & CStr(element) & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[5,0]").SetFocus 'Conditions
    If Err.Number = "0" Then
        adr = element
        GoTo Fin
    End If
    Err.Clear
Next
Fin:
End Sub


Sub PO_Adr(session As Variant, adr)
'alias for PO_address
Call PO_Address(session, adr)

End Sub

Sub PO_Item(session As Variant, adr, item As Integer)
'adr: mobile address of the PO
Dim nazwa As String
Dim it As Integer
Sprawdz:
nazwa = session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & CStr(adr) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").Text
nazwa = Left(nazwa, 3)
nazwa = Replace(nazwa, " ", "")
nazwa = Replace(nazwa, "[", "")
it = CInt(nazwa) * 10

If it = item Then
    Exit Sub
ElseIf it < item Then
    session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & CStr(adr) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").Press
    GoTo Sprawdz
Else
    session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & CStr(adr) & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT001").Press
    GoTo Sprawdz
End If

End Sub


Sub Open_OA(session As Variant, OA As String)


session.SendCommand ("/nME32K")
session.FindById("wnd[0]/usr/ctxtRM06E-EVRTN").Text = OA
session.FindById("wnd[0]").SendVKEy 0


End Sub


Sub PR_Line(session As Variant, item As Integer)

'adr: mobile address of the PO
Dim nazwa As String
Dim it As Integer
Sprawdz:
nazwa = session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").Text
nazwa = Left(nazwa, 3)
nazwa = Replace(nazwa, " ", "")
nazwa = Replace(nazwa, "[", "")
it = CInt(nazwa) * 10

If it = item Then
    Exit Sub
ElseIf it < item Then
    session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT002").Press
    GoTo Sprawdz
Else
    session.FindById("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/btn%#AUTOTEXT001").Press
    GoTo Sprawdz
End If

End Sub


Sub MM03_Tab(session As Variant, tab_text As String)
'selecting tab inside MM03 transaction
'bug works only for till Work scheduling tab

k = 1
tabb = session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP0" & CStr(k) & "").Text
Do Until tabb Like "*" & tab_text & "*"
Err.Clear
On Error GoTo Next_k
        k = k + 1
    If k = 24 Then
        Exit Do
    End If
    If k < 10 Then
        tabb = session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP0" & CStr(k) & "").Text
    Else
        tabb = session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP" & CStr(k) & "").Text
    End If
Next_k:
Loop

If k < 10 Then
    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP0" & CStr(k) & "").Select
Else
    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP" & CStr(k) & "").Select
End If
End Sub
