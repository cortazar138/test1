Attribute VB_Name = "Module2"
Sub Cena()

Dim a As Boolean
Dim b As Boolean
Dim c As Boolean
Dim scales As Boolean
Dim Blad As String
Dim k As Integer
Dim Valid As Date

If Not IsObject(App) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set App = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = App.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject App, "on"
End If

For i = 3 To Cells(1, 1).Value

'sprawdza czy sa roznice miedzy MDC a SAP
a = Cells(i, 19).Value <> Cells(i, 15).Value
b = Cells(i, 20).Value <> Cells(i, 16).Value
c = Cells(i, 21).Value <> Cells(i, 17).Value

If a Or b Or c Then
'sprawdza czy sa skale
If Cells(i, 24).Value = "" Then
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme32k"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRM06E-EVRTN").Text = Cells(i, 12).Value
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/txtRM06E-EBELP").Text = Cells(i, 13).Value
    session.findById("wnd[0]/usr/txtRM06E-EBELP").SetFocus
    session.findById("wnd[0]/usr/txtRM06E-EBELP").caretPosition = 3
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[18]").Press
    'scorluje do ostatniego wiersza
    k = 0
    Valid = "01.01.1999"
    Do
    Valid = session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATBI[1,0]").Text
    session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102").verticalScrollbar.Position = k + 1
    session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATBI[1,0]").SetFocus
    session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATBI[1,0]").caretPosition = 4
    k = k + 1
    Loop While Valid <> session.findById("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATBI[1,0]").Text
    session.findById("wnd[1]/tbar[0]/btn[8]").Press
    k = 0
    scales = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/chkRV13A-KOSTKZ[7," & CStr(k) & "]").Selected
    Condition = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/ctxtKONP-KSCHL[0," & CStr(k) & "]").Text
    'szuka nieusunietej ceny PB00
    Do Until Condition = "PB00" And scales = False
    scales = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/chkRV13A-KOSTKZ[7," & CStr(k) & "]").Selected
    Condition = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/ctxtKONP-KSCHL[0," & CStr(k) & "]").Text
    k = k + 1
    Loop
    'sprawdza czy zmienic validity date jak nie daye dzisiejsza
    If Cells(i, 27).Value Then
    session.findById("wnd[0]/usr/ctxtRV13A-DATAB").Text = Cells(i, 23).Value 'valid from
    Else
    session.findById("wnd[0]/usr/ctxtRV13A-DATAB").Text = Date
    End If
    'reszta danych (cena, waluta itp)
    session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/txtKONP-KBETR[2," & CStr(k) & "]").Text = Cells(i, 19).Value 'price
    session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/ctxtKONP-KONWA[3," & CStr(k) & "]").Text = Cells(i, 20).Value
    session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/txtKONP-KPEIN[4," & CStr(k) & "]").Text = Cells(i, 21).Value
    session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/ctxtKONP-KMEIN[5," & CStr(k) & "]").Text = Cells(i, 22).Value
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[3]").Press
    Blad = ""
    
    'overlaping validity periods
    On Error Resume Next
    Blad = session.findById("wnd[1]").Text
    If Blad Like "Errors as*" Then
    session.findById("wnd[1]").sendVKey 5
    End If
    'save
    session.findById("wnd[0]/tbar[1]/btn[48]").Press
    session.findById("wnd[0]").sendVKey 3
    session.findById("wnd[0]").sendVKey 11
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").Press
Else
End If
End If
Next i

End Sub
