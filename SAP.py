import win32com.client
import os
def Connect(n=0):
    SapGui = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    session = SapGui.FindById("ses[" + str(n) + "]")
    return session

def Verify(session, control):
    try:
        session.findByID(control)
        return session.findByID(control)
    except:
        print("could not find " + control)
        return

def PO_adr(session):
    rozne = [16, 19, 15, 17]
    for ele in rozne:
        try:
            zmienna = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" + str(ele) + "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[5,0]").Text
            address = str(ele)
            return address
            break
        except:
            pass

def Select_OA_Line(session='session', linia=10):
    #do ogarniec
    pass

def OA_Line_Select(session, linia):
    #Select_OA_Line(session, linia)
    pass

def Select_PO_Line(session, linia):
    #do ogarniecia
    pass 

def PO_Line_Select(session, linia):
    #Select_PO_Line(session, linia)
    pass
def Select_PR_Line(session, linia):
    #do ogarniecia
    pass

def PR_Line_Select(session, linia):
    Select_PR_Line(session, linia)

def Save(session):
    session.findById("wnd[0]").SendvKey(11)

def Last_Price(session):
    i = 0
    session.findbyid("wnd[0]").SendVkey(18)
    while True:
        try:
            session.findbyid("wnd[1]").SendVkey(24)
            data = session.findbyid("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATBI[1," + str(i) + "]").Text
            if data == '__________':
                session.findbyid("wnd[1]").SendVkey(0)
                break
            session.findbyid("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATBI[1," + str(i) + "]").SetFocus()
            session.findbyid("wnd[1]/usr/tblSAPLV14ATCTRL_D0102/ctxtVAKE-DATBI[1," + str(i) + "]").caretPosition = 4
            i = i + 1
        except:
            try:
                session.findbyid("wnd[1]").SendVkey(0)
                break
            except:
                break