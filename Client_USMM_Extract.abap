' Script for extracting USMM-data from Client.
' Avaa CompanyReportExtraction.xlsm ja avaa Module5, seuraavaksi aja skripti
Option Explicit
Public SapGuiAuto
Public objGui  As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

Sub SAPCustomerReport()

Set SapGuiAuto = GetObject("SAPGUI")
Set objGui = SapGuiAuto.GetScriptingEngine
Set objConn = objGui.Children(0)
Set session = objConn.Children(0)

session.FindById("wnd[0]").ResizeWorkingPane 98, 38, False
session.FindById("wnd[0]/tbar[0]/okcd").Text = "usmm"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[18]").Press
session.FindById("wnd[0]/usr/cntlSLIM_USER_CONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
session.FindById("wnd[0]/usr/cntlSLIM_USER_CONTAINER/shellcont/shell").SelectContextMenuItem "&PC"
session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
session.FindById("wnd[1]/tbar[0]/btn[0]").Press 'Muuta kuukausi seuraavalta riviltä joka kerta muotoon *09 Syyskuu*
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\Otto Karppinen\OneDrive - GRC Nordic\Tiedostot\MG\Raportointi\TC Reporting\2022\kuukausi\02 Client\1 Original"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "WOOD_USMM" 'Muuta WOOD_USMM riviltä jos kyseessa toinen company
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 9
session.FindById("wnd[1]/tbar[0]/btn[0]").Press

MsgBox "Scripti valmis ja ladattu systeemin original-kansioon"

End Sub