' Script for extracting LOG-data from ClientName.
' Avaa CompanyReportExtraction.xlsm ja avaa Module19
' Aja skripti
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

session.FindById("wnd[0]").ResizeWorkingPane 159, 38, False
session.FindById("wnd[0]/tbar[0]/okcd").Text = "suim"
session.FindById("wnd[0]").SendVKey 0
session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").ExpandNode "02  1      2"
session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").TopNode = "01  1      1"
session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").SelectItem "03  2     16", "1"
session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").EnsureVisibleHorizontalItem "03  2     16", "1"
session.FindById("wnd[0]/usr/cntlTREE_CONTROL_CONTAINER/shellcont/shell").ClickLink "03  2     16", "1"
session.FindById("wnd[0]/usr/chkREFUSER").Selected = False
session.FindById("wnd[0]/usr/chkSERVUSER").Selected = False
session.FindById("wnd[0]/usr/chkSYSUSER").Selected = False
session.FindById("wnd[0]/usr/chkCOMMUSER").Selected = False
session.FindById("wnd[0]/usr/chkCOMMUSER").SetFocus
session.FindById("wnd[0]/tbar[1]/btn[8]").Press
session.FindById("wnd[0]/tbar[1]/btn[45]").Press
session.FindById("wnd[1]/tbar[0]/btn[0]").Press ' Vaihda seuraavalta riviltä kuukausi
session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\Users\Otto Karppinen\OneDrive - GRC Nordic\Tiedostot\MG\Raportointi\TC Reporting\2022\09 Syyskuu\05 ClientName\1 Original"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ClientName_LOG"
session.FindById("wnd[1]/usr/ctxtDY_FILENAME").CaretPosition = 9
session.FindById("wnd[1]/tbar[0]/btn[0]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press
session.FindById("wnd[0]/tbar[0]/btn[3]").Press

MsgBox "Scripti valmis ja ladattu P02 Original-kansioon"

End Sub

