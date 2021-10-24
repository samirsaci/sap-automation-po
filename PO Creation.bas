Attribute VB_Name = "Module1"
Option Explicit

'Variables for SAP GUI Tool
Public SapGuiAuto, WScript, msgcol
Public objGui  As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession
Public objSBar As GuiStatusbar
Public objSheet As Worksheet
'Variables for Functions
Public Plant, SAP_CODE, Listing_Procedure As String
Dim W_System
Dim iCtr As Integer

'Transactions Code
Const tcode = "WSM3"

'Function to Connect with SAP GUI Sessions
Function Create_SAP_Session() As Boolean
    'Variables for Session Creation
    Dim il, it
    Dim W_conn, W_Sess, tcode, Transac, Info_System
    Dim N_Gui As Integer
    Dim A1, A2 As String
    'Get Transaction Code
    tcode = Sheets(1).Range("B3")
    'Get System Name in Cell(2,1) of Sheet1
    If mysystem = "" Then
        W_System = Sheets(1).Cells(2, 2)
    Else
        W_System = mysystem
    End If
    'If we are already connected to a Session we exit this function
    If W_System = "" Then
    Create_SAP_Session = False
    Exit Function
    End If
    'If Object Session is not null and the system is matching with the one we target: we use this object
    If Not session Is Nothing Then
        If session.Info.SystemName & session.Info.Client = W_System Then
            Create_SAP_Session = True
            Exit Function
        End If
    End If
    'If we are not connected to anything and GUI Object is Nothing we create one
    If objGui Is Nothing Then
    Set SapGuiAuto = GetObject("SAPGUI")
    Set objGui = SapGuiAuto.GetScriptingEngine
    End If
    'Loop through all SAP GUI Sessions to find the one with the right transaction
    For il = 0 To objGui.Children.Count - 1
        Set W_conn = objGui.Children(il + 0)
        
        For it = 0 To W_conn.Children.Count - 1
            Set W_Sess = W_conn.Children(it + 0)
            Transac = W_Sess.Info.Transaction
            Info_System = W_Sess.Info.SystemName & W_Sess.Info.Client
            
            'Check if Session Name and Transaction Code are matching then connect to it
            If W_Sess.Info.SystemName & W_Sess.Info.Client = W_System Then
            'If W_Sess.Info.SystemName & W_Sess.Info.Client = W_System And W_Sess.Info.Transaction = tcode Then
                Set objConn = objGui.Children(il + 0)
                Set session = objConn.Children(it + 0)
                Exit For
            End If         
        Next 
    Next
    ' If we can't find Session with the right System Name and Transaction Code: display error message
    If session Is Nothing Then
    MsgBox "No active session to system " + W_System + " with transaction " + tcode + ", or scripting is not enabled.", vbCritical + vbOKOnly
    Create_SAP_Session = False
    Exit Function
    End If
    ' Turn on scripting mode
    If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject objGui, "on"
    End If
    'Confirm connection to a session
    Create_SAP_Session = True
End Function

' Function to create the PO
Function PO_Function()

    'Declare Variables
    Dim W_BPNumber, W_SearchTerm, PON
    Dim lineitems As Long
    Dim Sht_Name As String
    Dim N_Lines As Integer
    Dim t As Integer
    Sht_Name = "PO"


    'Launch_Transaction
    session.findById("wnd[0]").Maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me21n"
    session.findById("wnd[0]").sendVKey 0
    Application.Wait (Now + TimeValue("0:00:1"))


    'Fill Vendor Code
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = Sheets(Sht_Name).Cells(2, 3)
    'PurchOrg Code
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").Text = Sheets(Sht_Name).Cells(2, 7)
    'Purch Group
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").Text = Sheets(Sht_Name).Cells(2, 8)
    session.findById("wnd[0]").sendVKey 0


    'Loop for SAP Code
    For t = 0 To N_Lines - 3
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4," & t & "]").Text = Sheets(Sht_Name).Cells(t + 2, 4)
    Next t

    'Loop for Quantities
    For t = 0 To N_Lines - 3
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[" & "6," & t & "]").Text = Sheets(Sht_Name).Cells(t + 2, 6)
    Next t

    'Loop for  Plants
    For t = 0 To N_Lines - 3
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15," & t & "]").Text = Sheets(Sht_Name).Cells(t + 2, 1)
    Next t


    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]").SetFocus
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,1]").caretPosition = 4

    'Click Save Button
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
    session.findById("wnd[0]/sbar").DoubleClick


    'Leave and go to Me23n
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "me23n"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4002/btnDYN_4000-BUTTON").press
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/txtMEPO_TOPLINE-EBELN").SetFocus
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/txtMEPO_TOPLINE-EBELN").caretPosition = 1

    'Copy PO# and paste it in Excel File
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/txtMEPO_TOPLINE-EBELN").SetFocus
    PON = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/txtMEPO_TOPLINE-EBELN").Text
    N_Lines = 1
    While Not (Sheets(Sht_Name).Cells(N_Lines, 1) = "")
        N_Lines = N_Lines + 1
        Sheets(Sht_Name).Cells(N_Lines, 9) = PON
    Wend

    'Leave Ready to go back to Menu
    session.findById("wnd[0]/tbar[0]/btn[3]").press

End Function



