VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTray 
   ClientHeight    =   615
   ClientLeft      =   9480
   ClientTop       =   6795
   ClientWidth     =   615
   ControlBox      =   0   'False
   Icon            =   "frmTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   615
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   30
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":27A2
            Key             =   ""
            Object.Tag             =   "Open Control Panel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":28FC
            Key             =   "OFF"
            Object.Tag             =   "Screensaver Disabled"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":2A56
            Key             =   "ON"
            Object.Tag             =   "Screensaver Enabled"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":2BB0
            Key             =   "NONE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":2D0A
            Key             =   ""
            Object.Tag             =   "Exit SaverSwitch"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":32A4
            Key             =   ""
            Object.Tag             =   "About SaverSwitch"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":33FE
            Key             =   ""
            Object.Tag             =   "Run Screensaver Now"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":3558
            Key             =   ""
            Object.Tag             =   "Load At Startup"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":36B2
            Key             =   ""
            Object.Tag             =   "Not Loaded At Startup"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":380C
            Key             =   ""
            Object.Tag             =   "Email Author"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":3966
            Key             =   ""
            Object.Tag             =   "Visit Website"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":3AC0
            Key             =   ""
            Object.Tag             =   "Current Screensaver"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":3C1A
            Key             =   ""
            Object.Tag             =   "Password Protected"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":3D74
            Key             =   ""
            Object.Tag             =   "No Password"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":3ECE
            Key             =   ""
            Object.Tag             =   "Configure"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTray.frx":4028
            Key             =   ""
            Object.Tag             =   "Start"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Begin VB.Menu zmnuSepWeb 
         Caption         =   "- www.blueknot.com "
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "Visit Website"
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "Email Author"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About SaverSwitch"
      End
      Begin VB.Menu zmnuSepTitle 
         Caption         =   "- S a v e r S w i t c h "
      End
      Begin VB.Menu mnuStartup 
         Caption         =   ""
      End
      Begin VB.Menu mnuRun 
         Caption         =   "Run Screensaver Now"
      End
      Begin VB.Menu mnuOpenCP 
         Caption         =   "Open Control Panel"
      End
      Begin VB.Menu zmnuOptions 
         Caption         =   "- O P T I O N S "
      End
      Begin VB.Menu mnuCurrent 
         Caption         =   "Current Screensaver"
         Begin VB.Menu mnuSaver 
            Caption         =   "(None)"
            Index           =   0
         End
      End
      Begin VB.Menu mnuConfigure 
         Caption         =   "Configure Screensaver..."
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Start after"
         Begin VB.Menu mnuMin 
            Caption         =   "1 minute"
            Index           =   0
         End
         Begin VB.Menu mnuMin 
            Caption         =   "2 minutes"
            Index           =   1
         End
         Begin VB.Menu mnuMin 
            Caption         =   "3 minutes"
            Index           =   2
         End
         Begin VB.Menu mnuMin 
            Caption         =   "5 minutes"
            Index           =   3
         End
         Begin VB.Menu mnuMin 
            Caption         =   "10 minutes"
            Index           =   4
         End
         Begin VB.Menu mnuMin 
            Caption         =   "15 minutes"
            Index           =   5
         End
         Begin VB.Menu mnuMin 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuMin 
            Caption         =   ""
            Index           =   7
         End
      End
      Begin VB.Menu mnuPassword 
         Caption         =   ""
      End
      Begin VB.Menu mnuDisable 
         Caption         =   ""
      End
      Begin VB.Menu zmnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit SaverSwitch"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Boolean for disabled state and long for length of a double-click, path of curr SS, boolean for password
'protected, long for timeout length (minutes)
Private blnSSDis As Boolean, lDblClick As Long, strCurr As String, blnSSSecure As Boolean, lTimeout As Long
Private WithEvents HelpObj As HelpCallBack
Attribute HelpObj.VB_VarHelpID = -1

Private Sub Form_Load()
Dim lIcon As Long, sMess As String
    'quit if already running another copy
    If App.PrevInstance Then
        MsgBox "Only one copy of SaverSwitch may be running at one time on a computer!", vbCritical, App.Title
        Unload frmTray
        Exit Sub
    End If
    
    LoadSettings
    'get current SS
    strCurr = GetCurrSS
    'is saver already disabled?
    blnSSDis = Not IsSSEnabled
    'Is SS password protected
    blnSSSecure = GetSSSecure
    'setup menu items
    SetCaptions
    SetTimeoutMenu
    'create tray icon
    TrayIcon True
    'Get max length of time in seconds for a double-click
    lDblClick = GetDoubleClickTime()
    'load list of screensavers with proper names
    RefreshList
    
    'connect coolmenu icons
    Set HelpObj = New HelpCallBack
    'This is a slightly modified CoolMenu.  The second param says whether to subclass or not
    '(I had another proj that already subclassed the form... form's don't like to be subclassed
    'twice...)
    mCoolMenu.Install frmTray.hwnd, True, HelpObj, ilsMenu
    mCoolMenu.ComplexChecks frmTray.hwnd, False
    
    'no need to see this useless form!
    frmTray.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static blnRunSaver As Boolean, lTime As Long
    'Actually the callback routine from the tray icon
    '(Cheaper than subclassing)
    
    'X is actually 'uMsg'
    'VB multiplied the return value by TwipsPerPixelX for us...
    '(...how thoughtful...)
    Select Case X / Screen.TwipsPerPixelX
        Case &H202 'LBUTTONUP
            'See below for explanation of this section
            If blnRunSaver Then
                blnRunSaver = False
                If GetTickCount() - lTime <= lDblClick Then
                    Exit Sub
                End If
            End If
            'switch SS state
            If strCurr <> "" Then
                ToggleDisabled
            End If
        Case &H203 'LBUTTONDBLCLICK
            If strCurr <> "" Then 'only try to run is SS selected
                'toggle again because LBUTTONUP has already triggered once
                ToggleDisabled
                'Run
                StartSS
                'this is a little odd.  On 2000 then LBUTTONUP doesn't
                'come again after the LBUTTONDBLCLICK, but on XP it does
                '(and maybe others).  So I set a flag to basically say that
                'a double click occured and saved a tickcount.
                'Above, if the flag is on and the tickcount difference
                'is within the system limit for a double-click, it
                'doesn't toggle a third time
                blnRunSaver = True
                lTime = GetTickCount()
            End If
        Case &H205 'RBUTTONUP
            'This neat little trick of setting the foreground window
            'makes the menu actually go away like it supposed to when
            'it loses focus!
            SetForegroundWindow frmTray.hwnd
            'show the tray menu
            PopupMenu mnuPopup
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case vbFormCode 'closing through menu
            'if disabled, ask to restore
            If blnSSDis Then
                'is "reenable" really a word?  Maybe it needs a hyphen...
                If (MsgBox("Reenable Screensaver before closing?", vbYesNo + vbQuestion, "SaverSwitch") = vbYes) Then
                    SetSSEnabled True
                End If
            End If
        Case Else 'closing for other reason (windows exiting, etc.)
            'Always restore
            If blnSSDis Then
                SetSSEnabled True
            End If
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'save the one setting that exists (load at startup)
    SaveSettings
    'Disconnect collmenu
    Call mCoolMenu.Uninstall(Me.hwnd, True)
    Set HelpObj = Nothing
    'Kill the icon
    RemoveTrayIcon frmTray.hwnd
End Sub

Private Sub mnuConfigure_Click()
    'launch the saver throught shell w/ '/C' parameter -- SHOULD
    'open configuration menu if there is one
    Shell """" & strCurr & """ /C", vbNormalFocus
End Sub

Private Sub mnuDisable_Click()
    'toggle through menu
    ToggleDisabled
End Sub

Private Sub mnuEmail_Click()
    'Launch mailto: link
    ShellExecute 0&, vbNullString, "mailto:Dan@blueknot.com?Subject=SaverSwitch%20Feedback", _
        vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub mnuExit_Click()
    'exit program
    Unload frmTray
End Sub

Private Sub ToggleDisabled()
    'toggle the screensaver
    '(Sorry about the negative logic.  The boolean represents 'disabled' but the
    'SetSSEnabled routine is designed to take True for 'Enabled'.  So I set it before
    'I negate it.  I know it's weird, it just makes sense to me.)
    SetSSEnabled blnSSDis
    blnSSDis = Not blnSSDis
    'show appropriate icon
    If blnSSDis Then
        ModifyTrayIcon DMESS, frmTray.hwnd, ilsMenu.ListImages("OFF").Picture
    Else
        ModifyTrayIcon EMESS, frmTray.hwnd, ilsMenu.ListImages("ON").Picture
    End If
    'update caption
    SetCaptions
End Sub

Private Sub mnuAbout_Click()
    'Hi there!
    MsgBox "SaverSwitch 0.99" & vbCrLf & vbCrLf & _
        "© March 2003" & vbCrLf & _
        "Dan Redding / Blue Knot Software" & vbCrLf & vbCrLf & vbTab & _
        "http://www.blueknot.com" & vbCrLf & vbCrLf & _
        "Purpose:" & vbCrLf & "- Temporarily disable the screensaver with a single" & vbCrLf & _
        "click of a system tray icon; or run it with a double-click." & vbCrLf & vbCrLf & _
        "- Select & configure screensaver from tray icon or launch" & vbCrLf & _
        "the control panel to choose with preview." & vbCrLf & vbCrLf & _
        "I plan to release a full version on the website later with" & vbCrLf & _
        "a proper help file and maybe a few more features...", _
        vbInformation, "About SaverSwitch"
End Sub


Private Sub mnuMin_Click(Index As Integer)
Dim strIn As String
    'get the new timeout
    If Index < 7 Then
        'Preset values (1,2,3,5,10,15)
        lTimeout = CLng(Trim$(Left$(mnuMin(Index).Caption, 2)))
    Else
        'input from user
        'could be nicer, this is quickie
        strIn = InputBox("Number of idle minutes before starting screensaver:", "Set Timeout", "30")
        If IsNumeric(strIn) Then
            On Error GoTo NoGood
            lTimeout = CLng(strIn)
        Else
            Exit Sub
        End If
    End If
    'set the timeout and update the menu
    SetSSTimeout lTimeout
    SetTimeoutMenu
NoGood:
End Sub

Private Sub mnuOpenCP_Click()
Dim blnSSDisTemp As Boolean, lCP As Long, iLoop As Integer
    blnSSDisTemp = blnSSDis
    'If currently disabled, enable it.  Otherwise panel
    'may show '(None)' for screen saver selection
    If blnSSDis Then
        'disable but don't change icon
        SetSSEnabled True
    End If
    'launch the Display control panel, second tab (SS)
    Shell "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,1"
    
    'Wait for the control panel window to appear
    lCP = 0&
    Do While lCP = 0
        lCP = FindWindow(vbNullString, "Display Properties")
        DoEvents
    Loop
        
    'Wait for the control panel window to disappear
    Do While IsWindow(lCP)
        DoEvents
    Loop
    
    'Restore disabled if it was originally
    If blnSSDisTemp Then
        SetSSEnabled False
    End If
    'update the menu items to reflect any changes
    strCurr = GetCurrSS
    For iLoop = 0 To mnuSaver.UBound
        mnuSaver(iLoop).Checked = (UCase$(mnuSaver(iLoop).Tag) = UCase$(strCurr))
    Next iLoop
    SetCaptions
    SetTimeoutMenu
    TrayIcon
End Sub

Private Sub mnuPassword_Click()
    'toggle password protection
    SetSSSecure (Not blnSSSecure)
    blnSSSecure = GetSSSecure
    SetCaptions
End Sub

Private Sub mnuRun_Click()
    'Launch current SS
    StartSS
End Sub

Private Sub mnuSaver_Click(Index As Integer)
Dim iLoop As Integer
    'change check
    For iLoop = 0 To mnuSaver.UBound
        mnuSaver(iLoop).Checked = iLoop = Index
    Next iLoop
    'set the new saver
    SetCurrSS mnuSaver(Index).Tag
    strCurr = GetCurrSS
    'update tray icon in case SS changed to or from '(None)'
    TrayIcon
End Sub

Private Sub mnuStartup_Click()
    'load at startup setting
    blnStartup = Not blnStartup
    SetCaptions
    RunAtStartup blnStartup
End Sub


'"By giving us the opinions of the uneducated,
'    journalism keeps us in touch with the ignorance of the community."
'                               -- Oscar Wilde

'Heard it on the radio the other day, had to share it...

Private Sub mnuWebsite_Click()
    'launch website link in default browser
    ShellExecute 0&, vbNullString, "http://www.blueknot.com", _
        vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub RefreshList()
Dim lFolder As Long, strDir As String, strFolder As String, iLoop As Integer, _
    strPath As String, strFile As String, strName As String, blnNone As Boolean
    'wipe the list except for '(None)'
    For iLoop = mnuSaver.UBound To 1 Step -1
        Unload mnuSaver(iLoop)
    Next iLoop
    mnuSaver(0).Checked = False
    'wipe the memory list
    Erase SSList
    'system folder first, then windows folder
    For lFolder = 37& To 36& Step -1
        'get the real system & windows folder, don't assume "C:\Windows"
        'Thanks to Mr. Bobo for this beauty
        strFolder = SpecialFolder(lFolder)
        'find all the .scr's
        strDir = Dir$(getFullPath(strFolder, "*.scr"), vbNormal + vbHidden)
        Do While strDir <> ""
            'dir just returns the file, this rebuilds the full path
            strPath = getFullPath(strFolder, strDir)
            'get the proper display name for the SS
            strName = GetSSName(strPath, strDir)
            strFile = strDir
            'add to list if not already on the list (file, not path or display name)
            'in case copy of SS in windows and system folders
            If Not CheckSSFile(strFile) Then
                AddSSFile strPath, strFile, strName
            End If
            'get next
            strDir = Dir$
        Loop
    Next lFolder
    'simple bubble sort
    SortSS
    'load the menus
    On Error GoTo NoSS
    'if none found, UBound may generate an error
    blnNone = True
    For iLoop = 1 To UBound(SSList)
        Load mnuSaver(iLoop)
        mnuSaver(iLoop).Visible = True
        mnuSaver(iLoop).Caption = "#| |" & Replace$(SSList(iLoop).Name, "&", "&&")
        mnuSaver(iLoop).Tag = SSList(iLoop).Path
        If UCase$(mnuSaver(iLoop).Tag) = UCase$(strCurr) Then
            mnuSaver(iLoop).Checked = True
            blnNone = False
        End If
    Next iLoop
NoSS:
    If blnNone Then mnuSaver(0).Checked = True
End Sub

Private Sub TrayIcon(Optional blnCreate As Boolean = False)
Dim lIcon As Long, sMess As String
    'set the appropriate tray icon & tooltip
    If strCurr <> "" Then
        If blnSSDis Then
            'Show off
            lIcon = ilsMenu.ListImages("OFF").Picture
            sMess = DMESS
        Else
            'Show on
            lIcon = ilsMenu.ListImages("ON").Picture
            sMess = EMESS
        End If
        mnuDisable.Enabled = True
    Else
        lIcon = ilsMenu.ListImages("NONE").Picture
        sMess = NMESS
        mnuDisable.Enabled = False
    End If
    'these menus are only enabled if a screensaver is selected
    mnuRun.Enabled = mnuDisable.Enabled
    mnuStart.Enabled = mnuDisable.Enabled
    mnuPassword.Enabled = mnuDisable.Enabled
    mnuConfigure.Enabled = mnuDisable.Enabled
    'add or modify the tray icon
    If blnCreate Then
        SetTrayIcon sMess, frmTray.hwnd, lIcon
    Else
        ModifyTrayIcon sMess, frmTray.hwnd, lIcon
    End If
End Sub

Private Sub SetCaptions()
    'set appropriate menu captions
    'This also changes the associated icon from the imagelist
    mnuStartup.Caption = IIf(blnStartup, "Load", "Not Loaded") & " At Startup"
    mnuDisable.Caption = "Screensaver " & IIf(blnSSDis, "Dis", "En") & "abled"
    mnuPassword.Caption = IIf(blnSSSecure, "Password Protected", "No Password")
End Sub

Private Sub SetTimeoutMenu()
Dim iIndex As Integer, iLoop As Integer
    'set the correct checkmark and caption on other menu
    lTimeout = GetSSTimeout
    Select Case lTimeout
        Case 1, 2, 3
            iIndex = lTimeout - 1
        Case 5
            iIndex = 3
        Case 10
            iIndex = 4
        Case 15
            iIndex = 5
        Case Else
            iIndex = 7
    End Select
    For iLoop = 0 To mnuMin.UBound
        mnuMin(iLoop).Checked = iLoop = iIndex
    Next iLoop
    If iIndex = 7 Then
        mnuMin(7).Caption = "#| |Other (" & CStr(lTimeout) & " minutes)..."
    Else
        mnuMin(7).Caption = "#| |Other..."
    End If
End Sub
