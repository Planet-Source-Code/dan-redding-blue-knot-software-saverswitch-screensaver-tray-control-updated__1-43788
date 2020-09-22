VERSION 5.00
Begin VB.Form frmTray 
   ClientHeight    =   540
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   450
   ControlBox      =   0   'False
   Icon            =   "frmTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleWidth      =   450
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgOff 
      Height          =   240
      Left            =   0
      Picture         =   "frmTray.frx":27A2
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgOn 
      Height          =   240
      Left            =   0
      Picture         =   "frmTray.frx":28EC
      Top             =   240
      Width           =   240
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Begin VB.Menu mnunAbout 
         Caption         =   "About SaverSwitch"
      End
      Begin VB.Menu mnuStartup 
         Caption         =   "Load at Startup"
      End
      Begin VB.Menu mnuOpenCP 
         Caption         =   "Open Control Panel"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "Run Screensaver Now"
      End
      Begin VB.Menu mnuDisable 
         Caption         =   "Disable Screensaver"
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

'Boolean for disabled state and long for length of a double-click
Private blnSSDis As Boolean, lDblClick As Long

Private Sub Form_Load()
Dim lIcon As Long, sMess As String
    'quit if already running another copy
    If App.PrevInstance Then
        MsgBox "Only one copy of SaverSwitch may be running at one time on a computer!", vbCritical, App.Title
        Unload frmTray
        Exit Sub
    End If
    
    LoadSettings
    mnuStartup.Checked = blnStartup
    
    'is saver already disabled?
    blnSSDis = Not IsSSEnabled
    If blnSSDis Then
        'Show off
        lIcon = imgOff.Picture
        sMess = DMESS
    Else
        'Show on
        lIcon = imgOn.Picture
        sMess = EMESS
    End If
    'create tray icon
    SetTrayIcon sMess, frmTray.hwnd, lIcon
    'Get max length of time in seconds for a double-click
    lDblClick = GetDoubleClickTime()
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
            ToggleDisabled
        
        Case &H203 'LBUTTONDBLCLICK
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
    SaveSettings
    'Kill the icon
    RemoveTrayIcon frmTray.hwnd
End Sub

Private Sub mnuDisable_Click()
    'toggle through menu
    ToggleDisabled
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
        ModifyTrayIcon DMESS, frmTray.hwnd, imgOff.Picture
    Else
        ModifyTrayIcon EMESS, frmTray.hwnd, imgOn.Picture
    End If
    'synchonize menu checkmark
    mnuDisable.Checked = blnSSDis
End Sub

Private Sub mnunAbout_Click()
    'Hi there!
    MsgBox "SaverSwitch 0.9" & vbCrLf & vbCrLf & _
        "© March 2003" & vbCrLf & _
        "Dan Redding / Blue Knot Software" & vbCrLf & vbTab & _
        "http://www.blueknot.com" & vbCrLf & vbCrLf & _
        "Purpose: Temporarily disable the screensaver with a single" & vbCrLf & _
        "click of a system tray icon; or run it with a double-click." & vbCrLf & vbCrLf & _
        "This is a basic version, I plan to release a full version on" & vbCrLf & _
        "the website later with a lot more features." & vbCrLf & vbCrLf & _
        "But since a good portion of the techniques here were learned on" & vbCrLf & _
        "PSC, I thought it would be nice to 'give back' a little.", _
        vbInformation, "About SaverSwitch"
End Sub

Private Sub mnuOpenCP_Click()
Dim blnSSDisTemp As Boolean, lCP As Long
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
    
    Debug.Print Now
    'Restore disabled if it was originally
    If blnSSDisTemp Then
        SetSSEnabled False
    End If
End Sub

Private Sub mnuRun_Click()
    'Launch current SS
    StartSS
End Sub


'"By giving us the opinions of the uneducated,
'    journalism keeps us in touch with the ignorance of the community."
'                               -- Oscar Wilde

'Heard it on the radio the other day, had to share it...
Private Sub mnuStartup_Click()
    mnuStartup.Checked = Not mnuStartup.Checked
    blnStartup = mnuStartup.Checked
    RunAtStartup blnStartup
End Sub
