VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "HLRConsole v1.11"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10035
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   10035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPager 
      BackColor       =   &H00000000&
      Caption         =   "Pager "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   6720
      TabIndex        =   10
      Top             =   5520
      Width           =   3255
      Begin VB.Timer tmrStatus 
         Enabled         =   0   'False
         Left            =   2640
         Top             =   720
      End
      Begin VB.Timer tmrPager 
         Interval        =   4000
         Left            =   2160
         Top             =   720
      End
      Begin VB.CommandButton cmdClearPager 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   3240
         Width           =   3015
      End
      Begin MSWinsockLib.Winsock wsStatus 
         Left            =   2640
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsLog 
         Left            =   2160
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsRCon 
         Left            =   1680
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblPager 
         BackColor       =   &H00000000&
         Caption         =   "lblPager"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2895
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Frame frmStatus 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   9840
      Width           =   9015
      Begin VB.Label lblPortStatus 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Status Port: [00000]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4680
         TabIndex        =   15
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblServerStatus 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Server: [000.000.000.000:00000]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   120
         Width           =   2910
      End
      Begin VB.Label lblPagerStatus 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Pager: [OFF]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox txtCommand 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "txtCommand"
      Top             =   9480
      Width           =   7575
   End
   Begin VB.Frame frmRCon 
      BackColor       =   &H00000000&
      Caption         =   "RCon "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   0
      TabIndex        =   5
      Top             =   5520
      Width           =   6615
      Begin VB.TextBox txtRCon 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "frmMain.frx":0A02
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame frmChat 
      BackColor       =   &H00000000&
      Caption         =   "Chat "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   0
      TabIndex        =   4
      Top             =   3480
      Width           =   9975
      Begin VB.TextBox txtChat 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "frmMain.frx":0A0C
         Top             =   240
         Width           =   9735
      End
   End
   Begin VB.Frame frmMisc 
      BackColor       =   &H00000000&
      Caption         =   "Log "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9975
      Begin VB.TextBox txtMisc 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "frmMain.frx":0A16
         Top             =   240
         Width           =   9735
      End
   End
   Begin VB.Label lblRCon 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "RCON"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   9480
      Width           =   510
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuEditINI 
         Caption         =   "Edit HLRCONSOLE.&INI"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset Console"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ByteCode = "ÿÿÿÿ"            'Prefix sent before all UDP packets
Const Quote = """"                 'The quote "
Const MaxChar = 30000              'Max characters to store in a textbox
Const ProgTitle = "HLRConsole"     'Title of the program
Const VersionNum = "v1.11"         'Program version number
Const Spacing = 10                 'Space between controls
Const WidthBias = 150              'Width of the window borders
Const HeightBias = 800             'Height of the window borders
Const MinWidth = 10155             'Minimum Window Width
Const MinHeight = 10785            'Minimum Window Height
Const LabelON = "Pager: [ON]"      'Text to display when the pager is on
Const LabelOFF = "Pager: [OFF]"    'Text to display when the pager is off

Dim ProgPath As String             'Path the EXE was started from
Dim INIFile As String              'Path/File of the INI file
Dim ThemeFile As String            'Path/File to theme INI file
Dim LocalIP As String              'Local IP Address
Dim ServerIP As String             'Server IP Address
Dim ServerPort As Integer          'Game Server Port Number
Dim Password As String             'Server RCon Password
Dim LogPort As Integer             'Port for remote logging
Dim RConPort As Integer            'Port for sending RCon commands
Dim StatusPort As Integer          'Port for retreiving server status
Dim RefreshInterval As Integer     'Interval (in ms) to check server status
Dim PageWAV As String              'WAV used when paging the admin
Dim PageON As String               'Text to display when paging is on
Dim PageOFF As String              'Text to display when paging is off
Dim PageAWAY As String             'Text to display when page is missed
Dim Challenge As String            'Stores the challenge response code
Dim PageCount As Integer           'Counts the number of times the WAV is played
Dim PagerDefault As String         'Does the pager start ON or OFF?
Dim CmdHistory(1 To 20) As String  'Sent command history
Dim CmdCurrent As Integer          'Currently selected command in history

Private Enum enLogType  'Used in LogItem subroutine
    ltMisc = 0          'Log to the txtMisc box
    ltChat = 1          'Log to the txtChat box
    ltRCon = 2          'Log to the txtRCon box
End Enum

Private Sub cmdClearPager_Click()
    tmrPager.Enabled = False   'Disable the pager sound
    lblPager.Caption = "No Pages"   'Clear the pager text
    txtCommand.SetFocus   'Move the cursor back to the Command box
End Sub

Private Sub cmdReset_Click()
    wsRCon.Close   'Close all Winsock connections
    wsLog.Close
    wsStatus.Close
    Form_Load    'Load the INI, reset the WinSock, and reconnect
    txtCommand.SetFocus   'Move the cursor back to the Command box
End Sub

Private Sub Form_Load()
    'Initialize variables
    ProgPath = App.Path
    If Right(App.Path, 1) <> "\" Then ProgPath = ProgPath & "\"
    INIFile = ProgPath & "HLRConsole.ini"
    ThemeFile = ProgPath & "Theme.ini"
    Challenge = ""
    PageCount = 0
    
    'Initialize controls
    txtMisc.Text = ""
    txtChat.Text = ""
    txtRCon.Text = ""
    txtCommand.Text = ""
    lblPager.Caption = "No Pages"
    tmrPager.Enabled = False
    CmdCurrent = 0
        
    'Show the window
    Me.Show
    Me.Caption = "Initializing... - " & ProgTitle
    
    'Log startup progress
    LogItem ProgTitle & " " & VersionNum
    LogItem String(80, "-")
    LogItem "INI File = " & INIFile
    LogItem "Reading INI..."
    LoadINI INIFile
    LoadTheme ThemeFile
    frmMisc.Caption = "Log [Port " & LogPort & "] "
    frmRCon.Caption = "RCon [Port " & RConPort & "] "
    lblServerStatus = "Server: [" & ServerIP & ":" & ServerPort & "]"
    lblPortStatus = "Status Port: [" & StatusPort & "]"
    lblServerStatus.Left = lblPagerStatus.Left + lblPagerStatus.Width + 500
    lblPortStatus.Left = lblServerStatus.Left + lblServerStatus.Width + 500
    If PagerDefault = "ON" Then
        lblPagerStatus = LabelON
    Else
        lblPagerStatus = LabelOFF
    End If
    tmrStatus.Interval = RefreshInterval
    tmrStatus.Enabled = True
    LogItem "Done!" & vbCrLf
    
    'Configure WinSock
    LogItem "Configuring WinSock and Binding ports..."
    BindWinsock
    LogItem "Done!" & vbCrLf

    'Send the RCON Challenge
    LogItem "Attempting to challenge RCon..."
    wsRCon.SendData ByteCode & "challenge rcon"

End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub   'Don't change controls if minimized
    If Me.Width < MinWidth Then Me.Width = MinWidth   'Sets minimum form size
    If Me.Height < MinHeight Then Me.Height = MinHeight
    
    ResizeControls
End Sub

Private Sub BindWinsock()
    'Configure WinSock and bind to local ports
    With wsLog
        .Protocol = sckUDPProtocol
        .Bind LogPort
        If LocalIP = "0" Then LocalIP = .LocalIP
    End With
    With wsRCon
        .Protocol = sckUDPProtocol
        .RemoteHost = ServerIP
        .RemotePort = ServerPort
        .Bind RConPort
    End With
    With wsStatus
        .Protocol = sckUDPProtocol
        .RemoteHost = ServerIP
        .RemotePort = ServerPort
        .Bind StatusPort
    End With
End Sub

Private Sub LogItem(sData As String, Optional Destination As enLogType = 0)
    'Add a new line to a text box
    Select Case Destination
        Case ltMisc
            With txtMisc
                If Len(.Text) > MaxChar Then
                    .Text = Mid(.Text, InStr(100 + Len(sData), .Text, vbCrLf) + 2)
                End If
                .Text = .Text & sData & vbCrLf
                .SelStart = Len(.Text)
            End With
            
        Case ltChat
            With txtChat
                If Len(.Text) > MaxChar Then
                    .Text = Mid(.Text, InStr(100 + Len(sData), .Text, vbCrLf) + 2)
                End If
                .Text = .Text & sData & vbCrLf
                .SelStart = Len(.Text)
            End With
        
        Case ltRCon
            With txtRCon
                If Len(.Text) > MaxChar Then
                    .Text = Mid(.Text, InStr(100 + Len(sData), .Text, vbCrLf) + 2)
                End If
                .Text = .Text & sData & vbCrLf
                .SelStart = Len(.Text)
            End With
    End Select
    
End Sub

Private Sub InsertCommand(cData As String)
    'Insert a command in the history
    CmdCurrent = 0
    
    For A = 19 To 1 Step -1   'Make the most recent first
        CmdHistory(A + 1) = CmdHistory(A)
    Next A
    
    CmdHistory(1) = cData
End Sub

Private Sub ResizeControls()
    'Moves and resizes controls as the window size changes
    With frmStatus
        .Left = 0
        .Width = Me.Width - WidthBias
        .Top = Me.Height - .Height - HeightBias
    End With
    
    With cmdReset
        .Top = frmStatus.Top - txtCommand.Height - Spacing
        .Left = Me.Width - .Width - WidthBias
        lblRCon.Top = .Top
        txtCommand.Top = .Top
    End With
    
    With txtCommand
        .Width = cmdReset.Left - .Left - (Spacing * 5)
    End With
    
    With frmRCon
        .Top = txtCommand.Top - .Height - (Spacing * 5)
        frmPager.Top = .Top
    End With
    
    With frmPager
        .Left = Me.Width - frmPager.Width - WidthBias
        frmRCon.Width = .Left - Spacing
    End With
    
    With frmChat
        .Top = frmRCon.Top - .Height - Spacing
        .Width = Me.Width - WidthBias
    End With
    
    With frmMisc
        .Height = frmChat.Top - Spacing
        .Width = frmChat.Width
    End With
    
    With txtRCon
        .Left = 100
        .Top = 225
        .Height = frmRCon.Height - 350
        .Width = frmRCon.Width - 200
    End With
    
    With txtChat
        .Left = 100
        .Top = 225
        .Height = frmChat.Height - 350
        .Width = frmChat.Width - 200
    End With
    
    With txtMisc
        .Left = 100
        .Top = 225
        .Height = frmMisc.Height - 350
        .Width = frmMisc.Width - 200
    End With
    
    Me.Refresh
End Sub

Private Sub LoadTheme(INI As String)
    Dim TempColor As Long
    
    TempColor = QBColor(INIGetSettingInteger("Theme", "Background", INI))
    Me.BackColor = TempColor
    frmMisc.BackColor = TempColor
    frmChat.BackColor = TempColor
    frmRCon.BackColor = TempColor
    frmStatus.BackColor = TempColor
    frmPager.BackColor = TempColor
    txtMisc.BackColor = TempColor
    txtChat.BackColor = TempColor
    txtRCon.BackColor = TempColor
    txtCommand.BackColor = TempColor
    lblRCon.BackColor = TempColor
    lblPager.BackColor = TempColor
    lblPagerStatus.BackColor = TempColor
    lblServerStatus.BackColor = TempColor
    lblPortStatus.BackColor = TempColor
    
    TempColor = QBColor(INIGetSettingInteger("Theme", "Frame", INI))
    frmMisc.ForeColor = TempColor
    frmChat.ForeColor = TempColor
    frmRCon.ForeColor = TempColor
    frmPager.ForeColor = TempColor
    frmStatus.ForeColor = TempColor
    
    TempColor = QBColor(INIGetSettingInteger("Theme", "Log", INI))
    txtMisc.ForeColor = TempColor
    
    TempColor = QBColor(INIGetSettingInteger("Theme", "Chat", INI))
    txtChat.ForeColor = TempColor
    
    TempColor = QBColor(INIGetSettingInteger("Theme", "RCon", INI))
    txtRCon.ForeColor = TempColor
    txtCommand.ForeColor = TempColor
    lblRCon.ForeColor = TempColor
    
    TempColor = QBColor(INIGetSettingInteger("Theme", "Pager", INI))
    lblPager.ForeColor = TempColor
    
    TempColor = QBColor(INIGetSettingInteger("Theme", "Status", INI))
    lblPagerStatus.ForeColor = TempColor
    lblServerStatus.ForeColor = TempColor
    lblPortStatus.ForeColor = TempColor
    
End Sub

Private Sub LoadINI(INI As String)
    'Load the INI file into variables
    LocalIP = INIGetSettingString("Main", "LocalIP", INI)
    ServerIP = INIGetSettingString("Main", "ServerIP", INI)
    ServerPort = INIGetSettingString("Main", "ServerPort", INI)
    Password = INIGetSettingString("Main", "RConPass", INI)
    LogPort = INIGetSettingInteger("Main", "LogPort", INI)
    RConPort = INIGetSettingInteger("Main", "RConPort", INI)
    StatusPort = INIGetSettingInteger("Main", "StatusPort", INI)
    RefreshInterval = INIGetSettingInteger("Main", "StatusInterval", INI)
    PageWAV = INIGetSettingString("Main", "PagerFile", INI)
    
    PageON = INIGetSettingString("Main", "PagerOn", INI)
    If PageON = "" Then PageON = "Paging the Admin..."
    
    PageOFF = INIGetSettingString("Main", "PagerOff", INI)
    If PageOFF = "" Then PageOFF = "Paging is disabled.  A message will be logged for the Admin."
    
    PageAWAY = INIGetSettingString("Main", "PagerAway", INI)
    If PageAWAY = "" Then PageAWAY = "Sorry but the Admin is away.  A message will be logged."
    
    PagerDefault = UCase(INIGetSettingString("Main", "PagerDefault", INI))
    If PagerDefault <> "ON" And PagerDefault <> "OFF" Then PagerDefault = "OFF"
End Sub

Private Sub lblPagerStatus_DblClick()
    'Toggle the pager on and off
    With lblPagerStatus
        If .Caption = LabelON Then
            .Caption = LabelOFF
        Else
            .Caption = LabelON
        End If
    End With
End Sub

Private Sub mnuAbout_Click()
MsgBox "HLRConsole v1.11b. (c) Fear-Otaku Software.  This program is provided 'AS-IS'."
End Sub

Private Sub mnuEditINI_Click()
    MsgBox "Click RESET when done editing the INI file for changes to take effect.", , "Info"
    Shell "notepad " & INIFile, vbNormalFocus
End Sub



Private Sub mnuExit_Click()
    Select Case MsgBox("Are you sure you want to exit?", vbYesNo, "Quitting?")
        Case vbYes
            End
    End Select
End Sub

Private Sub mnuReset_Click()
    cmdReset_Click
End Sub

Private Sub tmrPager_Timer()
    'Increment the counter of times the WAV is played and play the WAV
    PageCount = PageCount + 1
    PlayWav ProgPath & PageWAV
    If PageCount = 6 Then
        tmrPager.Enabled = False
        PageCount = 0
        SendRCon "say " & PageAWAY   'Page wasn't answered
    End If
End Sub

Private Sub tmrStatus_Timer()
    'Send 'info' request to server
    wsStatus.SendData ByteCode & "info" & Chr(0)
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then   'Enter pressed
        KeyAscii = 0
        If txtCommand.Text = "" Then Exit Sub   'If no command entered then do nothing
        SendRCon txtCommand.Text   'Send the command to WinSock
        InsertCommand txtCommand.Text   'Add command to history
        txtCommand.Text = ""
    End If
End Sub

Private Sub txtCommand_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38   'Up arrow (get previous command)
            CmdCurrent = CmdCurrent + 1
            If CmdCurrent > 20 Then
                CmdCurrent = 0
                txtCommand.Text = ""
            Else
                txtCommand.Text = CmdHistory(CmdCurrent)
                If txtCommand.Text = "" Then CmdCurrent = 0
            End If
            txtCommand.SelStart = Len(txtCommand.Text)   'Put the cursor at the end of the line
        
        Case 40   'Down arrow (get next command)
            CmdCurrent = CmdCurrent - 1
            If CmdCurrent <= 0 Then
                CmdCurrent = 0
                txtCommand.Text = ""
            Else
                txtCommand.Text = CmdHistory(CmdCurrent)
                If txtCommand.Text = "" Then CmdCurrent = 0
            End If
            txtCommand.SelStart = Len(txtCommand.Text)   'Put the cursor at the end of the line
    End Select
    
End Sub

Private Sub txtMisc_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtRCon_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub wsLog_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo SkipPacket
    
    If Challenge = "" Then Exit Sub  'Do nothing if we're not authorized
    Dim Buffer As String
    wsLog.GetData Buffer
    Buffer = Replace(Buffer, ByteCode, "")
    Buffer = Replace(Buffer, Chr(10) & Chr(0), "")
    Buffer = Replace(Buffer, Chr(0), "")
    Buffer = Replace(Buffer, Chr(10), vbCrLf)
    Buffer = Right(Buffer, Len(Buffer) - 6)
    LogItem Buffer
    
    If InStr(LCase(Buffer), Quote & " say " & Quote) Or _
       InStr(LCase(Buffer), Quote & " say_team " & Quote) Or _
       InStr(LCase(Buffer), "server say " & Quote) Then
        LogItem Buffer, ltChat
    End If
    
    If InStr(Buffer, ": Rcon: " & Quote) Then
        LogItem Buffer, ltRCon
    End If
    
    If InStr(LCase(Buffer), Quote & "page admin") Then
        lblPager.Caption = Buffer
        Select Case lblPagerStatus
            Case LabelON
                SendRCon "say " & PageON
                PageCount = 1
                tmrPager.Enabled = True
                PlayWav ProgPath & PageWAV
            Case LabelOFF
                SendRCon "say " & PageOFF
        End Select
    End If
    
    If InStr(LCase(Buffer), Quote & "admin status") Or _
       InStr(LCase(Buffer), Quote & "pager status") Then
        Select Case lblPagerStatus
            Case LabelON
                SendRCon "say The pager is ON"
            Case LabelOFF
                SendRCon "say The pager is OFF"
        End Select
    End If
    
Exit Sub

SkipPacket:
    Exit Sub
End Sub

Private Sub wsRCon_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo SkipPacket
    
    Dim Buffer As String
    wsRCon.GetData Buffer
    Buffer = Replace(Buffer, ByteCode, "")
    Buffer = Replace(Buffer, Chr(10) & Chr(0), "")
    Buffer = Replace(Buffer, Chr(0), "")
    Buffer = Replace(Buffer, Chr(10), vbCrLf)
    
    If Challenge = "" Then
        Challenge = Trim(Right(Buffer, Len(Buffer) - 14))
        LogItem "Received challenge response of: " & Challenge & vbCrLf
        LogItem "Setting LogAddress..." & vbCrLf
        LogItem "Console started!"
        Me.Caption = "Waiting for status update... - " & ProgTitle
        LogItem "============================="
        SendRCon "logaddress " & LocalIP & " " & LogPort
    Else
        Buffer = Right(Buffer, Len(Buffer) - 1)
        LogItem Buffer, ltRCon & vbCrLf & vbCrLf
    End If
Exit Sub

SkipPacket:
    Exit Sub
End Sub

Private Sub wsStatus_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo SkipPacket
    
    If Challenge = "" Then Exit Sub  'Do nothing if we're not authorized
    Dim Buffer As String
    Dim TempIP As String
    Dim TempName As String
    Dim TempMap As String
    Dim TempGame As String
    Dim TempDesc As String
    Dim TempRemain As String
    Dim TempClients As String
    Dim TempMaxclients As String
    Dim TempProtocolVer As String
    
    wsStatus.GetData Buffer
    Buffer = Replace(Buffer, ByteCode, "")
    
    'Parse the received info packet
    If UCase(Left(Buffer, 1)) = "C" Then
        infoArray = Split(Buffer, Chr(0))
        TempIP = Right(infoArray(0), Len(infoArray(0)) - 1)
        TempName = infoArray(1)
        TempMap = infoArray(2)
        TempGame = infoArray(3)
        TempDesc = infoArray(4)
        TempRemain = infoArray(5)
        TempClients = Asc(Mid(TempRemain, 1, 1))
        TempMaxclients = Asc(Mid(TempRemain, 2, 1))
        TempProtocolVer = Asc(Mid(TempRemain, 3, 1))
        Me.Caption = TempMap & " - (" & TempClients & "/" & TempMaxclients & ") - " & ProgTitle
    End If
Exit Sub

SkipPacket:
    Exit Sub
End Sub

Private Sub SendRCon(cData As String)
    'Send an RCon command to the server
    LogItem vbCrLf & ">> " & cData & " <<", ltRCon
    wsRCon.SendData ByteCode & "rcon " & Challenge & " " & Quote & Password & Quote & " " & cData
End Sub
