VERSION 5.00
Begin VB.Form frmListner 
   BorderStyle     =   0  'None
   Caption         =   "Listener"
   ClientHeight    =   4950
   ClientLeft      =   150
   ClientTop       =   810
   ClientWidth     =   7710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   2505
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   2280
      Width           =   7515
   End
   Begin VB.Frame Frame2 
      Caption         =   "System"
      Height          =   1155
      Left            =   120
      TabIndex        =   7
      Top             =   1050
      Width           =   3885
      Begin VB.CheckBox chkAllowStopService 
         Caption         =   "Allow Stop Service"
         Enabled         =   0   'False
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   780
         Width           =   1815
      End
      Begin VB.CheckBox chkStartService 
         Caption         =   "Allow Start Service"
         Enabled         =   0   'False
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   510
         Width           =   1815
      End
      Begin VB.CheckBox chkLoggedInUser 
         Caption         =   "Logged in user"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   3885
      Begin VB.TextBox txtSubject 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Text            =   "SMS From 919768292949"
         Top             =   570
         Width           =   2415
      End
      Begin VB.TextBox txtInterval 
         Height          =   315
         Left            =   1350
         TabIndex        =   3
         Text            =   "1"
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Check Subject"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   6
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Second(s)"
         Height          =   195
         Index           =   1
         Left            =   1950
         TabIndex        =   4
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Check Interval"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.PictureBox pichook 
      Height          =   555
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   2820
      Top             =   2220
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   690
      Top             =   2160
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   1500
      Picture         =   "Listner.frx":0000
      Top             =   2955
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   2055
      Picture         =   "Listner.frx":030A
      Top             =   2940
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   2580
      Picture         =   "Listner.frx":0614
      Top             =   2955
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Listener"
      Begin VB.Menu mnuSetup 
         Caption         =   "Set up"
      End
      Begin VB.Menu mnuMM 
         Caption         =   "Minimise"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmListner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`*
'Author     : Ameetkumar
'Description: This application used to control the PC using
'             mobile phone through SMS.
''*`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`**`*

'Option Explicit
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim TrayI As NOTIFYICONDATA

Dim sParam As String
Dim sAlertedMails As String

Private Sub Form_Load()
On Error GoTo ErrTrap
'
'Purpose    : Load this application into systray
'
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd 'Link the trayicon to this picturebox
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = imgIcon(2).Picture
    TrayI.szTip = "Listener 0.1.1" & Chr$(0)
    'Create the icon
    Shell_NotifyIcon NIM_ADD, TrayI
    Me.Hide
    sAlertedMails = ""
    Exit Sub
ErrTrap:
    txtLog = txtLog & "Error In Form_Load(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Sub

Private Sub mnuExit_Click()
On Error GoTo ErrTrap
'
'Purpose    : Load this application into systray
'
    Unload Me
    Exit Sub
ErrTrap:
    txtLog = txtLog & "Error In mnuExit_Click(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Sub

Private Sub mnuMM_Click()
    'Minimise Form
    Timer2.Interval = Val(Trim(txtInterval.Text)) * 1000
    Me.Hide
End Sub

Private Sub mnuSetup_Click()
    Me.Show
End Sub

Private Sub Timer2_Timer()
On Error GoTo ErrTrap
'
'Purpose    This procedure will check every minute if any mails sent from my mobile
'           If any mails sent then the system will act accordingly
'
Dim iDir As Integer
Dim iFile As Integer
Dim sParam1 As String
Dim sParam2 As String
Dim sParam3 As String
Dim sStr As String
Dim i As Integer
    Select Case (UCase(Trim(ParseMail)))
        Case "SHUTDOWN"
            '"SHUTDOWN"
            Shell "shutdown -s -t 10"
        Case "LOGOFF"
            '"LOGOFF"
            Shell "shutdown -l -t 10"
        Case "RESTART"
            '"RESTART"
            Shell "shutdown -r -t 10"
        Case "RUN"
            '"RUN~c:\test.exe"
            Shell sParam
        Case "SENDFILE"
            '"SENDFILE~c:\test.txt~MTO"
            sParam1 = Mid(sParam, 1, InStr(1, sParam, "~") - 1)
            sParam2 = Mid(sParam, InStr(1, sParam, "~") + 1)
            SendMail sParam2, "Send File " & Now, sParam1, "test"
        Case "WHO"
            'To get logged in user
            If chkLoggedInUser.Value = 1 Then
                sStr = LoggedInUser
                SendMail "919768292949.ameet_kumar2003@smscountry.net", "who", "", "User logged in : " & sStr
            End If
        Case "CHECKMAIL"
               SendMail "919768292949.ameet_kumar2003@smscountry.net", "checkmail", "", "U have " & CheckMail & " Unread Mails"
        Case "READMAILHEADER"
            ReadMailHeader
        Case "READMSG"
            ReadMail sParam
        Case "CREATE_FOLDER"
            CreateFolder sParam
        Case "DELETE_FOLDER"
            DeleteFolderWithSubFolders sParam
        Case "DELETE_FILES"
            DeleteFiles sParam
        Case "MOVE_FOLDER"
            sParam1 = Mid(sParam, 1, InStr(1, sParam, "~") - 1)
            sParam2 = Mid(sParam, InStr(1, sParam, "~") + 1)
            MoveFolder sParam1, sParam2
        Case "COPY_FOLDER"
            sParam1 = Mid(sParam, 1, InStr(1, sParam, "~") - 1)
            sParam2 = Mid(sParam, InStr(1, sParam, "~") + 1)
            CopyFolder sParam1, sParam2
        Case "CREATE_TEXT_FILE"
            sParam1 = Mid(sParam, 1, InStr(1, sParam, "~") - 1)
            sParam2 = Mid(sParam, InStr(1, sParam, "~") + 1)
            CreateTextFile sParam1, sParam2
        Case "PC_CONFIG"
            GetSystemConfig
        
    End Select
    Exit Sub
ErrTrap:
    txtLog = txtLog & "Error In ParseMail(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Sub

Private Sub mnuPop_Click(Index As Integer)
On Error GoTo ErrTrap
'
'Purpose    : Pop-up menu actions wre handled here
'
    'Me.Show
    Select Case Index
        Case 0  'About
            MsgBox "All rights reserved." + vbCrLf + "E-Mail: projtest21012@gmail.com", vbInformation + vbOKOnly
        Case 2  'End
            Unload Me
    End Select
    Exit Sub
ErrTrap:
    txtLog = txtLog & "Error In ParseMail(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Sub

Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrTrap
'
'Purpose    : This function is called when Double click or right
'             mouse button clicked in Notification Icon (System tray)
'
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then  'If the user dubbel-clicked on the icon
        mnuPop_Click 0
    ElseIf Msg = WM_RBUTTONDOWN Then  'Right click
        Me.PopupMenu mnuPopUp
    End If
    Exit Sub
ErrTrap:
    txtLog = txtLog & "Error In ParseMail(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrTrap
'
'Purpose    : Remove notification Icon From the systray
'
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd
    TrayI.uId = 1&
    'Delete the icon
    Shell_NotifyIcon NIM_DELETE, TrayI
    End
    Exit Sub
ErrTrap:
    txtLog = txtLog & "Error In ParseMail(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Sub
Private Sub Timer1_Timer()
'
'Purpose    : To Animate Notification icon
'
    Static mPic As Integer
    Me.Icon = imgIcon(mPic).Picture
    TrayI.hIcon = imgIcon(mPic).Picture
    mPic = mPic + 1
    If mPic = 3 Then mPic = 0
    Shell_NotifyIcon NIM_MODIFY, TrayI
End Sub

Private Function ParseMail() As String
On Error GoTo ErrTrap
'
'Purpose      This procedure will check and let us know what action to be
'             taken aginst the command sent from mobile
'
'Set up our variables
Dim oApp        As Outlook.Application  'Create an object for outlook application
Dim oNpc        As NameSpace            'Name space to drildown message folder
Dim oMails      As MailItem             'To find our mail
Dim sCommand    As String
Dim iMsgCount   As Integer
Dim sMsgHead    As String
    'Lets apply values to our variables
    Set oApp = CreateObject("Outlook.Application")
    Set oNpc = oApp.GetNamespace("MAPI")
    iMsgCount = 0
    'Lets iterate through an easy For Each loop
    For Each oMails In oNpc.GetDefaultFolder(olFolderInbox).Items
        If oMails.UnRead Then
            sParam = ""
            'Change the Subject comparition string based on your service provider message
            If UCase(oMails.Subject) = UCase(Trim(txtSubject.Text)) Then
                'For visualgsm.com as the service provider.
                'Dim startIndex As String
                'Dim endIndex As String
                'startIndex = InStr(1, oMails.Body, "@") + 11
                'endIndex = InStr(startIndex, oMails.Body, "^")
                'sCommand = Mid(oMails.Body, startIndex, endIndex - startIndex)
                'For general e-mail.
                'sCommand = Mid(oMails.Body, 1, InStr(1, oMails.Body, "^") - 1)
                'For smscountry.com as the service provider.
                'sCommand = oMails.Body
                'sCommand = Mid(oMails.HTMLBody, 4, InStr(1, oMails.HTMLBody, "</p>") - 4)
                'sCommand = Mid(oMails.Body, 1, InStr(1, oMails.Body, "^") - 1)
                'sCommand = oMails.HTMLBody
                If InStr(1, oMails.Body, "~") <> 0 Then
                    ParseMail = Mid(oMails.Body, 1, InStr(1, oMails.Body, "~") - 1)
                    sParam = Mid(oMails.Body, InStr(1, oMails.Body, "~") + 1)
                Else
                    ParseMail = oMails.Body
                End If
                oMails.UnRead = False
            End If
        End If
    Next oMails
    Exit Function
ErrTrap:
    Select Case Err.Number
        Case 5
            Resume Next
        Case 424
            Resume Next
        Case Else
            txtLog = txtLog & "Error In ParseMail(): "
            txtLog = txtLog & Err.Description & vbCrLf
    End Select
End Function

Private Function SendMail(sTo As String, sSubject As String, Optional sAttachment As String = "", Optional sBody As String)
On Error GoTo ErrTrap
'
'Parameters : sTo - Send mail to; sSubject - Subject of the mail; sAttachment - If any attachment path going to send with mail
'
'Purpose    : This procedure will send mail to specific address
'
'To send mail to specified address
Dim objApp As Outlook.Application
Dim objMailMessage As Outlook.MailItem
'
    Set objApp = New Outlook.Application
    Set objMailMessage = objApp.CreateItem(olMailItem)
    '
    objMailMessage.To = sTo
    objMailMessage.Subject = sSubject
    
    If Trim(sBody) <> "" Then
        objMailMessage.Body = sBody
    End If
    If Trim(sAttachment) <> "" Then
        objMailMessage.Attachments.Add sAttachment
        
    End If
    objMailMessage.Send
    '
    Exit Function
ErrTrap:
    txtLog = txtLog & "Error In ParseMail(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Function

Private Function ReadMail(sMsgId As String)
On Error GoTo ErrTrap
'
'Purpose      This Procedure will read message and send it as SMS
'
'Set up our variables
Dim oApp        As Outlook.Application  'Create an object for outlook application
Dim oNpc        As NameSpace            'Name space to drildown message folder
Dim oMails      As MailItem             'To find our mail
Dim sMsg        As String
Dim lPos        As Long
'
    'Lets apply values to our variables
    Set oApp = CreateObject("Outlook.Application")
    Set oNpc = oApp.GetNamespace("MAPI")
    '
    'Lets iterate through an easy For Each loop
    For Each oMails In oNpc.GetDefaultFolder(olFolderInbox).Items
        If oMails.UnRead Then
            If Trim(UCase(Trim(oMails.EntryID))) = Trim(UCase(Trim(sMsgId))) Then
                'Read mail and send it as SMS
                sMsg = "Sub: " & oMails.Subject & vbCrLf
                If oMails.Attachments.Count <> 0 Then
                    sMsg = sMsg & "Att: Haves Attach"
                End If
                sMsg = sMsg & "Body: " & oMails.Body & "  "
                lPos = InStr(1, sMsg, "Original Message")
                If lPos > 0 Then
                    sMsg = Mid(sMsg, 1, lPos)
                End If
                SendMail "919768292949.ameet_kumar2003@smscountry.net", "readmail", "", "The mail is : " & sMsg
            End If
        End If
    Next oMails
    Exit Function
ErrTrap:
    txtLog = txtLog & Err.Description & vbCrLf
End Function
Private Function CheckMail() As Integer
On Error GoTo ErrTrap
'
'Purpose      This Procedure Count number of unread mails in inbox
'
'Set up our variables
Dim oApp        As Outlook.Application  'Create an object for outlook application
Dim oNpc        As NameSpace            'Name space to drildown message folder
Dim oMails      As MailItem             'To find our mail
Dim iMsgCount   As Integer
'
    'Lets apply values to our variables
    Set oApp = CreateObject("Outlook.Application")
    Set oNpc = oApp.GetNamespace("MAPI")
    '
    'Lets iterate through an easy For Each loop
    iMsgCount = 0
    For Each oMails In oNpc.GetDefaultFolder(olFolderInbox).Items
        If oMails.UnRead Then
            iMsgCount = iMsgCount + 1
        End If
    Next
    CheckMail = iMsgCount
    Exit Function
ErrTrap:
    txtLog = txtLog & Err.Description & vbCrLf

End Function

Private Function ReadMailHeader()
On Error GoTo ErrTrap
'
'Purpose      This Procedure Reads Mail Header info and send it as SMS
'
'Set up our variables
Dim oApp        As Outlook.Application  'Create an object for outlook application
Dim oNpc        As NameSpace            'Name space to drildown message folder
Dim oMails      As MailItem             'To find our mail
Dim sMsgHead    As String
'
    'Lets apply values to our variables
    Set oApp = CreateObject("Outlook.Application")
    Set oNpc = oApp.GetNamespace("MAPI")
    '
    'Lets iterate through an easy For Each loop
    sMsgHead = ""
    For Each oMails In oNpc.GetDefaultFolder(olFolderInbox).Items
        If oMails.UnRead Then
            sMsgHead = "From: " & oMails.SenderName & vbCrLf
            sMsgHead = sMsgHead & "Sub: " & oMails.Subject & vbCrLf
            sMsgHead = sMsgHead & "DT: " & oMails.SentOn & vbCrLf
            sMsgHead = sMsgHead & "MSGID: " & oMails.EntryID & "  "
            SendMail "919768292949.ameet_kumar2003@smscountry.net", "readmailheader", "", "The header is : " & sMsgHead
            
        End If
    Next
    Exit Function
ErrTrap:
    txtLog = txtLog & Err.Description & vbCrLf
End Function

Private Sub txtInterval_LostFocus()
    If Val(Trim(txtInterval)) > 60 Or Val(Trim(txtInterval)) < 0 Then
        MsgBox "Check Interval should be between 1 to 60 seconds", vbInformation
        txtInterval.SetFocus
    End If
End Sub

Private Function CreateFolder(folderPath As String)
On Error GoTo ErrTrap
    Dim oFso As New FileSystemObject
    
    If oFso.FolderExists(folderPath) Then
        'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of CREATE_FOLDER command", "", "Folder already Exists!!!"
    Else
        MkDir folderPath
        'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of CREATE_FOLDER command", "", "Folder created successfully!!!"
    End If
    Exit Function
ErrTrap:
    txtLog = txtLog & "Error In CreateFolder(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Function

Private Function DeleteFolderWithSubFolders(folderPath As String)
On Error GoTo ErrTrap
    Dim oFso As New FileSystemObject
    
    If oFso.FolderExists(folderPath) Then
        oFso.DeleteFolder folderPath, True
        'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of DELETE_FOLDER command", "", "Folder deleted successfully!!!"
    Else
        'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of DELETE_FOLDER command", "", "Folder not found!!!"
    End If
    Exit Function
ErrTrap:
    txtLog = txtLog & "Error In DeleteFolderWithSubFolders(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Function
    
Private Function DeleteFiles(filePath As String)
On Error GoTo ErrTrap
    Kill folderPath & "\*.*"
    Exit Function
ErrTrap:
    txtLog = txtLog & "Error In DeleteFiles(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Function

Private Function MoveFolder(oldPath As String, newPath As String)
On Error GoTo ErrTrap
    Dim oFso As New FileSystemObject
    
    If oFso.FolderExists(oldPath) Then
        oFso.MoveFolder oldPath, newPath
        'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of RENAME_FOLDER command", "", "Folder has been renamed!!!"
    Else
        'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of RENAME_FOLDER command", "", "Folder not found!!!"
    End If
    Exit Function
ErrTrap:
    txtLog = txtLog & "Error In RenameFolder(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Function

Private Function CopyFolder(source As String, destination As String)
On Error GoTo ErrTrap
    Dim oFso As New FileSystemObject
    
    If oFso.FolderExists(source) Then
        oFso.CopyFolder source, destination, True
        'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of COPY_FOLDER command", "", "Folder has been copied!!!"
    Else
        'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of COPY_FOLDER command", "", "Folder not found!!!"
    End If
    Exit Function
ErrTrap:
    txtLog = txtLog & "Error In RenameFolder(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Function

Private Function CreateTextFile(filePath As String, content As String)
On Error GoTo ErrTrap
    Dim oFso As New FileSystemObject
    Dim newFile
    Const ForReading = 1, ForWriting = 2
    
    If oFso.FileExists(filePath) Then
        'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of CREATE_FILE command", "", "File already Exists!!!"
    Else
        Set newFile = oFso.CreateTextFile(filePath, True)
        newFile.WriteLine (content)
        newFile.Close
        'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of CREATE_FILE command", "", "File created!!!"
    End If
    Exit Function
ErrTrap:
    txtLog = txtLog & "Error In CreateTextFile(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Function

Private Function GetSystemConfig()
On Error GoTo ErrTrap
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_ComputerSystem")
    For Each Object In List
        Msg = Msg & "Physical memomy: " & Object.TotalPhysicalMemory & vbCrLf
    Next
        
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_BaseBoard")
    For Each Object In List
        Msg = Msg & "Motherboard Serial Number: " & Object.SerialNumber & vbCrLf
    Next
        
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_LogicalDisk")
    For Each Object In List
        Msg = Msg & "Disk Serial Number: " & Object.VolumeSerialNumber & vbCrLf
    Next
    
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_Processor")
    For Each Object In List
        Msg = Msg & "CPU Caption: " & Object.Caption & vbCrLf
    Next
    
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_Processor")
    For Each Object In List
        Msg = Msg & "CPU Clock speed (in mega-hertz): " & Object.MaxClockSpeed & vbCrLf
    Next
    
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_Processor")
    For Each Object In List
        Msg = Msg & "CPU data width (32 or 64 bit): " & Object.DataWidth & vbCrLf
    Next
    
    MsgBox Msg
    
    'SendMail "919768292949.ameet_kumar2003@smscountry.net", "Result of PC_CONFIG command", "", "System information : " & Msg
            
ErrTrap:
    txtLog = txtLog & "Error In GetSystemConfig(): "
    txtLog = txtLog & Err.Description & vbCrLf
End Function

