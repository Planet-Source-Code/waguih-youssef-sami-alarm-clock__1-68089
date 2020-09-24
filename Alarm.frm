VERSION 5.00
Begin VB.Form Alarm 
   Caption         =   "Waguih Alarm"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   Icon            =   "Alarm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   1560
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdMinimise 
      Caption         =   "OK"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtAlarm2 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtAlarm1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Minutes 0-60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hours 1-24"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Alarm Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Alarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyFile As String
Private Type NOTIFYICONDATA
    cbSize As Long              'Size of the structure
    hwnd As Long                'Window handle of the icon's owner
    uId As Long                 'Unique identifier, use for multiple icons
    uFlags As Long              'Flags
    uCallBackMessage As Long    'Window message (WM) sent to the icon's owner
    hIcon As Long               'Handle of the icon to use (use VB's Form.Icon property)
    szTip As String * 64        'ToolTip textType
End Type

'Shell_NotifyIcon messages
Private Const NIM_ADD = &H0         'Add to tray
Private Const NIM_MODIFY = &H1      'Change Icon
Private Const NIM_DELETE = &H2      'Delete Icon
 
'NotifyIconData uFlags parameters, specify which members of
'NOTIFYICONDATA are valid and should be used by Shell_NotifyIcon
'AND these together in uFlags member
Private Const NIF_MESSAGE = &H1     'Honor uCallbackMessage member
Private Const NIF_ICON = &H2        'Honor hIcon member
Private Const NIF_TIP = &H4         'Honor szInfo member
 

'Window messages Shell_NotifyIcon sends to your app
Private Const WM_MOUSEMOVE = &H200       'MouseMove message
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
 
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
 
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
 
Private m_IconData As NOTIFYICONDATA
Private m_lngLastMessage As Long



Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdMinimise_Click()
Alarm.WindowState = 1
End Sub

Private Sub Form_DblClick()
Me.Show

End Sub

Private Sub Form_Load()
Text1.Text = Time
With m_IconData
        .cbSize = Len(m_IconData)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE    'Use MouseMove message
        .hIcon = Me.Icon
        .szTip = "Waguih Alarm" & vbNullChar
              
    End With

txtAlarm1.Enabled = True
txtAlarm2.Enabled = True

Dim X
X = InputBox("Put in the Period in minutes!")
Dim Mytime
If X <> "" Then
Mytime = DateAdd("n", X, Time)
txtAlarm1.Text = Hour(Mytime)
txtAlarm2.Text = Minute(Mytime)
End If

Dim MyPath
MyPath = App.Path
'MMControl1.filename = MyPath & "\" & "cook1.WAV"
''MMControl1.filename = "E:\MUSIC\Wav-files\cook1.WAV"
'MMControl1.Command = "Open"
MyFile = MyPath & "\" & "cook1.WAV"

Shell_NotifyIcon NIM_ADD, m_IconData
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
   
    Dim lRet As Long
    Dim lMessage As Long
   
    If Me.ScaleMode = vbPixels Then
        'No conversion needed
        lMessage = X
    Else
        'VB mangled X to convert it from Pixels to Twips
        lMessage = X / Screen.TwipsPerPixelX
    End If
   
    'We assume a click on BUTTONUP messages
    Select Case lMessage
          
    Case WM_LBUTTONDBLCLK
        'Double-click, restore the form
        Result = SetForegroundWindow(Me.hwnd)
        'MsgBox "Double-Click"
        Me.WindowState = vbNormal
        Me.Show
        txtAlarm1.Enabled = True
        txtAlarm2.Enabled = True
                  
    End Select
   
    m_lngLastMessage = lMessage
End Sub
Private Sub Form_Resize()
    'Hide the window if it's been minimized
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    'Remove the icon from the tray
    Shell_NotifyIcon NIM_DELETE, m_IconData
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub Timer1_Timer()
Text1.Text = Time
Alarm.Refresh
If txtAlarm1.Text & ":" & txtAlarm2.Text = CStr(Hour(Time) & ":" & CSng(Minute(Time))) Then
'  MMControl1.Command = "Play"
'    If MMControl1.Position = MMControl1.Length Then
'    MMControl1.Command = "Prev"
'    End If
PlaySoundX MyFile

txtAlarm2.Text = CSng(txtAlarm2.Text) + 1
End If
If CSng(txtAlarm2.Text) = 60 Then
txtAlarm2.Text = 0
txtAlarm1.Text = txtAlarm1.Text + 1
End If
End Sub
Private Sub PlaySoundX(filename As String)


    PlaySound filename, CLng(0), SND_FILENAME


End Sub
