VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{220F55E8-7AAE-11D3-9D68-F74ED5721646}#18.0#0"; "TAni.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   1830
   ClientTop       =   3480
   ClientWidth     =   5910
   Icon            =   "Form1.frx":0000
   ScaleHeight     =   5055
   ScaleWidth      =   5910
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "7"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtIp 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtIp 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtIp 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   4440
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   5160
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin TAni.TMaxAni GifAnim 
      Height          =   1455
      Left            =   4560
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2566
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Scan"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "The Dancing Girl"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   2895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuBlank 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuBlank2 
      Caption         =   ""
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutHelp 
         Caption         =   "&Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LoadOnce As Boolean
Dim LPbaSe As Integer
Dim MyIpAdd As String
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Private Sub Command1_Click()
Dim x As Integer
If LoadOnce = False Then LoadWinsck
If LoadOnce = True Then CloseUDPWinsck
Label3 = 15: Label3.Visible = True: GifAnim.Visible = True: Label2.Visible = True

Timer1 = True
On Error GoTo TrapAnError
    List1.Clear
    MyIpAdd = txtIp(0) & "." & txtIp(1) & "." & txtIp(2) & "."
    
For x = 1 To 254
    Winsock1(x).LocalPort = LPbaSe
    Winsock1(x).Protocol = sckUDPProtocol
    Winsock1(x).RemoteHost = MyIpAdd & x
    Winsock1(x).RemotePort = txtPort
    Winsock1(x).SendData "1"
    LPbaSe = LPbaSe + 1
Next x
    
    Exit Sub

TrapAnError:
    If Err.Number = 10048 Then ' Port Address in use
    LPbaSe = LPbaSe + 1
        Winsock1(x).LocalPort = LPbaSe
        
        Resume
    ElseIf Err.Number = 10065 Then ' No route to host
        MsgBox "No route to Host..." & vbCrLf & "I can not reach that destination because it -" & _
        vbCrLf & "Does not exist or you are not connected to a LAN or Internet." & _
        vbCrLf & "Selected a Host IP range that you can Ping.", vbQuestion Or vbOKOnly, "Warning! Bill Robinson!"
        LPbaSe = LPbaSe + 5
        CloseUDPWinsck
    Else
    MsgBox "Error in the Scan Button..." & vbCrLf & Err.Number & vbCrLf & Err.Description, _
     vbCritical Or vbOKOnly, "Shane's Port Scan Error Control"
    'Text1 = testb & "fark" & vbCrLf
    End If
    Exit Sub
End Sub
Private Sub Form_Load()
Form1.Caption = "Shane's Worlds Fastest Host Scanner..."
Label1 = "IP Address                       Port Number "
    LPbaSe = 5009
    Call RemoveCancelMenuItem(Me)
GifAnim.FileName = App.Path & ("\dance.gif")
GifAnim.ShowGif
End Sub
Private Sub mnuAboutHelp_Click()
    frmHelp.Show
End Sub

Private Sub mnuFileExit_Click()
    Dim x As Integer
    If LoadOnce = False Then GoTo GetOut
    
    For x = 1 To 254
        Unload Winsock1(x)
    Next
GetOut:
    frmExit2.Show
    MsgBox "elvis007now@hotmail.com", vbInformation Or vbOKOnly, "Sianara..."
    Unload frmExit2
    Unload Me
End Sub
Private Sub Timer1_Timer()
Dim x As Integer

Label3 = Label3 - 1
    If Label3 = 0 Then
        Timer1 = False
        Label2.Visible = False: Label3.Visible = False: GifAnim.Visible = False
    End If
End Sub

Private Sub txtIp_GotFocus(Index As Integer)
    txtIp(0).SelStart = 0
    txtIp(0).SelLength = Len(txtIp(0))
    txtIp(1).SelStart = 0
    txtIp(1).SelLength = Len(txtIp(1))
    txtIp(2).SelStart = 0
    txtIp(2).SelLength = Len(txtIp(2))
End Sub

Private Sub Winsock1_DataArrival(x As Integer, ByVal bytesTotal As Long)
    Dim TempData As String, MyData As String
    On Error GoTo TrapAnError
    
    Winsock1(x).GetData TempData
    
    If TempData = 1 Then
    MyData = MyIpAdd & x & " Host is Alive & Port " & txtPort & " is responding"
    End If
7
    
    Exit Sub
TrapAnError:
    If Err.Number = 10054 Then ' Connection Reset by Remote Side
       
        MyData = MyIpAdd & x & " Host is Alive..."
        List1.AddItem MyData
    Else
    End If
    Resume 7
End Sub
Private Sub LoadWinsck()
Dim x As Integer
LoadOnce = True

On Error GoTo TrapAnError

    For x = 1 To 254
        Load Winsock1(x)
    Next x
    Exit Sub
TrapAnError:
    MsgBox "Could not load the WinSocks..."
End Sub
Private Sub CloseUDPWinsck()
Dim x As Integer

    For x = 1 To 254
        Winsock1(x).Close
    Next x
End Sub
Private Sub RemoveCancelMenuItem(frm As Form)
   Dim hSysMenu As Long
   'get the system menu for this form
   hSysMenu = GetSystemMenu(frm.hWnd, 0)
   'remove the close item
   Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
   'remove the separator that was over the close item
   Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub
