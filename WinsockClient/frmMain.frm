VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock Client Example"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabControl 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Connection"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fmeConnectionSettings"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtDisplay"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtMessage"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSend"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Debug Window"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDebug"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtDebug 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "frmMain.frx":0038
         Top             =   600
         Width           =   6735
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtMessage 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "<Data to send to server here> (Pressing enter sends data)"
         Top             =   4320
         Width           =   5535
      End
      Begin VB.TextBox txtDisplay 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "frmMain.frx":0109
         Top             =   1800
         Width           =   6735
      End
      Begin VB.Frame fmeConnectionSettings 
         Caption         =   "Connection Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6735
         Begin MSWinsockLib.Winsock wskClient 
            Left            =   6240
            Top             =   0
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "&Connect"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   11
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtPass 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3840
            PasswordChar    =   "*"
            TabIndex        =   10
            Text            =   "default"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtUser 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   8
            Text            =   "default"
            Top             =   720
            Width           =   2175
         End
         Begin VB.CommandButton cmdClearFields 
            Caption         =   "C&lear All Fields"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtPort 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3840
            TabIndex        =   5
            Text            =   "2323"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtIP 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Text            =   "127.0.0.1"
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "P&ass:"
            Height          =   255
            Left            =   3360
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "&User:"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "&Port:"
            Height          =   255
            Left            =   3360
            TabIndex        =   4
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "&IP Address:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer:  Tony Drummond/Bone
'Comments:
'In this particular example you will find that it is probably
'one of the most informative you will ever download.  Everything should
'seem VERY clear to even a newbie.  I hope you learn from this and
'don't forget to visit www.e-programmer.net for more of my work.
'-------------------------------------------------------------------
'This API call get's the number of ticks since your computer has been started
'This allows us programmers to easily determine how much time has gone by.
'Also, the value of this API call is returned in milliseconds.
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Sub wskSendData()
    'This sends data to the server
    'Use the winsock control to send data to the server you're
    'connected to.
    wskClient.SendData ("03_DATA " & txtMessage.Text)
    'Clear the text box that holds the message to send
    txtMessage.Text = ""
End Sub
Private Sub UpdateText(strText As String)
    'This just updates the text box with the new data
    'I used a sub here to make things easier and more
    'organized.
    txtDisplay.Text = txtDisplay.Text & vbCrLf & strText
End Sub
Private Sub cmdClearFields_Click()
    'Clear all text boxes on the form, except for txtDisplay and txtMessage
    txtIP.Text = ""
    txtPort.Text = ""
    txtUser.Text = ""
    txtPass.Text = ""
End Sub

Private Sub cmdConnect_Click()
    Dim lngTickCount As Long 'Stores the tick count (used for timeout)
    Dim lngSeconds As Long 'Stores the seconds after tick count is converted
    'Check to make sure there are no empty text boxes
    If txtIP.Text = "" Or txtPort.Text = "" Or txtUser.Text = "" Or txtPass.Text = "" Then
        'Notify the user of the problem
        MsgBox "You have not properly filled out all of the required text boxes needed to connect.  Please check to make sure you have filled in the following:" & vbCrLf & vbCrLf & "Remote IP Address" & vbCrLf & "Port Number" & vbCrLf & "User Name" & vbCrLf & "Password", vbOKOnly, "Error"
        'Exit the sub, we have nothing else to do here
        Exit Sub
    End If
    'Allow different code segments to execute depending on
    'the button's caption
    Select Case cmdConnect.Caption
        'Now, if you're ever writing this on your own and you can't get
        'your code to execute check and make sure you include the
        'ampersand (&) in the proper place.  This is used for short cuts
        'In your program, so the user can press alt and trigger a button
        Case "&Connect"
            'Well, it's good practice to close a socket before you connect
            'That way if some idiot does decide to play around with your
            'program by using API to set the button's caption to "Connect"
            'you won't get any errors.  Also, if the server shuts down,
            'the winsock control likes to think it's still connected, which
            'will give you some errors as well.
            wskClient.Close
            'Check to see if the port number is numerical
            If IsNumeric(txtPort.Text) = False Then
                'Notify the user so they can fix the problem
                MsgBox "You are using a non-numerical port.  Please check your settings and try again.", vbOKOnly, "Error"
                'Exit the sub
                Exit Sub
            End If
            'Connect to the server
            Call wskClient.Connect(txtIP.Text, Val(txtPort.Text))
            'Get the current tick count (used for timeout)
            lngTickCount = GetTickCount
            'Loop until the client is connected or times out
            Do
                'Prevents system hang caused by loops, also continues
                'processing messages properly
                DoEvents
                'Get the number of seconds gone by.
                'This API call returns milliseconds so divide by 1000
                'and round it to simplify things a bit.
                lngSeconds = Round((GetTickCount - lngTickCount) / 1000)
                'Check to see if the connection attempt times out
                'In this particular example I allowed for 8 seconds
                'This should be more than enough time
                If lngSeconds > 8 Then
                    'Notify the user so they can check the server
                    MsgBox "The connection attempt has timed out.  Check to see if the server is currently enabled and try again later.", vbOKOnly, "Error"
                    'Exit the sub -- We don't need to do anything else
                    Exit Sub
                End If
            'Keep looping until the client is connected to the server
            Loop Until wskClient.State = sckConnected
            'Don't forget to change the caption of the button
            cmdConnect.Caption = "&Disconnect"
            'Enable the text box so the user can send data
            txtMessage.Enabled = True
            'Enable the "Send" button
            cmdSend.Enabled = True
            'Send the logon packet to the server
            '01_LOGIN USER:PASS
            wskClient.SendData ("01_LOGIN " & txtUser.Text & ":" & txtPass.Text)
        Case "&Disconnect"
            'Close the connection to the server
            wskClient.Close
            'Change the caption back to "Connect"
            cmdConnect.Caption = "&Connect"
            'Disable the text box, sending data while not connected causes
            'errors
            txtMessage.Enabled = False
            'Disable the "Send" button
            cmdSend.Enabled = False
    End Select
End Sub

Private Sub cmdSend_Click()
    'Send the specified data to the server
    wskSendData
End Sub

Private Sub txtDisplay_Change()
    'If you are using a text box that might possibly exceed 32,768 characters
    'then it's best to include a code to truncate the text, doing so stops
    'buffer overflow errors caused by the text box control.
    If Len(txtDisplay.Text) > 15000 Then
        'Truncate the text, we're going to keep the last 5000 characters
        txtDisplay.Text = Right(txtDisplay, 5000)
    End If
    'Now, to keep the cursor at the bottom (auto scroll text box) you
    'simply do the following
    txtDisplay.SelStart = Len(txtDisplay.Text)
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    'Check to see if the enter key was pressed
    If KeyAscii = 13 Then
        'Send the data to the server
        wskSendData
        'This makes that annoying *beep* noise go away
        KeyAscii = 0
    End If
End Sub

Private Sub wskClient_Close()
    'Well, it is possible for the server to disconnect you, or
    'maybe you get kicked offline by your ISP (AOL enjoys this)
    'Anyway, it's proper to disable the "Send" button and the text
    'box that holds the data.  Also, don't forget to change your
    'connection button back to "Connect"
    cmdConnect.Caption = "&Connect"
    txtMessage.Enabled = False
    cmdSend.Enabled = False
    'I like to make sure the winsock control is closed all the way
    'It doesn't take much effort to type this anyway, and it may
    'prevent some errors.
    wskClient.Close
End Sub

Private Sub wskClient_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String 'Stores the data sent by the server
    Dim strCmd As String 'Stores the command ie. 00_DISP
    Dim intSpace As Integer 'Stores the space character's position
    Dim strNewData As String 'Stores the new data (after parsing)
    'Get the data sent by the server
    Call wskClient.GetData(strData)
    'Get the position of the space character
    intSpace = InStr(strData, " ")
    'Parse the command sent by the server
    strCmd = Left(strData, intSpace - 1)
    'Truncate the text and trim off the command
    strNewData = Mid(strData, intSpace + 1)
    'This is sent when the server wants to display data to the user
    'using the txtDisplay text box
    If strCmd = "00_DISP" Then
        'Update the text box with the new data
        UpdateText (strNewData)
        'Exit the sub, nothing else to do
        Exit Sub
    End If
    'This is sent when you enter an invalid password
    If strCmd = "02_INVALID" Then
        'Update the text box
        UpdateText (strNewData)
        'Since it's an invalid password we have to disconnect from the server
        'since I already had my wskClient_Close sub programmed, I used it
        'to save myself a little time.
        Call wskClient_Close
        'Exit the sub, nothing else to do here
        Exit Sub
    End If
End Sub
