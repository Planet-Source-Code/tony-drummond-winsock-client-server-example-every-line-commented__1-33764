VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock Server Example"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   0
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "User Files (*.usr) | *.usr"
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   0
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
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
      Left            =   3120
      TabIndex        =   13
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "&Restart"
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
      Left            =   2160
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdDisable 
      Caption         =   "D&isable"
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
      Left            =   1200
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "&Enable"
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
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3975
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
      Left            =   3360
      TabIndex        =   0
      Text            =   "2323"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "[Connection Options]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   3735
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Idle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   14
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label lblUsers 
      Caption         =   "Users:  0"
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
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblIP 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "127.0.0.1"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer:  Tony Drummond/Bone
'Visit http://www.e-programmer.net
Public intCount As Integer 'Stores the number of winsock controls loaded
Public IsEnabled As Boolean 'True if the server is enabled
Private Sub LoadList(strDirectory As String, lst As ListBox)
    Dim strText As String 'Text buffer
    On Error Resume Next 'If an error occurs, skip to next line
    lst.Clear 'Clear the list box before loading
    Open strDirectory$ For Input As #1 'Open the file for reading
    While Not EOF(1) 'Loop while the program isn't at the End Of File marker
        Input #1, strText 'Insert the current line into the text buffer
        If strText <> "" Then 'Don't load blank list items
            DoEvents 'Reduces system hang caused by loops
            lst.AddItem strText 'Add the item to the list box
        End If
    Wend 'Ends the while loop
    Close #1 'Always close your file
End Sub
Private Sub SaveList(strDirectory As String, lst As ListBox)
    Dim i As Long 'Used in the for...next loop
    On Error Resume Next 'If an error occurs, skip to the next line
    Open strDirectory$ For Output As #1 'Open the file for writing
    For i& = 0 To lst.ListCount - 1 'Loop through all list items
        Print #1, lst.List(i&) 'Write text to file
    Next i&
    Close #1 'Always close your file
End Sub
Public Sub UpdateList()
    'This sub updates the label that is used to
    'display the amount of users
    lblUsers.Caption = "Users:  " & lstUsers.ListCount
End Sub
Private Sub cmdAdd_Click()
    frmMain.Enabled = False 'Disable the main form
    frmAddUser.Show 'Show the dialog window
End Sub
Private Sub cmdDelete_Click()
    'If list index is less than 0, user didn't select an item
    If lstUsers.ListIndex < 0 Then
        'Notify the user
        MsgBox "You did not select any item to delete.", vbOKOnly, "Error"
        Exit Sub 'Exit the sub
    End If
    'If all is good, remove the list item
    lstUsers.RemoveItem (lstUsers.ListIndex)
    UpdateList 'Update the user count
End Sub
Private Sub cmdDisable_Click()
    'If an error occurs, we don't want to hear about it because
    'it's more than likely from trying to send data to a socket
    'that isn't connected due to someone disconnecting on the other side.
    On Error Resume Next
    Dim i As Long 'Used in the for...next loop
    For i& = 0 To wskServer.Count - 1
        'Notify user that they are going to be disconnected.
        wskServer(i).SendData ("Server was disabled by the administrator." & vbCrLf)
        wskServer(i).Close 'Close the connection
        Unload wskServer(i) 'Unload the socket
    Next i&
    'Update status
    lblStatus.Caption = "Server Successfully Disabled."
    'Tell the user that the server was disabled
    MsgBox "Server successfully disabled.", vbOKOnly, "Winsock Server"
    cmdDisable.Enabled = False 'Disable the "Disable server" button
    cmdEnable.Enabled = True 'Enable the "Enable server" button
    IsEnabled = False 'Make sure you set the enabled flag
    'Update status
    lblStatus.Caption = "Idle"
End Sub
Private Sub cmdEnable_Click()
    If IsNumeric(txtPort.Text) Then
        wskServer(0).LocalPort = Val(txtPort.Text) 'Set the listening port
        wskServer(0).Listen 'Listen for connections
        'Change the status label's caption
        lblStatus.Caption = "Listening for connections [Server Enabled]"
        'Disable this button and enable the button "Disable"
        cmdDisable.Enabled = True
        cmdEnable.Enabled = False
        IsEnabled = True 'Make sure you set the enabled flag
    Else
        'Let the user know they need to fix the port number
        MsgBox "The port number you have chosen is invalid.  This usually occurs when you use letters in your port number.", vbOKOnly, "Error"
    End If
End Sub
Private Sub cmdExit_Click()
    Dim intResp As Integer 'Stores the response from the message box function
    intResp = MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Exit Program?")
    Select Case intResp 'Case structure
        Case vbYes 'User pressed yes?
            End 'End the program
        Case vbNo 'User pressed no?
            Exit Sub 'Exits the subroutine without executing anymore code
    End Select
End Sub
Private Sub cmdLoad_Click()
    Dim strFile As String 'Stores the file name
    comDialog.ShowOpen 'Show the load dialog
    strFile = comDialog.FileName 'Retrieve the file name
    'If the user presses cancel, strFile will be empty
    If strFile = vbNullString Then Exit Sub
    'All else is good, let's run the load list sub
    Call LoadList(strFile, lstUsers)
    'Update the list count
    UpdateList
End Sub

Private Sub cmdRestart_Click()
    MsgBox "This feature has not been written.  Want to know why?  It's your job to figure out how to write it.  I feel that I have provided you with more than enough information to accomplish this task.", vbOKOnly, "Error"
End Sub

Private Sub cmdSave_Click()
    Dim strFile As String 'Stores the file name
    comDialog.ShowSave 'Show the save dialog
    strFile = comDialog.FileName 'Retrieve the file name
    'If the user presses cancel, strFile will be empty
    If strFile = vbNullString Then Exit Sub
    'Everything is good, time to save the list box
    Call SaveList(strFile, lstUsers)
End Sub
Private Sub Form_Load()
    'Display the users IP
    '--------------------
    'I'm using control arrays in this particular example to allow
    'multiple people to connect to this program
    lblIP.Caption = wskServer(0).LocalIP
    lstUsers.AddItem "default:default" 'Add the default username and password
    UpdateList 'Update the list count
End Sub
Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    intCount = wskServer.Count 'Get the total number of winsock controls
    Load wskServer(intCount) 'Load a new control
    wskServer(intCount).Close 'Make sure it's closed
    wskServer(intCount).Accept (requestID) 'Accept the connection request
    'Send output to the person who just connected
    'For your reference:  vbCrLf is a vb constant which Stores
    'ascii values for the carriage return(CR), and the line feed (LF)
    wskServer(intCount).SendData ("00_DISP Welcome to Bone's Winsock Server" & vbCrLf & "Type 'help' for a command list." & vbCrLf)
End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'If some idiot decides to send your server a packet like...
    '01_LOGIN and the server is expecting 01_LOGIN BONE:PASS
    'It will crash when your program tries to parse the packet
    'that's why I always put "On Error Resume Next" Also, it's best
    'to add this line when you're done programming, that way you can
    'still receive errors when you're still in the development stage.
    On Error Resume Next
    Dim strData As String 'Stores the incoming data
    Dim strUser As String 'Stores the user name
    Dim strPass As String 'Stores the password
    Dim strIPs As String 'Stores the IP addresses (used for IPs command)
    Dim strCmd As String 'Stores the command ie. 01_LOGIN
    Dim strNewData As String 'Stores new data (after parsing)
    Dim intSpace As Integer 'Stores the position of a space character
    Dim intColon As Integer 'Stores the position of the colon
    Dim i As Long 'Used in a for...next loop
    'Store the incoming data in our variable
    Call wskServer(Index).GetData(strData)
    'Now, I want to restrict the user from doing certain things if
    'they aren't properly logged in.  If I didn't do this, anyone
    'could use a winsock connection utility and connect to my server
    'and begin executing commands, even though they don't belong here.
    'To prevent this I'm simply going to check the value of the
    'socket's tag property.  Then, when the user logs in correctly I'm
    'going to set the tag to "LOGGED IN" allowing them to have full
    'access
    Select Case wskServer(Index).Tag
        'The tag property is blank if the user isn't logged in yet
        Case ""
            'Get the position of the first space in the packet
            intSpace = InStr(strData, " ")
            'Use the space's position to get the actual command
            'sent to the server.
            strCmd = Left(strData, intSpace - 1)
            'Is the command the login packet?
            If strCmd = "01_LOGIN" Then
                'Message Structure:
                '01_LOGIN USER:PASS
                'Trim off the 01_LOGIN
                strNewData = Mid(strData, intSpace + 1)
                'Find the position of the colon that seperates
                'the username and the password
                intColon = InStr(strNewData, ":")
                'Get the user name
                strUser = Left(strNewData, intColon - 1)
                'Get the password
                strPass = Mid(strNewData, intColon + 1)
                'Now, you might be wondering why I seperated the username
                'from the password, when they're stored in the list box
                'as USER:PASS.  Well, I did this incase I felt like storing
                'the user name somewhere for future use, such as displaying
                'all users currently logged in.
                'Next, we loop through the list box and check the login info
                For i& = 0 To lstUsers.ListCount - 1
                    'Check to see if it matches
                    If lstUsers.List(i) = strUser & ":" & strPass Then
                        'The user logged in successfully
                        wskServer(Index).SendData ("00_DISP " & strUser & " logged in successfully at " & Now)
                        'Don't forget to set your tag property to "LOGGED IN"
                        wskServer(Index).Tag = "LOGGED IN"
                        'Well, we're done now so we can exit the sub
                        Exit Sub
                    End If
                Next i&
                'This segment of code only executes if the username and password
                'were invalid.  I did this by putting "Exit Sub" in my if statement
                '------------------------------------------------------------
                'Well, they didn't have the right login info, let's tell them
                wskServer(Index).SendData ("02_INVALID The username/password you have entered is invalid.  Please try again.")
            End If
        'The tag property is "LOGGED IN" after the user has logged in
        Case "LOGGED IN"
            'The 03_DATA command will only be allowed if the person is logged
            'into the server.  See above for a better explination
            'Get the position of the space character
            intSpace = InStr(strData, " ")
            'Get the command sent to the server
            strCmd = Left(strData, intSpace - 1)
            'Check to see if it's the 03_DATA command.  This command is sent
            'whenever a user sends data through use of the text box or the
            'send button on the client.
            If strCmd = "03_DATA" Then
                'Now we parse the strData string and trim off the command
                'We do this because it's no longer needed.
                strNewData = Mid(strData, intSpace + 1)
                'Below is the code where it accepts all of the command
                'the users enters from the client side.
                'Did the user enter the help command?  Also when checking
                'your commands lowercase them so if they type hElP it still
                'works properly.
                If LCase(strNewData) = "help" Then
                    'Send the help information to the user
                    'Also, the text is formatted properly because I'm
                    'using Courier New text on my client.
                    wskServer(Index).SendData _
                    "00_DISP " & _
                    "Welcome To The Help Center" & vbCrLf & _
                    "--------------------------" & vbCrLf & _
                    "Commands:" & vbCrLf & _
                    "Close           - Closes your connection" & vbCrLf & _
                    "Hi              - YOU MUST WRITE THIS CODE -- PRACTICE" & vbCrLf & _
                    "IPs             - Displays all IP addresses" & vbCrLf & _
                    "Help            - Displays help information" & vbCrLf & _
                    "Say <text here> - Says something to everyone logged in" & vbCrLf
                    'Exit the sub routine
                    Exit Sub
                End If
                'Check to see if the user entered the "Close" command
                'If so, close their connection.
                If LCase(strNewData) = "close" Then
                    'Send the disconnect message
                    wskServer(Index).SendData ("00_DISP Disconnecting you from the server...")
                    'If you don't put this, the client does not display the data you just sent
                    'This allows the program to finish up before going to the next line of code
                    DoEvents
                    'Close the connection
                    wskServer(Index).Close
                End If
                'Check to see if the user entered the "IPs" command
                'This command sends the IP's of all logged in users
                'to the client
                If LCase(strNewData) = "ips" Then
                    'Give our IPs string a header to display
                    strIPs = "IP Addresses" & vbCrLf & "------------" & vbCrLf
                    'Loop through all winsock controls that are connected
                    'and logged in properly
                    For i& = 1 To wskServer.UBound
                        'Check to see if the socket is connected and logged in
                        If wskServer(i).State = sckConnected And wskServer(i).Tag = "LOGGED IN" Then
                            'Add the IP to buffer I have setup to hold them.  We will
                            'send these all at once when we're finished
                            strIPs = strIPs & i & ":  " & wskServer(i).RemoteHostIP & vbCrLf
                        End If
                    Next i&
                    'Now, we have the IP's of everyone connected, let's send it to the user
                    wskServer(Index).SendData "00_DISP " & strIPs
                    'Exit the sub routine
                    Exit Sub
                End If
                'The like operator allows us programmers to use if statements
                'that allow for wild cards.  Although you could do it other
                'ways as well.
                If LCase(strNewData) Like "say *" Then
                    'Trim "say " it's not needed anymore
                    strNewData = Mid(strNewData, 5)
                    'Next we loop through all winsock controls loaded
                    'and check to see if they're connected.  It's rather
                    'pointless to send data to a disconnected socket, it
                    'will also give you errors.
                    For i& = 1 To wskServer.UBound
                        'Check to see if the socket is connected.
                        'Also, check and make sure the user is LOGGED IN.
                        'If you don't check for this, you will send data to
                        'everyone connected, and they can snoop on conversations
                        If wskServer(i).State = sckConnected And wskServer(i).Tag = "LOGGED IN" Then
                            'The socket is connected, let's send the data
                            wskServer(i).SendData ("00_DISP " & Index & ":  " & strNewData)
                        End If
                        'If you do not include this, your program will not send data properly
                        'I know because I experienced this problem and I was like, "WTF"
                        'Then, I simply put this here to see if it would fix the problem
                        'and sure enough, BINGO!
                        DoEvents
                    Next i&
                    'Exit the sub routine
                    Exit Sub
                End If
            End If
    End Select
End Sub
