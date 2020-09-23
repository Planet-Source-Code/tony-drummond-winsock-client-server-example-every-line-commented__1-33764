VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add User Dialog"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   975
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
      Left            =   720
      TabIndex        =   3
      Text            =   "Password"
      Top             =   480
      Width           =   2535
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
      Left            =   720
      TabIndex        =   1
      Text            =   "UserName"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Pass:"
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
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "User:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    'Enable the main form
    frmMain.Enabled = True
    'Hide this form
    frmAddUser.Hide
End Sub

Private Sub cmdOK_Click()
    'Did the user fill out both text boxes?
    If txtUser.Text = vbNullString Or txtPass.Text = vbNullString Then
        'Tell the user to fill out both text boxes
        MsgBox "Sorry, but you must fill out both text boxes to continue.", vbOKOnly, "Error"
        'Exit the sub
        Exit Sub
    End If
    'Add the User and Pass to the list box
    frmMain.lstUsers.AddItem txtUser.Text & ":" & txtPass.Text
    'Update the list count
    frmMain.UpdateList
    'Enable the main form
    frmMain.Enabled = True
    'Hide this window
    frmAddUser.Hide
End Sub

Private Sub cmdReset_Click()
    'Reset both text boxes
    txtUser.Text = ""
    txtPass.Text = ""
End Sub
