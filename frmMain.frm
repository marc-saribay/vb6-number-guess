VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Number Guess"
   ClientHeight    =   5070
   ClientLeft      =   2790
   ClientTop       =   2130
   ClientWidth     =   5655
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "Click here if the number is NOT present in the card"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Yes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Click here if the number is present in the card"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame fraCard 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   5175
      Begin VB.Label lblNumbers 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number Guess™"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "©2001 by Marc Christian Saribay"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3075
      TabIndex        =   6
      ToolTipText     =   "NumGuess™ (August 18, 2001)"
      Top             =   120
      Width           =   2385
   End
   Begin VB.Label Label3 
      Caption         =   "Is the number in this card?"
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
      Left            =   360
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMain.frx":0442
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   400
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare varibles globally
Public intRange, intGuess, intCard, intInitial, intNumber, intCounter1, intCounter2 As Integer

Public Sub InitializeValues()
  'Reset variables to initial values
  intRange = 100
  intGuess = 0
  intCard = 1
  intInitial = 1
End Sub

Public Sub InitializeCounters()
  'Reset counters to initial values
  intCounter1 = 1
  intCounter2 = 1
End Sub

Public Sub GenerateNumbers()
  'Determine card number
  If intInitial <= intRange Then
    fraCard.Caption = "Card " + Str(intCard)
    intNumber = intInitial
    lblNumbers.Caption = ""
    InitializeCounters
    'Generate the numbers of current card
    Do While intRange >= intNumber
      If intCounter1 <= intInitial Then
        lblNumbers.Caption = lblNumbers.Caption + Str(intNumber)
        intNumber = intNumber + 1
        intCounter1 = intCounter1 + 1
      Else
        'Generate skipped numbers
        If intCounter2 <= intInitial Then
          intNumber = intNumber + 1
          intCounter2 = intCounter2 + 1
        Else
          InitializeCounters
        End If
      End If
    Loop
  Else
    'Message Box contents
    strPrompt1 = "The number is " & Str(intGuess) & ". Try Again?"
    strPrompt2 = "Please look carefully and answer honestly! Try Again?"
    strTitle = "Number Guess"
    'Display the guess number
    If intGuess > intRange Then
      Msg = MsgBox(strPrompt2, vbYesNo + vbExclamation, strTitle)
    Else
      Msg = MsgBox(strPrompt1, vbYesNo + vbQuestion, strTitle)
    End If
    If Msg = 6 Then
      'Retry program
      InitializeValues
      GenerateNumbers
    Else
      'Quit program
      Unload Me
    End If
  End If
End Sub

Private Sub cmdNo_Click()
  'Skip addition of the integer initial
  intCard = intCard + 1
  intInitial = intInitial * 2
  GenerateNumbers
End Sub

Private Sub cmdYes_Click()
  'Add the value of integer initial to integer guess
  intGuess = intGuess + intInitial
  intCard = intCard + 1
  intInitial = intInitial * 2
  GenerateNumbers
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    'Escape key function to quit program
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  'Initialize program
  InitializeValues
  GenerateNumbers
End Sub

