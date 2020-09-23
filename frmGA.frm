VERSION 5.00
Begin VB.Form frmGA 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A new Genetic Algorithm"
   ClientHeight    =   5715
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   693
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Worst Chromosome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Width           =   4575
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   11
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "Fitness"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Genome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Made By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Best Chromosome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   4575
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Made By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Genome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Fitness"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmGA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim TimeElapsed As Long
Private Sub About_Click()
Yesorno = MsgBox("Made by Paras Chopra, paraschopra@lycos.com, Do you want to se the readme?", vbYesNo, "About")
If Yesorno = vbYes Then
Shell "start " & App.Path & "\readme.txt"
End If
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Y_OFF_E = Me.ScaleHeight '- (Command3.Top + Command3.Height)
X_OFF = 0
Y_OFF_old = Me.ScaleHeight
X_OFF_old = 0
Me.Cls
Main
End Sub

Sub Main()
Dim intnum As Long
'Timer1.Enabled = False
Allow = 100
inputnum = Int(InputBox("Enter the number which you want to resolve in form: x^2 + y^2 + z^2"))
If inputnum <> vbCancel Or IsNumeric(inputnum) = True Then
NumToFind = inputnum
'Timer1.Enabled = True
Call BuildPopu(100, 5, 100, 97, 5, 75)
Evolve
'MsgBox "Solution in : " & CLng(TimeElapsed / 10) & " seconds"
'Timer1.Enabled = False
'TimeElapsed = 0
End If
End Sub

Private Sub Command3_Click()
PopuMain.StopEvolution = True
End Sub

Private Sub Form_Load()
Y_OFF_E = Me.ScaleHeight '- (Command3.Top + Command3.Height)
X_OFF = 0
Y_OFF_old = Me.ScaleHeight
X_OFF_old = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
PopuMain.StopEvolution = True
End Sub

'Private Sub Timer1_Timer()
'TimeElapsed = TimeElapsed + 1
'End Sub


Public Sub PrntTextBest(Fitness As String, Genome As String, MadeBy As String)
Text1.Text = Fitness
Text2.Text = Genome
Text3.Text = MadeBy
End Sub

Public Sub PrntTextWorst(Fitness As String, Genome As String, MadeBy As String)
Text6.Text = Fitness
Text5.Text = Genome
Text4.Text = MadeBy
End Sub

