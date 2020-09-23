VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Skinnable OptionButton"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin Project1.OptionEx OptionEx12 
      Height          =   195
      Left            =   720
      TabIndex        =   16
      Top             =   5040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicChecked      =   "Form1.frx":0000
      PicDisabled     =   "Form1.frx":017A
      PicUnchecked    =   "Form1.frx":0224
   End
   Begin Project1.OptionEx OptionEx11 
      Height          =   195
      Left            =   720
      TabIndex        =   15
      Top             =   4680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   344
      PicChecked      =   "Form1.frx":038D
      PicDisabled     =   "Form1.frx":0507
      PicUnchecked    =   "Form1.frx":05B1
   End
   Begin Project1.OptionEx OptionEx10 
      Height          =   195
      Left            =   720
      TabIndex        =   14
      ToolTipText     =   "test"
      Top             =   4320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicChecked      =   "Form1.frx":071A
      PicDisabled     =   "Form1.frx":0894
      PicUnchecked    =   "Form1.frx":093E
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3375
      Left            =   2880
      TabIndex        =   4
      Top             =   360
      Width           =   5295
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   2415
         Left            =   2040
         TabIndex        =   10
         Top             =   840
         Width           =   3135
         Begin Project1.OptionEx OptionEx9 
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1200
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   344
            PicChecked      =   "Form1.frx":0AA7
            PicDisabled     =   "Form1.frx":0C21
            PicUnchecked    =   "Form1.frx":0CCB
         End
         Begin Project1.OptionEx OptionEx8 
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   344
            PicChecked      =   "Form1.frx":0E34
            PicDisabled     =   "Form1.frx":0FAE
            PicUnchecked    =   "Form1.frx":1058
         End
         Begin Project1.OptionEx OptionEx7 
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   344
            PicChecked      =   "Form1.frx":11C1
            PicDisabled     =   "Form1.frx":133B
            PicUnchecked    =   "Form1.frx":13E5
         End
      End
      Begin Project1.OptionEx OptionEx6 
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   423
         PicChecked      =   "Form1.frx":154E
         PicDisabled     =   "Form1.frx":17AF
         PicUnchecked    =   "Form1.frx":1942
      End
      Begin Project1.OptionEx OptionEx5 
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   423
         PicChecked      =   "Form1.frx":1B97
         PicDisabled     =   "Form1.frx":1DF8
         PicUnchecked    =   "Form1.frx":1F8B
      End
      Begin Project1.OptionEx OptionEx4 
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   423
         PicChecked      =   "Form1.frx":21E0
         PicDisabled     =   "Form1.frx":2441
         PicUnchecked    =   "Form1.frx":25D4
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2175
      Begin VB.CommandButton Command2 
         Caption         =   "Enable"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Disable"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1455
      End
      Begin Project1.OptionEx OptionEx3 
         Height          =   210
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         PicChecked      =   "Form1.frx":2829
         PicDisabled     =   "Form1.frx":2A77
         PicUnchecked    =   "Form1.frx":2BD7
         ForeColor       =   0
      End
      Begin Project1.OptionEx OptionEx2 
         Height          =   210
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         PicChecked      =   "Form1.frx":2E21
         PicDisabled     =   "Form1.frx":306F
         PicUnchecked    =   "Form1.frx":31CF
      End
      Begin Project1.OptionEx OptionEx1 
         Height          =   210
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   370
         PicChecked      =   "Form1.frx":3419
         PicDisabled     =   "Form1.frx":3667
         PicUnchecked    =   "Form1.frx":37C7
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
OptionEx1.Enabled = False
End Sub

Private Sub Command2_Click()
OptionEx1.Enabled = True
End Sub
