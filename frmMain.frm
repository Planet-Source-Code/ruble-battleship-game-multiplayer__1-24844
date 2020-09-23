VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sea Battle"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":08A6
   ScaleHeight     =   7125
   ScaleWidth      =   10080
   Begin VB.PictureBox radar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   8160
      Picture         =   "frmMain.frx":6ECE
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   31
      Top             =   4680
      Visible         =   0   'False
      Width           =   9030
      Begin VB.PictureBox waterAnim 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "frmMain.frx":87A5
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   120
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.PictureBox fire 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1800
         Picture         =   "frmMain.frx":8C08
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   569
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   8565
      End
      Begin VB.PictureBox hit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "frmMain.frx":9A6F
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   608
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   9150
      End
   End
   Begin VB.PictureBox radarMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   8160
      Picture         =   "frmMain.frx":B51B
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   32
      Top             =   3480
      Visible         =   0   'False
      Width           =   9030
      Begin VB.PictureBox fireMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   840
         Picture         =   "frmMain.frx":C169
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   569
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   8565
      End
      Begin VB.PictureBox hitMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "frmMain.frx":C8C4
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   608
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   9150
      End
   End
   Begin VB.Frame frameType 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0C4A6&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   480
      TabIndex        =   21
      Top             =   1800
      Width           =   3135
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0C4A6&
         Caption         =   "Single Player"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0C4A6&
         Caption         =   "2 Player Network Server"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   840
         Width           =   2055
      End
      Begin VB.Frame frameGame 
         BackColor       =   &H00F0C4A6&
         Height          =   975
         Left            =   240
         TabIndex        =   24
         Top             =   1500
         Visible         =   0   'False
         Width           =   2655
         Begin VB.TextBox txtServer 
            Height          =   285
            Left            =   240
            TabIndex        =   25
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F0C4A6&
            Caption         =   "Enter IP of Server Computer"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0C4A6&
         Caption         =   "2 Player Network Client"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         Picture         =   "frmMain.frx":D073
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   3135
         Left            =   0
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F0C4A6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Choose Your Game Type Please"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   29
         Top             =   90
         Width           =   2775
      End
   End
   Begin VB.Timer turnTaggle 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   0
   End
   Begin VB.Timer radarAnim 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   0
   End
   Begin VB.Timer missAnim 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer fireAnim 
      Interval        =   100
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer hitAnim 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   2880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0C4A6&
      Caption         =   "Start"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   140
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FF00&
      Height          =   1095
      Left            =   3840
      TabIndex        =   44
      Top             =   1440
      Visible         =   0   'False
      Width           =   6495
      Begin VB.PictureBox tile 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   960
         Picture         =   "frmMain.frx":DFE5
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox water 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1440
         Picture         =   "frmMain.frx":EAEF
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipMid1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2880
         Picture         =   "frmMain.frx":F5F9
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipMid1Mask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3360
         Picture         =   "frmMain.frx":10103
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipDead 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         Picture         =   "frmMain.frx":10C0D
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipDeadMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":11717
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipMid2Mask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4320
         Picture         =   "frmMain.frx":12221
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipMid2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3840
         Picture         =   "frmMain.frx":12D2B
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipLeftEndMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2400
         Picture         =   "frmMain.frx":13835
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipLeftEnd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1920
         Picture         =   "frmMain.frx":1433F
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   58
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipRightEndMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5280
         Picture         =   "frmMain.frx":14E49
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipRightEnd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4800
         Picture         =   "frmMain.frx":15953
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipSingleHorMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5760
         Picture         =   "frmMain.frx":1645D
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   55
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipSingleHor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5760
         Picture         =   "frmMain.frx":16F67
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   54
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipMid1DownMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3360
         Picture         =   "frmMain.frx":17A71
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipMid1Down 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2880
         Picture         =   "frmMain.frx":1857B
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipMid2DownMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4320
         Picture         =   "frmMain.frx":19085
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipMid2Down 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3840
         Picture         =   "frmMain.frx":19B8F
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipDownEndMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2400
         Picture         =   "frmMain.frx":1A699
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   49
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipDownEnd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1920
         Picture         =   "frmMain.frx":1B1A3
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipUpEndMask 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   5280
         Picture         =   "frmMain.frx":1BCAD
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox shipUpEnd 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4800
         Picture         =   "frmMain.frx":1C7B7
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   46
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox waterMiss 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1440
         Picture         =   "frmMain.frx":1D2C1
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.TextBox txtReply 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Height          =   315
      Left            =   120
      MaxLength       =   50
      TabIndex        =   42
      Top             =   6600
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0C4A6&
      BorderStyle     =   0  'None
      Caption         =   "Choose a ship to place on your field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5880
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   3375
      Begin VB.OptionButton Option8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0C4A6&
         Caption         =   "Clear"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   1920
         Width           =   2055
      End
      Begin VB.PictureBox picCurrent 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00F0C4A6&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   240
         ScaleHeight     =   450
         ScaleWidth      =   2865
         TabIndex        =   11
         Top             =   2640
         Width           =   2895
      End
      Begin VB.OptionButton Option7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0C4A6&
         Caption         =   "Destroyer"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton Option6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0C4A6&
         Caption         =   "Submarine"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0C4A6&
         Caption         =   "Cruiser"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0C4A6&
         Caption         =   "Battleship"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   3
         Height          =   3255
         Left            =   0
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F0C4A6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Choose a ship to place on your field"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   130
         TabIndex        =   20
         Top             =   90
         Width           =   3105
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F0C4A6&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Picture         =   "frmMain.frx":1DDCB
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Picture         =   "frmMain.frx":1F019
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Picture         =   "frmMain.frx":20267
         Top             =   810
         Width           =   1095
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Picture         =   "frmMain.frx":214B5
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label lbl7 
         BackColor       =   &H00F0C4A6&
         Caption         =   "4 left"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lbl6 
         BackColor       =   &H00F0C4A6&
         Caption         =   "3 left"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lbl5 
         BackColor       =   &H00F0C4A6&
         Caption         =   "2 left"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbl4 
         BackColor       =   &H00F0C4A6&
         Caption         =   "1 left"
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F0C4A6&
         Caption         =   "Current Selection"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   960
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
   End
   Begin VB.Frame frameField1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Your Fleet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4970
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   4770
      Begin VB.PictureBox field1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   4530
         Left            =   120
         ScaleHeight     =   300
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   2
         Top             =   360
         Width           =   4530
         Begin VB.Shape box 
            BorderColor     =   &H00F0C4A6&
            BorderWidth     =   5
            Height          =   450
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   1800
         End
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         Height          =   4970
         Left            =   0
         Top             =   0
         Width           =   4770
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00F0C4A6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Your Fleet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   38
         Top             =   70
         Width           =   4525
      End
   End
   Begin VB.PictureBox radarPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000C0C0&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   120
      Picture         =   "frmMain.frx":22703
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Frame frameField2 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4970
      Left            =   5160
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   4770
      Begin VB.PictureBox field2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   4530
         Left            =   120
         MouseIcon       =   "frmMain.frx":22DB1
         ScaleHeight     =   300
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   4
         Top             =   360
         Width           =   4530
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F0C4A6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Opponent's Fleet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   39
         Top             =   70
         Width           =   4525
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         BorderWidth     =   3
         Height          =   4970
         Left            =   0
         Top             =   0
         Width           =   4770
      End
   End
   Begin VB.TextBox txtReceive 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   910
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      Top             =   360
      Visible         =   0   'False
      Width           =   8590
   End
   Begin VB.Image bigShip 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1155
      Left            =   1320
      Picture         =   "frmMain.frx":23093
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   8595
   End
   Begin VB.Label lblReply 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00F0C4A6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type in your message here and press the return key:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   43
      Top             =   6360
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.Label lblReceive 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00F0C4A6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Incoming messages from your opponent:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1320
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0C4A6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   400
      TabIndex        =   0
      Top             =   5100
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0C4A6&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status of the game:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   5170
      TabIndex        =   18
      Top             =   6480
      Visible         =   0   'False
      Width           =   4745
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'* Made by Ruble
'* ruble_19@yahoo.com
'**********************
'*
'* You may use this code any way you like.
'*
'* Description:
'* This program is an old game of Battleship which you can play
'* 1 player with the computer and two player over the network or
'* internet providing you know the IP of the server computer
'******************************************************************

Option Explicit

Dim mess As String      'used for displaying game status messages
Dim curX As Integer     'latest shot X coord
Dim curY As Integer     'latest shot Y coord
Dim user As String      'opponents nick
Dim user1 As Boolean    'used for turn determination (true if user1's turn)
Dim user2 As Boolean    'used for turn determination (true if user2's turn)
Dim zone1(0 To max - 1, 0 To max - 1)   'Stores info on where your ships are located
Dim zone2(0 To max - 1, 0 To max - 1)   'Stores info on where you have shot before
Dim zoneComp(0 To max - 1, 0 To max - 1)    'Stores info on where computer ships are located
Dim zoneComp2(0 To max - 1, 0 To max - 1)   'Stores info on where the computer shot before
Dim zone1flame(1 To 20) As flame    'Stores flames
Dim ships(1 To 10) As shipClass     'Stores your ships
Dim compShips(1 To 10) As shipClass 'Stores computer ships
Function allSet() As Boolean
'Function dermines if all of the ships have been placed
'in order to start the game

Dim i As Integer    'counter

allSet = True   'initialize to true

'Loop through all ships to see if any of them
'are not set
For i = 1 To 10
    If ships(i) Is Nothing Then
        allSet = False
        Exit Function
    End If
Next i
End Function


Sub changeTurn()
'This sub changes the turns

'Set every boolean to opposite of what it is now
user1 = Not user1
user2 = Not user2
field2.Enabled = Not field2.Enabled

'user1 = true means your turn
If user1 Then
    'Set state caption
    lblState.Caption = "It is your turn!"
    'Enable the turnTaggle timer which makes the state caption blink
    turnTaggle.Enabled = True
Else
    'Set state caption
    lblState.Caption = "It is " & user & "'s turn!"
    'Disable blinking
    turnTaggle.Enabled = False
    lblState.BackColor = &HF0C4A6
    
    'If gameType = single mode then computer shoots
    If gameType = 1 Then
        compShoot
    End If
End If
End Sub
Sub compShoot()
On Error Resume Next

Dim cX As Integer           'current X
Dim cY As Integer           'current Y
Dim tempX As Integer        'temp x
Dim tempY As Integer        'temp y
Dim headX As Integer        'head x of ship
Dim headY As Integer        'head y of ship
Dim i As Integer            'counter
Dim j As Integer            'counter
Dim k As Integer            'counter
Dim r As Integer            'return value
Dim counter As Integer      'counter
Dim xStart As Integer       'used for shot selection
Dim yStart As Integer       'used for shot selection
Dim xEnd As Integer         'used for shot selection
Dim yEnd As Integer         'used for shot selection
Dim minX As Integer         'used for shot selection
Dim maxX As Integer         'used for shot selection
Dim minY As Integer         'used for shot selection
Dim maxY As Integer         'used for shot selection
Dim mark As Integer         'used for shot selection

'If no shot was predetermined then randomly choose one
If nextShotX = -1 Then
    Do
        Randomize
        cX = Int(Rnd() * 10)
        Randomize
        cY = Int(Rnd() * 10)
    Loop While zoneComp2(cX, cY) > 0
Else
    cX = nextShotX
    cY = nextShotY
End If

'Mark shot
zoneComp2(cX, cY) = 1


tempX = cX
tempY = cY

'Hit
If zone1(cX, cY) > 0 And zone1(cX, cY) < 5 Then
    zone1(cX, cY) = zone1(cX, cY) + 4
    
    'Find ship headX
    Do
        tempX = tempX - 1
        If tempX < 0 Then
            Exit Do
        End If
    Loop Until zone1(tempX, cY) = 0
    headX = tempX + 1
    
    'Find ship headY
    Do
        tempY = tempY - 1
        If tempY < 0 Then
            Exit Do
        End If
    Loop Until zone1(cX, tempY) = 0
    headY = tempY + 1
    
    'Find ship from array
    For i = 1 To 10
        If ships(i).headX = headX And ships(i).headY = headY Then
            'Wound ship
            r = ships(i).cripple()
            
            'Mark as hit
            zoneComp2(cX, cY) = 2
            
            'If dead
            If r Then
                If isOver() Then
                    drawRemainComp
                    MsgBox ("You have been defeated by " & user & vbNewLine & "You Suck!")
                    Unload frmMain
                Else
                    'Get area in and around the ship and mark as shot
                    'since there cannot be any ships around a ship
                    mark = ships(i).length
                    If ships(i).direction = vbKeyRight Then
                        If ships(i).headX > 0 Then
                            xStart = ships(i).headX - 1
                        Else
                            xStart = 0
                        End If
                        If ships(i).headX + mark < max Then
                            xEnd = ships(i).headX + mark
                        Else
                            xEnd = max - 1
                        End If
                        If ships(i).headY > 0 Then
                            yStart = ships(i).headY - 1
                        Else
                            yStart = 0
                        End If
                        If ships(i).headY < max - 1 Then
                            yEnd = ships(i).headY + 1
                        Else
                            yEnd = max - 1
                        End If
                     Else
                        If ships(i).headX > 0 Then
                            xStart = ships(i).headX - 1
                        Else
                            xStart = 0
                        End If
                        If ships(i).headX < max - 1 Then
                            xEnd = ships(i).headX + 1
                        Else
                            xEnd = max - 1
                        End If
                        If ships(i).headY > 0 Then
                            yStart = ships(i).headY - 1
                        Else
                            yStart = 0
                        End If
                        If ships(i).headY + mark < max Then
                            yEnd = ships(i).headY + mark
                        Else
                            yEnd = max - 1
                        End If
                    End If
                    For k = xStart To xEnd
                        For j = yStart To yEnd
                            zoneComp2(k, j) = 1
                        Next j
                    Next k
                    
                    'Get next shot randomly
                    Do
                        Randomize
                        nextShotX = Int(Rnd() * 10)
                        Randomize
                        nextShotY = Int(Rnd() * 10)
                    Loop While zoneComp2(nextShotX, nextShotY) > 0
                End If
            Else    'Crippled
                minX = cX - 1
                maxX = cX + 1
                minY = cY - 1
                maxY = cY + 1
                
                'If it is the first part of ship hit then mark areas around it
                'as potential candidates for hit
                If zoneComp2(minX, cY) < 2 And zoneComp2(maxX, cY) < 2 And zoneComp2(cX, maxY) < 2 And zoneComp2(cX, minY) < 2 Then
                    If zoneComp2(minX, cY) <> 1 And zoneComp2(minX, cY) <> 2 Then
                        zoneComp2(minX, cY) = 5
                    End If
                    If zoneComp2(maxX, cY) <> 1 And zoneComp2(maxX, cY) <> 2 Then
                        zoneComp2(maxX, cY) = 5
                    End If
                    If zoneComp2(cX, maxY) <> 1 And zoneComp2(cX, maxY) <> 2 Then
                        zoneComp2(cX, maxY) = 5
                    End If
                    If zoneComp2(cX, minY) <> 1 And zoneComp2(cX, minY) <> 2 Then
                        zoneComp2(cX, minY) = 5
                    End If
                Else
                    'If this not the first part of ship then mark previously selected
                    'squares for candidates as misses because now you know the
                    'direction of the ship
                    If (zoneComp2(minX, cY) = 2 And minX <> cX) Or (zoneComp2(maxX, cY) = 2 And maxX <> cX) Then
                        If (zoneComp2(minX, cY) = 2 And minX <> cX) Then
                            zoneComp2(maxX, cY) = 5
                        Else
                            zoneComp2(minX, cY) = 5
                        End If
                        zoneComp2(cX, minY) = 1
                        zoneComp2(cX, maxY) = 1
                        zoneComp2(maxX, minY) = 1
                        zoneComp2(maxX, maxY) = 1
                        zoneComp2(minX, minY) = 1
                        zoneComp2(minX, maxY) = 1
                        'You can unrem these bitblt's to see how the comp thinks
                        'BitBlt field1.hDC, cX * 30, minY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                        'BitBlt field1.hDC, cX * 30, maxY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                        'BitBlt field1.hDC, maxX * 30, minY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                        'BitBlt field1.hDC, maxX * 30, maxY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                        'BitBlt field1.hDC, minX * 30, minY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                        'BitBlt field1.hDC, minX * 30, maxY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                    ElseIf (zoneComp2(cX, minY) = 2 And minY <> cY) Or (zoneComp2(cX, maxY) = 2 And maxY <> cY) Then
                        If (zoneComp2(cX, minY) = 2 And minY <> cY) Then
                            zoneComp2(cX, maxY) = 5
                        Else
                            zoneComp2(cX, minY) = 5
                        End If
                        zoneComp2(minX, cY) = 1
                        zoneComp2(maxX, cY) = 1
                        zoneComp2(maxX, minY) = 1
                        zoneComp2(minX, minY) = 1
                        zoneComp2(minX, maxY) = 1
                        zoneComp2(maxX, maxY) = 1
                        'You can unrem these bitblt's to see how the comp thinks
                        'BitBlt field1.hDC, minX * 30, cY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                        'BitBlt field1.hDC, maxX * 30, cY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                        'BitBlt field1.hDC, maxX * 30, minY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                        'BitBlt field1.hDC, maxX * 30, maxY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                        'BitBlt field1.hDC, minX * 30, minY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                        'BitBlt field1.hDC, minX * 30, maxY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
                    End If
                End If
                
                'This gets the next potential shot if any
                For k = 0 To max - 1
                    For j = 0 To max - 1
                        If zoneComp2(k, j) = 5 Then
                            nextShotX = k
                            nextShotY = j
                            changeTurn
                            field1.Refresh
                            
                            'Set new flame
                            zone1flame(curFlame).frame = 0
                            Randomize
                            zone1flame(curFlame).typ = Int(Rnd() * 3) + 1
                            zone1flame(curFlame).x = cX
                            zone1flame(curFlame).y = cY
                            getBackImage cX, cY, zone1flame(curFlame).backImage, zone1flame(curFlame).backImageMask
                            curFlame = curFlame + 1
                            Exit Sub
                        End If
                    Next j
                Next k
                
                nextShotX = -1
                nextShotY = -1
            End If
            
            'Set new flame
            zone1flame(curFlame).frame = 0
            Randomize
            zone1flame(curFlame).typ = Int(Rnd() * 3) + 1
            zone1flame(curFlame).x = cX
            zone1flame(curFlame).y = cY
            getBackImage cX, cY, zone1flame(curFlame).backImage, zone1flame(curFlame).backImageMask
            curFlame = curFlame + 1
            Exit For
        End If
        Next i
Else
    'Miss
    BitBlt field1.hDC, cX * 30, cY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
    
    'See if any candidates exist otherwise set next to -1 so it chooses randomly next time
    For k = 0 To max - 1
        For j = 0 To max - 1
            If zoneComp2(k, j) = 5 Then
                nextShotX = k
                nextShotY = j
                changeTurn
                field1.Refresh
                Exit Sub
            End If
        Next j
    Next k
    nextShotX = -1
    nextShotY = -1
End If
changeTurn
field1.Refresh
End Sub

Sub drawField()
'This sub draws all the tiles on field1 (left field)

Dim i As Integer    'counter
Dim j As Integer    'counter

'Draw tile on each available square
For i = 0 To max - 1
    For j = 0 To max - 1
        BitBlt field1.hDC, i * 30, j * 30, tile.ScaleWidth, tile.ScaleHeight, tile.hDC, 0, 0, vbSrcCopy
    Next j
Next i

field1.Refresh

'Make field1 visible
field1.Visible = True
End Sub
Sub drawRemainComp()
Dim i As Integer

For i = 1 To 10
    If Not compShips(i).isDead Then
        compShips(i).drawShip2
    End If
Next i
End Sub

Sub drawShips()
'This sub draws all the ships on field1 (left field)

Dim i As Integer    'counter

'For each ship draw it
For i = 1 To 10
    ships(i).drawShip
Next
End Sub

Sub drawWater()
'This sub draws water tiles on field1 (left field) and empty tiles on field2 (right field)

Dim i As Integer    'counter
Dim j As Integer    'counter

'Go through each available sqaure
For i = 0 To max - 1
    For j = 0 To max - 1
        BitBlt field1.hDC, i * 30, j * 30, tile.ScaleWidth, tile.ScaleHeight, water.hDC, 0, 0, vbSrcCopy
        BitBlt field2.hDC, i * 30, j * 30, tile.ScaleWidth, tile.ScaleHeight, tile.hDC, 0, 0, vbSrcCopy
    Next j
Next i
field1.Refresh
field2.Refresh
End Sub



Sub getBackImage(ByVal x, ByVal y, ByRef backImage As PictureBox, ByRef backImageMask As PictureBox)
'This sub gets the image and its mask from x and y coords supplied

Dim tempX As Integer    'temp X coord
Dim tempY As Integer    'temp Y coord
Dim headX As Integer    'head X of ship
Dim headY As Integer    'head Y of ship
Dim size As Integer     'size of ship
Dim direc As Integer    'direction of ship
Dim i As Integer        'counter
Dim section As Integer  'current section

tempX = x   'set temp X to x
tempY = y   'set temp Y to y

'Check if coords actually point to a ship
If zone1(tempX, tempY) <> 0 Then
    'Find headX of ship by going backwards from given point
    Do
        tempX = tempX - 1
        If tempX < 0 Then
            Exit Do
        End If
    Loop Until zone1(tempX, y) = 0
    headX = tempX + 1
    
    'Find headY of ship by going backwards from given point
    Do
        tempY = tempY - 1
        If tempY < 0 Then
            Exit Do
        End If
    Loop Until zone1(x, tempY) = 0
    headY = tempY + 1
End If

'Loop through all ships to find the one that is on x and y
For i = 1 To 10
    If ships(i) Is Nothing Then
        'Do nothing
    ElseIf ships(i).headX = headX And ships(i).headY = headY Then
        direc = ships(i).direction
        size = ships(i).length
        section = ships(i).sectionNumber(x, y)
        Exit For
    End If
Next i

'Direction determines the back image
If direc = vbKeyRight Then
    'Size determines the back image
    If size = 1 Then
        Set backImage = shipSingleHor
        Set backImageMask = shipSingleHorMask
    ElseIf size = 2 Then
        If section = 1 Then
            Set backImage = shipLeftEnd
            Set backImageMask = shipLeftEndMask
        ElseIf section = 2 Then
            Set backImage = shipRightEnd
            Set backImageMask = shipRightEndMask
        End If
    ElseIf size = 3 Then
        If section = 1 Then
            Set backImage = shipLeftEnd
            Set backImageMask = shipLeftEndMask
        ElseIf section = 2 Then
            Set backImage = shipMid1
            Set backImageMask = shipMid1Mask
        ElseIf section = 3 Then
            Set backImage = shipRightEnd
            Set backImageMask = shipRightEndMask
        End If
    ElseIf size = 4 Then
        If section = 1 Then
            Set backImage = shipLeftEnd
            Set backImageMask = shipLeftEndMask
        ElseIf section = 2 Then
            Set backImage = shipMid1
            Set backImageMask = shipMid1Mask
        ElseIf section = 3 Then
            Set backImage = shipMid2
            Set backImageMask = shipMid2Mask
        ElseIf section = 4 Then
            Set backImage = shipRightEnd
            Set backImageMask = shipRightEndMask
        End If
    End If
Else
    'Size determines the back image
    If size = 1 Then
        Set backImage = shipSingleHor
        Set backImageMask = shipSingleHorMask
    ElseIf size = 2 Then
        If section = 1 Then
            Set backImage = shipUpEnd
            Set backImageMask = shipUpEndMask
        ElseIf section = 2 Then
            Set backImage = shipDownEnd
            Set backImageMask = shipDownEndMask
        End If
    ElseIf size = 3 Then
        If section = 1 Then
            Set backImage = shipUpEnd
            Set backImageMask = shipUpEndMask
        ElseIf section = 2 Then
            Set backImage = shipMid2Down
            Set backImageMask = shipMid2DownMask
        ElseIf section = 3 Then
            Set backImage = shipDownEnd
            Set backImageMask = shipDownEndMask
        End If
    ElseIf size = 4 Then
        If section = 1 Then
            Set backImage = shipUpEnd
            Set backImageMask = shipUpEndMask
        ElseIf section = 2 Then
            Set backImage = shipMid2Down
            Set backImageMask = shipMid2DownMask
        ElseIf section = 3 Then
            Set backImage = shipMid1Down
            Set backImageMask = shipMid1DownMask
        ElseIf section = 4 Then
            Set backImage = shipDownEnd
            Set backImageMask = shipDownEndMask
        End If
    End If
End If
End Sub

Function isOver() As Boolean
'This function returns true if the game is over

Dim i As Integer    'counter

'Initialize isOver to true
isOver = True

'Go through each ship and test if all are dead
For i = 1 To 10
    If Not ships(i).isDead Then
        isOver = False
        Exit Function
    End If
Next i
End Function
Function isOverComp() As Boolean
'This function returns true if the game is over

Dim i As Integer    'counter

'Initialize isOver to true
isOverComp = True

'Go through each ship and test if all are dead
For i = 1 To 10
    If Not compShips(i).isDead Then
        isOverComp = False
        Exit Function
    End If
Next i
End Function
Sub setComputerShips()
'This sub randomly places the computer ships
'This sub could be done without code repetition
'but I am lazy

Dim OK As Boolean       'OK if valid position
Dim cX As Integer       'current X coord
Dim cY As Integer       'current Y coord
Dim direc As Integer    'direction
Dim i As Integer        'counter
Dim k As Integer        'counter

'Set battleship
OK = False
Do
    'Get random X
    Randomize
    cX = Int(Rnd() * 10)
    'Get random Y
    Randomize
    cY = Int(Rnd() * 10)
    'Get random direction
    direc = Int(Rnd() * 2) + 1
    If direc = 1 Then
        direc = vbKeyRight
    Else
        direc = vbKeyDown
    End If
    
    'Check validity of position
    If checkValidityComp(cX, cY, direc, 4) And zoneComp(cX, cY) = 0 Then
        Set compShips(1) = New shipClass
        compShips(1).setValues cX, cY, 4, direc
        If direc = vbKeyRight Then
            For i = 0 To 3
                zoneComp(cX + i, cY) = 4
            Next i
        Else
            For i = 0 To 3
                zoneComp(cX, cY + i) = 4
            Next i
        End If
        OK = True
    End If
    i = 0
Loop Until OK

'Set two cruisers
For k = 2 To 3
    OK = False
    i = 0
    Do
        Randomize
        cX = Int(Rnd() * 10)
        Randomize
        cY = Int(Rnd() * 10)
        direc = Int(Rnd() * 2) + 1
        If direc = 1 Then
            direc = vbKeyRight
        Else
            direc = vbKeyDown
        End If
        If checkValidityComp(cX, cY, direc, 3) And zoneComp(cX, cY) = 0 Then
            Set compShips(k) = New shipClass
            compShips(k).setValues cX, cY, 3, direc
            If direc = vbKeyRight Then
                For i = 0 To 2
                    zoneComp(cX + i, cY) = 3
                Next i
            Else
                For i = 0 To 2
                    zoneComp(cX, cY + i) = 3
                Next i
            End If
            OK = True
        End If
        i = 0
    Loop Until OK
Next k

'Set three subs
For k = 4 To 6
    OK = False
    i = 0
    Do
        Randomize
        cX = Int(Rnd() * 10)
        Randomize
        cY = Int(Rnd() * 10)
        direc = Int(Rnd() * 2) + 1
        If direc = 1 Then
            direc = vbKeyRight
        Else
            direc = vbKeyDown
        End If
        If checkValidityComp(cX, cY, direc, 2) And zoneComp(cX, cY) = 0 Then
            Set compShips(k) = New shipClass
            compShips(k).setValues cX, cY, 2, direc
            If direc = vbKeyRight Then
                For i = 0 To 1
                    zoneComp(cX + i, cY) = 2
                Next i
            Else
                For i = 0 To 1
                    zoneComp(cX, cY + i) = 2
                Next i
            End If
            OK = True
        End If
        i = 0
    Loop Until OK
Next k

'Set four destroyers
For k = 7 To 10
    OK = False
    i = 0
    Do
        Randomize
        cX = Int(Rnd() * 10)
        Randomize
        cY = Int(Rnd() * 10)
        direc = Int(Rnd() * 2) + 1
        If direc = 1 Then
            direc = vbKeyRight
        Else
            direc = vbKeyDown
        End If
        If checkValidityComp(cX, cY, direc, 1) And zoneComp(cX, cY) = 0 Then
            Set compShips(k) = New shipClass
            compShips(k).setValues cX, cY, 1, direc
            zoneComp(cX, cY) = 1
            OK = True
        End If
        i = 0
    Loop Until OK
Next k
End Sub

Private Sub Command4_Click()
'This is the begin button
'I got example of winsock usage from PSC so thank
'to whoever posted it

If Option1.Value Then
    'Set field for play in single mode
    timer1.Enabled = False
    lblStatus.Visible = False
    Frame1.Visible = True
    lblState.Visible = True
    Command5.Visible = True
    lblStatus.Visible = False
    frameType.Visible = False
    lblStatus.Visible = False
    frameField1.Visible = True
    drawField
    setComputerShips
    user = "Computer"
ElseIf Option2.Value Then
    'As server
    Winsock.Close
    Winsock.LocalPort = 1234
    Winsock.Listen
    'animate the waiting message
    mess = " Waiting for client computer"
    lblStatus.Caption = mess & " ."
    lblStatus.Visible = True
    timer1.Enabled = True
ElseIf Option3.Value Then
    'as client
    Winsock.Close
    If txtServer.Text <> "" Then
        Winsock.Connect txtServer.Text, 1234
        'animate the waiting message
        mess = " Waiting for reply from the server"
        lblStatus.Caption = mess & " ."
        lblStatus.Visible = True
        timer1.Enabled = True
    Else
        MsgBox ("Please indicate the server's IP or Host Name")
        Exit Sub
    End If
End If
End Sub
Private Sub Command5_Click()
'The start button

'Set field for play
field2.Enabled = False
radarAnim.Enabled = True
radarPic.Visible = True
txtReply.Visible = True
lblReply.Visible = True

If gameType = 2 Then
    txtReceive.Visible = True
    lblReceive.Visible = True
Else
    bigShip.Visible = True
End If

Frame1.Visible = False
frameField2.Visible = True
field1.Enabled = False
drawWater
box.Visible = False
field1.Refresh

'If multi player
If gameType = 2 Then
    'Send ready message
    Winsock.SendData "Info:Ready"
    user1 = True
    
    'if both ready, start game
    If user1 And user2 Then
        timer1.Enabled = False
        user2 = True
        user1 = False
        field2.Enabled = False
        changeTurn
    Else
        'waiting message
        mess = " " & user & " is still not ready, please wait"
        lblState.Caption = mess & " ."
        timer1.Enabled = True
    End If
    drawShips
    Command5.Visible = False
Else 'Single player
    field2.Enabled = True
    user1 = True
    user2 = False
    drawShips
    Command5.Visible = False
    txtReceive.Enabled = False
    txtReply.Enabled = False
    'Set state caption
    lblState.Caption = "It is your turn!"
    'Enable the turnTaggle timer which makes the state caption blink
    turnTaggle.Enabled = True
End If
End Sub
Private Sub field1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'field1 (left field)

On Error GoTo handler

Dim tempOption As OptionButton  'current option select
Dim tempLbl As Label            'corresponding label
Dim curShip As Integer          'current ship index
Dim tempX As Integer            'temp x
Dim tempY As Integer            'temp y
Dim headX As Integer            'head x
Dim headY As Integer            'head y
Dim direc As Integer            'direction of ship
Dim size As Integer             'size of ship
Dim i As Integer                'counter
Dim mark As Integer             'current mark (1-4 ship type)
Dim tempWidth As Integer        'temp width of box
Dim clickedX As Integer         'clicked x
Dim clickedY As Integer         'clicked y
Dim r                           'return value

'left button
If Button = 1 Then
    'Find the next available place for a ship in ship array
    'Loop through ship array until one is = nothing
    For curShip = 1 To 9
        If ships(curShip) Is Nothing Then
            Exit For
        End If
    Next
    'Determine what ship you are placing (1-4)
    If Option7.Value Then
        mark = 1
        Set tempOption = Option7
        Set tempLbl = lbl7
    ElseIf Option6.Value Then
        mark = 2
        Set tempOption = Option6
        Set tempLbl = lbl6
    ElseIf Option5.Value Then
        mark = 3
        Set tempOption = Option5
        Set tempLbl = lbl5
    ElseIf Option4.Value Then
        mark = 4
        Set tempOption = Option4
        Set tempLbl = lbl4
    Else
        'erase mark
        mark = 0
    End If
    
    'Erase code
    If mark = 0 Then
        clickedX = Int((box.Left + 10) / 30)
        clickedY = Int(y / 30)
        tempX = clickedX
        tempY = clickedY
        
        'Check if ship is there
        If zone1(tempX, tempY) <> 0 Then
            'Find headX of ship
            Do
                tempX = tempX - 1
                If tempX < 0 Then
                    Exit Do
                End If
            Loop Until zone1(tempX, clickedY) = 0
            headX = tempX + 1
            
            'Find shipY of ship
            Do
                tempY = tempY - 1
                If tempY < 0 Then
                    Exit Do
                End If
            Loop Until zone1(clickedX, tempY) = 0
            headY = tempY + 1
            
        'Find ship with headX and headY
        For i = 1 To 10
            If ships(i) Is Nothing Then
                'Do nothing
            ElseIf ships(i).headX = headX And ships(i).headY = headY Then
                direc = ships(i).direction
                size = ships(i).length
                Set ships(i) = Nothing 'erase ship
                Exit For
            End If
        Next i
        
        'Draw tiles where the ship was
        If direc = vbKeyRight Then
            For i = 0 To size - 1
                BitBlt field1.hDC, (headX + i) * 30, headY * 30, tile.ScaleWidth, tile.ScaleHeight, tile.hDC, 0, 0, vbSrcCopy
                zone1((headX + i), headY) = 0
            Next i
        Else
            For i = 0 To size - 1
                BitBlt field1.hDC, headX * 30, (headY + i) * 30, tile.ScaleWidth, tile.ScaleHeight, tile.hDC, 0, 0, vbSrcCopy
                zone1(headX, headY + i) = 0
            Next i
        End If
        
        'Modify the ship availability captions
        If size = 4 Then
            Option4.Enabled = True
            lbl4.Caption = Val(lbl4.Caption) + 1 & " left"
        ElseIf size = 3 Then
            Option5.Enabled = True
            lbl5.Caption = Val(lbl5.Caption) + 1 & " left"
        ElseIf size = 2 Then
            Option6.Enabled = True
            lbl6.Caption = Val(lbl6.Caption) + 1 & " left"
        ElseIf size = 1 Then
            Option7.Enabled = True
            lbl7.Caption = Val(lbl7.Caption) + 1 & " left"
        End If
        field1.Refresh
        End If
    Else    'set ship
        'if direction is horizontal
        If box.Width >= box.Height Then
            'Check validity
            If checkValidity(Int((box.Left + 10) / 30), Int(y / 30), vbKeyRight, mark) Then
                'play place sound
                r = sndPlaySound(App.Path & "\slap.wav", SND_ASYNC)
                
                'Place a mark on every square where ship placed
                For i = Int((box.Left + 10) / 30) To Int((box.Left + 10) / 30) + mark - 1
                    zone1(i, Int(y / 30)) = mark
                Next i
                
                'Initialize new ship
                Set ships(curShip) = New shipClass
                ships(curShip).setValues Int((box.Left + 10) / 30), Int(y / 30), mark, vbKeyRight
                
                'Draw ship
                ships(curShip).drawShip
            Else
                MsgBox ("Invalid position!")
                Exit Sub
            End If
        Else 'direction is vertical
            If checkValidity(Int(x / 30), Int((box.Top + 10) / 30), vbKeyDown, mark) Then
                'play place sound
                r = sndPlaySound(App.Path & "\slap.wav", SND_ASYNC)
                
                'Place a mark on every square where ship placed
                For i = Int((box.Top + 10) / 30) To Int((box.Top + 10) / 30) + mark - 1
                    zone1(Int(x / 30), i) = mark
                Next i
                
                'Initialize new ship
                Set ships(curShip) = New shipClass
                ships(curShip).setValues Int(x / 30), Int((box.Top + 10) / 30), mark, vbKeyDown
                
                'Draw ship
                ships(curShip).drawShip
            Else
                MsgBox ("Invalid position!")
                Exit Sub
            End If
        End If
        
        'Modify ship availablity captions and set focus to next available ship option
        tempLbl.Caption = Val(Left(tempLbl.Caption, 1)) - 1 & " left"
        If Left(tempLbl.Caption, 1) = "0" Then
            tempOption.Enabled = False
            If Option4.Enabled Then
               Option4.SetFocus
            ElseIf Option5.Enabled Then
               Option5.SetFocus
            ElseIf Option6.Enabled Then
               Option6.SetFocus
            ElseIf Option7.Enabled Then
               Option7.SetFocus
            Else
               Option8.SetFocus
            End If
        End If
    End If
Else 'right button
    'Change direction of box
    tempWidth = box.Width
    box.Width = box.Height
    box.Height = tempWidth
End If

'Check if all ships are set
If allSet() Then
    Command5.Enabled = True
Else
    Command5.Enabled = False
End If
Exit Sub
'This can only happen if ship is being placed out of bound
'but should never get here
handler:
    MsgBox ("Out of bounds error")
End Sub
Function checkValidity(ByVal x As Integer, ByVal y As Integer, ByVal dir As Integer, ByVal mark As Integer) As Boolean
'This long function checks if the placement of the ship
'is valid (in bounds and at least one square away from other ships

Dim xStart As Integer   'ship start X coord
Dim yStart As Integer   'ship start Y coord
Dim xEnd As Integer     'ship end X coord
Dim yEnd As Integer     'ship end Y coord
Dim i As Integer        'counter 1
Dim k As Integer        'counter 2

'Direction determines how to look for start and end of ship
If dir = vbKeyRight Then
    'These if's set the coord Start and End variables to be one square over
    'where they actually are. This is how I check if the ship being placed
    'is at least one square away from the others
    'Also makes surtain you don't go out of bounds
    If x > 0 Then
        xStart = x - 1
    Else
        xStart = 0
    End If
    If x + mark < max Then
        xEnd = x + mark
    Else
        xEnd = max - 1
    End If
    If y > 0 Then
        yStart = y - 1
    Else
        yStart = 0
    End If
    If y < max - 1 Then
        yEnd = y + 1
    Else
        yEnd = max - 1
    End If
    
    'Check if all clear
    For i = xStart To xEnd
        For k = yStart To yEnd
            If zone1(i, k) <> 0 Then
                checkValidity = False
                Exit Function
            End If
        Next k
    Next i
Else
    'Same as above but different direction
    If x > 0 Then
        xStart = x - 1
    Else
        xStart = 0
    End If
    If x < max - 1 Then
        xEnd = x + 1
    Else
        xEnd = max - 1
    End If
    If y > 0 Then
        yStart = y - 1
    Else
        yStart = 0
    End If
    If y + mark < max Then
        yEnd = y + mark
    Else
        yEnd = max - 1
    End If
    For i = xStart To xEnd
        For k = yStart To yEnd
            If zone1(i, k) <> 0 Then
                checkValidity = False
                Exit Function
            End If
        Next k
    Next i
End If
'if still in the function then all clear
checkValidity = True
End Function
Function checkValidityComp(ByVal x As Integer, ByVal y As Integer, ByVal dir As Integer, ByVal mark As Integer) As Boolean
'This is basically the same function as checkValidity
'but it is modified since computer can actually go out of bounds
'and you have to modify a bit, but the comments in checkValidity hold true
'here As well (i am lazy)

Dim xStart As Integer
Dim yStart As Integer
Dim xEnd As Integer
Dim yEnd As Integer
Dim i As Integer
Dim k As Integer

'Check for in bounds
If x + mark < max And y + mark < max Then
  If dir = vbKeyRight Then
    If x > 0 Then
        xStart = x - 1
    Else
        xStart = 0
    End If
    If x + mark < max Then
        xEnd = x + mark
    Else
        xEnd = max - 1
    End If
    If y > 0 Then
        yStart = y - 1
    Else
        yStart = 0
    End If
    If y < max - 1 Then
        yEnd = y + 1
    Else
        yEnd = max - 1
    End If
    For i = xStart To xEnd
        For k = yStart To yEnd
            If zoneComp(i, k) <> 0 Then
                checkValidityComp = False
                Exit Function
            End If
        Next k
    Next i
  Else
    If x > 0 Then
        xStart = x - 1
    Else
        xStart = 0
    End If
    If x < max - 1 Then
        xEnd = x + 1
    Else
        xEnd = max - 1
    End If
    If y > 0 Then
        yStart = y - 1
    Else
        yStart = 0
    End If
    If y + mark < max Then
        yEnd = y + mark
    Else
        yEnd = max - 1
    End If
    For i = xStart To xEnd
        For k = yStart To yEnd
            If zoneComp(i, k) <> 0 Then
                checkValidityComp = False
                Exit Function
            End If
        Next k
    Next i
  End If
  checkValidityComp = True
Else
  checkValidityComp = False
End If
End Function
Private Sub field1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Box movement sub

'Make box visible
If Not box.Visible Then
    box.Visible = True
End If

'Make cursor be in the middle of the box
box.Left = x - box.Width / 2
box.Top = y - box.Height / 2

'Make sure the box stays in bounds
If box.Left < 0 Then
    box.Left = 0
End If
If box.Top < 0 Then
    box.Top = 0
End If
If box.Left + box.Width > 300 Then
    box.Left = 300 - box.Width
End If
If box.Top + box.Height > 300 Then
    box.Top = 300 - box.Height
End If
End Sub

Private Sub field2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Shot sub

Dim r As Integer    'return value
Dim i As Integer
Dim tempX As Integer
Dim tempY As Integer
Dim headX As Integer
Dim headY As Integer


'multi-player
If gameType = 2 Then
    curX = Int(x / 30)
    curY = Int(y / 30)
    
    'If empty square
    If zone2(curX, curY) = 0 Then
        changeTurn
        zone2(curX, curY) = 1
        'Send shot message to opponent
        Winsock.SendData "Shot:" & curX & "," & curY
    End If
Else    'single player
    curX = Int(x / 30)
    curY = Int(y / 30)
    
    field2.Enabled = False
    'if empty square
    If zone2(curX, curY) = 0 Then
        zone2(curX, curY) = 1
        
        'Check shot result
        If zoneComp(curX, curY) > 0 And zoneComp(curX, curY) < 5 Then
            'animate explosion
            hitAnim.Enabled = True
            'play hit sound
            r = sndPlaySound(App.Path & "\bang.wav", SND_ASYNC)
            tempX = curX
            tempY = curY
            Do
                tempX = tempX - 1
                If tempX < 0 Then
                    Exit Do
                End If
            Loop Until zoneComp(tempX, curY) = 0
            headX = tempX + 1
            Do
                tempY = tempY - 1
                If tempY < 0 Then
                    Exit Do
                End If
            Loop Until zoneComp(curX, tempY) = 0
            headY = tempY + 1
            For i = 1 To 10
                If compShips(i).headX = headX And compShips(i).headY = headY Then
                    r = compShips(i).cripple
                    If r Then
                        isDead = True
                    End If
                End If
            Next i
        Else
            'animate miss
            missAnim.Enabled = True
            'play miss sound
            r = sndPlaySound(App.Path & "\miss.wav", SND_ASYNC)
        End If
    Else
        field2.Enabled = True
    End If
End If
End Sub

Private Sub field2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'Set mousepointer to crosshairs
field2.MousePointer = 99
End Sub


Private Sub fireAnim_Timer()
'Sub that animates the fire

Dim i As Integer    'counter
i = 1

'Here is how typ is used to check when to stop looping
'Loop through all fires and draw a frame of each
Do Until zone1flame(i).typ = 0
    'Draw fire mask
    BitBlt field1.hDC, zone1flame(i).x * 30, zone1flame(i).y * 30, 30, 30, fireMask.hDC, zone1flame(i).frame * 30, 0, vbSrcAnd
    'Draw fire
    BitBlt field1.hDC, zone1flame(i).x * 30, zone1flame(i).y * 30, 30, 30, fire.hDC, zone1flame(i).frame * 30, 0, vbSrcPaint
    
    'Increment frame
    zone1flame(i).frame = zone1flame(i).frame + 1
    
    'Reset frame
    If zone1flame(i).frame = 19 Then
        zone1flame(i).frame = 0
    End If
    
    'Increment counter
    i = i + 1
    
    'Should almost never be greater then 20 only at the end of game
    If i > 20 Then
        fireAnim.Enabled = False
        Exit Sub
    End If
Loop
field1.Refresh

'Reset counter
i = 1

'For each fire draw backImage and its mask
Do Until zone1flame(i).typ = 0
    BitBlt field1.hDC, zone1flame(i).x * 30, zone1flame(i).y * 30, tile.ScaleWidth, tile.ScaleHeight, zone1flame(i).backImageMask.hDC, 0, 0, vbSrcAnd
    BitBlt field1.hDC, zone1flame(i).x * 30, zone1flame(i).y * 30, tile.ScaleWidth, tile.ScaleHeight, zone1flame(i).backImage.hDC, 0, 0, vbSrcPaint
    i = i + 1
Loop
End Sub

Private Sub Form_Load()
'Set user status to not ready
user1 = False
user2 = False

'Set curFlame
curFlame = 1

'Set radar frame number
frameNum3 = 1

'Center frameType and status label
frameType.Left = Me.Width / 2 - frameType.Width / 2
frameType.Top = Me.Height / 2 - frameType.Height / 2 - 1000
lblStatus.Left = frameType.Left - 70
lblStatus.Top = frameType.Top + frameType.Height + 300

'Single user default
Option1.Value = True

'Initialize compShot
nextShotX = -1
End Sub

Private Sub hitAnim_Timer()
'Sub that animates explosion

Dim r As Integer    'return value

'Draw explosion frame
BitBlt field2.hDC, curX * 30, curY * 30, 30, 30, hitMask.hDC, frameNum * 32, 0, vbSrcAnd
BitBlt field2.hDC, curX * 30, curY * 30, 30, 30, hit.hDC, frameNum * 32, 0, vbSrcPaint

'Increment flame
frameNum = frameNum + 1

field2.Refresh

'Stop animation
If frameNum = 19 Then
    frameNum = 0
    hitAnim.Enabled = False
    BitBlt field2.hDC, curX * 30, curY * 30, tile.ScaleWidth, tile.ScaleHeight, shipDead.hDC, 0, 0, vbSrcCopy
    field2.Refresh
    
    'Play sunk sound if dead
    If isDead Then
        r = sndPlaySound(App.Path & "\sunk.wav", SND_ASYNC)
        If gameType = 1 Then
            MsgBox ("You sunk " & user & "'s ship")
            If isOverComp Then
                MsgBox ("You have defeated " & user & vbNewLine & "Great Job!")
                Unload frmMain
            End If
        End If
        isDead = False
    End If
    
    'Change turn if single player
    If gameType = 1 Then
        field2.Enabled = True
        changeTurn
    End If
End If
End Sub

Private Sub missAnim_Timer()
'sub that animates a miss

'Draw miss frame
BitBlt field2.hDC, curX * 30, curY * 30, 30, 30, waterAnim.hDC, frameNum * 30, 0, vbSrcCopy

'Increment miss frame number
frameNum = frameNum + 1

field2.Refresh

'Stop animation
If frameNum = 4 Then
    frameNum = 0
    missAnim.Enabled = False
    BitBlt field2.hDC, curX * 30, curY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
    field2.Refresh
    
    'Change turn if single player
    If gameType = 1 Then
        field2.Enabled = True
        changeTurn
    End If
End If
End Sub

Private Sub Option1_Click()
'Set single player variables
frameGame.Visible = False
gameType = 1
End Sub

Private Sub Option2_Click()
'Set multi player variables
frameGame.Visible = False
gameType = 2
End Sub

Private Sub Option3_Click()
'Set multi player variables
frameGame.Visible = True
txtServer.Text = Winsock.LocalIP
gameType = 2
End Sub

Private Sub Option4_Click()
'Battleship option

box.BorderColor = &HF0C4A6

'Draw current ship
picCurrent.Cls
BitBlt picCurrent.hDC, 0, 0, tile.ScaleWidth, tile.ScaleHeight, shipLeftEndMask.hDC, 0, 0, vbSrcAnd
BitBlt picCurrent.hDC, 0, 0, tile.ScaleWidth, tile.ScaleHeight, shipLeftEnd.hDC, 0, 0, vbSrcPaint
BitBlt picCurrent.hDC, 30, 0, tile.ScaleWidth, tile.ScaleHeight, shipMid1Mask.hDC, 0, 0, vbSrcAnd
BitBlt picCurrent.hDC, 30, 0, tile.ScaleWidth, tile.ScaleHeight, shipMid1.hDC, 0, 0, vbSrcPaint
BitBlt picCurrent.hDC, 60, 0, tile.ScaleWidth, tile.ScaleHeight, shipMid2Mask.hDC, 0, 0, vbSrcAnd
BitBlt picCurrent.hDC, 60, 0, tile.ScaleWidth, tile.ScaleHeight, shipMid2.hDC, 0, 0, vbSrcPaint
BitBlt picCurrent.hDC, 90, 0, tile.ScaleWidth, tile.ScaleHeight, shipRightEndMask.hDC, 0, 0, vbSrcAnd
BitBlt picCurrent.hDC, 90, 0, tile.ScaleWidth, tile.ScaleHeight, shipRightEnd.hDC, 0, 0, vbSrcPaint
picCurrent.Refresh

'Change box size
box.Width = 30 * 4
box.Height = 30
End Sub

Private Sub Option5_Click()
'Cruiser option

box.BorderColor = &HF0C4A6

'draw current ship
picCurrent.Cls
BitBlt picCurrent.hDC, 0, 0, tile.ScaleWidth, tile.ScaleHeight, shipLeftEndMask.hDC, 0, 0, vbSrcAnd
BitBlt picCurrent.hDC, 0, 0, tile.ScaleWidth, tile.ScaleHeight, shipLeftEnd.hDC, 0, 0, vbSrcPaint
BitBlt picCurrent.hDC, 30, 0, tile.ScaleWidth, tile.ScaleHeight, shipMid1Mask.hDC, 0, 0, vbSrcAnd
BitBlt picCurrent.hDC, 30, 0, tile.ScaleWidth, tile.ScaleHeight, shipMid1.hDC, 0, 0, vbSrcPaint
BitBlt picCurrent.hDC, 60, 0, tile.ScaleWidth, tile.ScaleHeight, shipRightEndMask.hDC, 0, 0, vbSrcAnd
BitBlt picCurrent.hDC, 60, 0, tile.ScaleWidth, tile.ScaleHeight, shipRightEnd.hDC, 0, 0, vbSrcPaint
picCurrent.Refresh

'Change box size
box.Width = 30 * 3
box.Height = 30
End Sub


Private Sub Option6_Click()
'sub option

box.BorderColor = &HF0C4A6

'Draw current ship
picCurrent.Cls
BitBlt picCurrent.hDC, 0, 0, tile.ScaleWidth, tile.ScaleHeight, shipLeftEndMask.hDC, 0, 0, vbSrcAnd
BitBlt picCurrent.hDC, 0, 0, tile.ScaleWidth, tile.ScaleHeight, shipLeftEnd.hDC, 0, 0, vbSrcPaint
BitBlt picCurrent.hDC, 30, 0, tile.ScaleWidth, tile.ScaleHeight, shipRightEndMask.hDC, 0, 0, vbSrcAnd
BitBlt picCurrent.hDC, 30, 0, tile.ScaleWidth, tile.ScaleHeight, shipRightEnd.hDC, 0, 0, vbSrcPaint
picCurrent.Refresh

'Change box size
box.Width = 30 * 2
box.Height = 30
End Sub


Private Sub Option7_Click()
'Destroyer option

box.BorderColor = &HF0C4A6

'Draw current ship
picCurrent.Cls
BitBlt picCurrent.hDC, 0, 0, tile.ScaleWidth, tile.ScaleHeight, shipSingleHorMask.hDC, 0, 0, vbSrcAnd
BitBlt picCurrent.hDC, 0, 0, tile.ScaleWidth, tile.ScaleHeight, shipSingleHor.hDC, 0, 0, vbSrcPaint
picCurrent.Refresh

'Change box size
box.Width = 30 * 1
box.Height = 30
End Sub


Private Sub Option8_Click()
'erase option

box.BorderColor = vbRed

'draw current ship
picCurrent.Cls
BitBlt picCurrent.hDC, 0, 0, tile.ScaleWidth, tile.ScaleHeight, tile.hDC, 0, 0, vbSrcCopy

'Change box size
box.Width = 30 * 1
box.Height = 30
End Sub

Private Sub radarAnim_Timer()
'sub that animates the radar

'Draw frame of radar
BitBlt radarPic.hDC, 0, 0, 75, 75, radarMask.hDC, frameNum3 * 75, 0, vbSrcAnd
BitBlt radarPic.hDC, 0, 0, 75, 75, radar.hDC, frameNum3 * 75, 0, vbSrcPaint

'Increment radar frame
frameNum3 = frameNum3 + 1

radarPic.Refresh

'Reset radar frame
If frameNum3 = 19 Then
    frameNum3 = 0
End If
End Sub

Private Sub timer1_Timer()
'This timer animates the waiting messages

If lblStatus.Caption = mess & " ." Then
    lblStatus.Caption = mess & " .."
ElseIf lblStatus.Caption = mess & " .." Then
    lblStatus.Caption = mess & " ..."
ElseIf lblStatus.Caption = mess & " ..." Then
    lblStatus.Caption = mess & " ."
End If

If lblState.Caption = mess & " ." Then
    lblState.Caption = mess & " .."
ElseIf lblState.Caption = mess & " .." Then
    lblState.Caption = mess & " ..."
ElseIf lblState.Caption = mess & " ..." Then
    lblState.Caption = mess & " ."
End If
End Sub




Private Sub turnTaggle_Timer()
'This sub makes the state caption blink

If lblState.BackColor = &HC0C000 Then
    lblState.BackColor = &HF0C4A6
Else
    lblState.BackColor = &HC0C000
End If
End Sub

Private Sub txtReceive_Change()
'Scroll to end of textbox
txtReceive.SelStart = Len(txtReceive.Text)
End Sub

Private Sub txtReceive_GotFocus()
radarPic.SetFocus
End Sub


Private Sub txtReply_KeyPress(KeyAscii As Integer)
On Error GoTo handler

'On enter send message
If KeyAscii = vbKeyReturn And txtReply.Text <> "" Then
    Winsock.SendData "Message:" & txtReply.Text
    
    'Display message in message window
    If txtReceive.Text <> "" Then
        txtReceive.Text = txtReceive.Text & vbNewLine & "You -> " & txtReply.Text
    Else
        txtReceive.Text = txtReceive.Text & "You -> " & txtReply.Text
    End If
    txtReply.Text = ""
End If
Exit Sub

'This is only used if message window gets filled to much
'it never happened to me yet
handler:
    txtReceive.Text = txtReply.Text
End Sub


Private Sub Winsock_Close()
'End game if connection is lost

MsgBox (user & " has quit his game!")
End
End Sub

Private Sub Winsock_Connect()
'On connection

'hide status label
timer1.Enabled = False
lblStatus.Visible = False

'Send news of successful connection
Winsock.SendData "Status:Connected"

frameType.Visible = False   'hide type frame
frameField1.Visible = True  'show field1 (left field)
drawField   'draw empty tiles

'Set game field
Frame1.Visible = True
lblState.Visible = True
Command5.Visible = True
lblStatus.Visible = False
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
'Accept connection

Winsock.Close
Winsock.Accept requestID
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
'Data arrival sub

Dim info As String  'message that is received
Dim r               'return value

'Get info
Winsock.GetData info

If info = "Status:Connected" Then
    'Set game field
    timer1.Enabled = False
    Frame1.Visible = True
    lblState.Visible = True
    Command5.Visible = True
    lblStatus.Visible = False
    lblStatus.Visible = False
    
    'Send acknowledge message
    Winsock.SendData "Status:Gotit"
    
    'Prompt for nick until it is not ""
    Do
        r = InputBox("Please enter your name:", "Enter Name", GetUser)
    Loop While r = ""
    
    'Send nick to opponent
    Winsock.SendData "User:" & r
    
    drawField
    frameType.Visible = False
    
ElseIf info = "Status:Gotit" Then
    'Get nick until not ""
    Do
        r = InputBox("Please enter your name:", "Enter Name", GetUser)
    Loop While r = ""
    
    'Send nick to opponent
    Winsock.SendData "User:" & r
    field2.Enabled = True
    drawField
    frameType.Visible = False
Else
    'Parse message
    parse (info)
End If
End Sub

Function GetUser() As String
'This function returns the current logged in user
'I got this from PSC somewhere so thanks to whoever posted it
    
Dim lpUserID As String
Dim nBuffer As Long
Dim ret As Long

lpUserID = String(25, 0)
nBuffer = 25
ret = GetUserName(lpUserID, nBuffer)
If ret Then
    GetUser$ = lpUserID$
End If
End Function
Sub parse(info As String)
'This sub pareses all the received from opponent

On Error GoTo handler

Dim msgtype As String   'stores the type of message
Dim msg As String       'stores the msg itself
Dim cX As Integer       'current X coord
Dim cY As Integer       'current Y coord
Dim tempX As Integer    'temp X coord
Dim tempY As Integer    'temp Y coord
Dim headX As Integer    'head X coord
Dim headY As Integer    'head Y coord
Dim r                   'return value
Dim i As Integer        'counter
Dim sendStr As String   'Message to send
Dim tempship As shipClass   'used to draw remaining ships
Dim tempstr As String       'temperary string
Dim tempdir As Integer      'used to draw remaining ships
Dim tempsize As Integer     'used to draw remaining ships
Dim tempheadx As Integer    'used to draw remaining ships
Dim tempheady As Integer    'used to draw remaining ships

'Get the type of message (string before ':')
msgtype = Left(info, InStr(1, info, ":") - 1)

'Get the message itself (string after ':')
msg = Mid(info, InStr(1, info, ":") + 1)

'Parse the messages
'"User" means player joined the game
If msgtype = "User" Then
    MsgBox (msg & " has joined the game!")
    user = msg  'set user to opponent nick
    frameType.Visible = False   'hide type frame
    frameField1.Visible = True  'show field1 (left field)
    drawField   'draw empty tiles
ElseIf msgtype = "Shot" Then
    'Shot comes in format x,y so parse that
    cX = Left(msg, InStr(1, msg, ",") - 1)
    cY = Mid(msg, InStr(1, msg, ",") + 1)
    
    'set temp coords
    tempX = cX
    tempY = cY
    
    'zone1(x,y) < 5 that means it's a hit
    If zone1(cX, cY) > 0 And zone1(cX, cY) < 5 Then
        'This is really useless now I though it would be used but it was not
        'It sets the zone(x,y) to plus 4 indicating hit section
        zone1(cX, cY) = zone1(cX, cY) + 4
        
        'Find headX of ship
        Do
            tempX = tempX - 1
            If tempX < 0 Then
                Exit Do
            End If
        Loop Until zone1(tempX, cY) = 0
        headX = tempX + 1
       
        'Find headY of ship
        Do
            tempY = tempY - 1
            If tempY < 0 Then
                Exit Do
            End If
        Loop Until zone1(cX, tempY) = 0
        headY = tempY + 1
        
        'Find the ship with headX and headY
        For i = 1 To 10
            If ships(i).headX = headX And ships(i).headY = headY Then
                'Cripple the ship
                r = ships(i).cripple()
                
                'If r = true then ship is dead
                If r Then
                    'Check if all ships are dead
                    If isOver() Then
                        'Send opponent news of their victory
                        Winsock.SendData "Result:Kayuk" 'its russian
                    Else
                        'Send news of sunken ship to opponent
                        Winsock.SendData "Result:Dead"
                    End If
                Else
                    'Send news of crippled ship to opponent
                    Winsock.SendData "Result:Wound"
                End If
                
                'Set the next available flame
                zone1flame(curFlame).frame = 0
                Randomize
                'Used before, but now basically useless (check fireAnim for usage)
                zone1flame(curFlame).typ = Int(Rnd() * 3) + 1
                zone1flame(curFlame).x = cX
                zone1flame(curFlame).y = cY
                'Send backImage and backImageMask byRef
                getBackImage cX, cY, zone1flame(curFlame).backImage, zone1flame(curFlame).backImageMask
                'Increment current flame
                curFlame = curFlame + 1
                Exit For
            End If
        Next i
    Else
        'Draw miss tile on field1 (left field)
        BitBlt field1.hDC, cX * 30, cY * 30, tile.ScaleWidth, tile.ScaleHeight, waterMiss.hDC, 0, 0, vbSrcCopy
        'Send new of a miss to opponent
        Winsock.SendData "Result:Miss"
    End If
    changeTurn
    field1.Refresh
'Result is the result of the shot returned from opponent
ElseIf msgtype = "Result" Then
    If msg = "Dead" Then
        'animate explosion
        hitAnim.Enabled = True
        isDead = True
        
        'play hit sound
        r = sndPlaySound(App.Path & "\bang.wav", SND_ASYNC)
        MsgBox ("You sunk " & user & "'s ship")
    ElseIf msg = "Kayuk" Then
        'animate explosion
        hitAnim.Enabled = True
        isDead = True
        
        'This is used to send ship data to the opponent in order to display
        'the ones that are still remaining
        sendStr = "Remain:"
        For i = 1 To 10
            If Not ships(i).isDead Then
                If sendStr <> "Remain:" Then
                   sendStr = sendStr & ":" & ships(i).headX & "," & ships(i).headY & "," & ships(i).direction & "," & ships(i).length
                Else
                   sendStr = sendStr & ships(i).headX & "," & ships(i).headY & "," & ships(i).direction & "," & ships(i).length
                End If
            End If
        Next i
        Winsock.SendData sendStr
        
        'play hit sound
        r = sndPlaySound(App.Path & "\bang.wav", SND_ASYNC)
        MsgBox ("You sunk " & user & "'s ship")
        MsgBox ("You have defeated " & user & vbNewLine & "Great Job!")
    ElseIf msg = "Wound" Then
        'animate explosion
        hitAnim.Enabled = True
        
        'play hit sound
        r = sndPlaySound(App.Path & "\bang.wav", SND_ASYNC)
    ElseIf msg = "Miss" Then
        'animate miss
        missAnim.Enabled = True
        
        'play miss sound
        r = sndPlaySound(App.Path & "\miss.wav", SND_ASYNC)
    End If
'Info is a status message
ElseIf msgtype = "Info" Then
    'Opponent is ready
    If msg = "Ready" Then
        user2 = True    'set opponent ready
        
        'if both ready start
        If user1 And user2 Then
            timer1.Enabled = False
            user1 = True
            user2 = False
            field2.Enabled = True
            changeTurn
        Else
            'keep waiting to start
            If user1 Then
                mess = " " & user & " is still not ready, please wait"
            Else
                mess = " " & user & " is ready to play!"
            End If
            lblState.Caption = mess & " ."
            timer1.Enabled = True
        End If
    End If
'Message is instant message type
ElseIf msgtype = "Message" Then
    'Display message in message window
    If txtReceive.Text <> "" Then
        txtReceive.Text = txtReceive.Text & vbNewLine & user & " -> " & msg
    Else
        txtReceive.Text = txtReceive.Text & user & " -> " & msg
    End If
    'Scroll down to end of message window
    txtReceive.SelStart = Len(txtReceive.Text)
ElseIf msgtype = "Remain" Then
    'This parses the ship data and paints it on field2
    Set tempship = New shipClass
    Do
        tempstr = Left(msg, InStr(1, msg, ":") - 1)
        msg = Mid(msg, InStr(1, msg, ":") + 1)
        tempheadx = Left(tempstr, InStr(1, tempstr, ",") - 1)
        tempstr = Mid(tempstr, InStr(1, tempstr, ",") + 1)
        tempheady = Left(tempstr, InStr(1, tempstr, ",") - 1)
        tempstr = Mid(tempstr, InStr(1, tempstr, ",") + 1)
        tempdir = Left(tempstr, InStr(1, tempstr, ",") - 1)
        tempstr = Mid(tempstr, InStr(1, tempstr, ",") + 1)
        tempsize = tempstr
        tempship.setValues tempheadx, tempheady, tempsize, tempdir
        tempship.drawShip2
    Loop While msg <> ""
    MsgBox ("You have been defeated by " & user & vbNewLine & "You Suck!")
    Unload frmMain
End If
Exit Sub
handler:
    MsgBox ("You have been defeated by " & user & vbNewLine & "You Suck!")
    Unload frmMain
End Sub

