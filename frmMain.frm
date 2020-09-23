VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GDI+ Load and Resize Demo"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "GDI+ Loaded and Resized PNG Picture"
      Height          =   3015
      Left            =   3120
      TabIndex        =   7
      Top             =   3960
      Width           =   5055
      Begin VB.PictureBox pctResizedPNG 
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2595
         ScaleWidth      =   4755
         TabIndex        =   8
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "GDI+ Loaded PNG Picture"
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   2895
      Begin VB.PictureBox pctPNG 
         Height          =   2685
         Left            =   120
         ScaleHeight     =   2625
         ScaleWidth      =   2625
         TabIndex        =   6
         Top             =   240
         Width           =   2685
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "GDI+ Resized from VB Picture (retaining the Ratio)"
      Height          =   3735
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.PictureBox pctResize 
         Height          =   3375
         Left            =   120
         ScaleHeight     =   3315
         ScaleWidth      =   4755
         TabIndex        =   4
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "GDI+ Loaded and Resized"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   2895
      Begin VB.PictureBox pctLoad 
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1395
         ScaleWidth      =   2595
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "VB Resized Picture"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.Image imgOriginal 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   120
         Picture         =   "frmMain.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BITMAP_FILE As String = "Picture.jpg"
Private Const PNG_FILE    As String = "Picture.png"

' When the form loads
Private Sub Form_Load()
    Dim Token    As Long

    ' Initialise GDI+
    Token = InitGDIPlus
    
    ' Load pictures
    pctLoad = LoadPictureGDIPlus(AppPath & BITMAP_FILE, pctLoad.ScaleWidth / Screen.TwipsPerPixelX, pctLoad.ScaleHeight / Screen.TwipsPerPixelY)
    pctPNG = LoadPictureGDIPlus(AppPath & PNG_FILE)
    pctResizedPNG = LoadPictureGDIPlus(AppPath & PNG_FILE, pctResizedPNG.ScaleWidth / Screen.TwipsPerPixelX, pctResizedPNG.ScaleHeight / Screen.TwipsPerPixelY)
    
    ' Resize already loaded picture (vbPicTypeIcon is not supported yet)
    pctResize = Resize(imgOriginal.Picture.Handle, imgOriginal.Picture.Type, pctResize.Width / Screen.TwipsPerPixelX, pctResize.Height / Screen.TwipsPerPixelY, vbBlack, True)
    
    ' Free GDI+
    FreeGDIPlus Token
End Sub

' Returns the application path
Private Function AppPath() As String
    AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
End Function

