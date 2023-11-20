VERSION 5.00
Begin VB.Form frmEditPalette 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Palette"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   Icon            =   "frmEditPalette.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   667
   Begin VB.TextBox txtEnd 
      Height          =   285
      Left            =   9360
      MaxLength       =   1
      TabIndex        =   303
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   9360
      MaxLength       =   1
      TabIndex        =   302
      Top             =   1920
      Width           =   495
   End
   Begin VB.Timer tmrColorPicker 
      Interval        =   1
      Left            =   8760
      Top             =   5640
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   299
      Top             =   5280
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   298
      Top             =   5040
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   297
      Top             =   4800
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   296
      Top             =   4560
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   295
      Top             =   4320
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   294
      Top             =   4080
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   293
      Top             =   3840
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   292
      Top             =   3600
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   291
      Top             =   3360
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   290
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   289
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   288
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   287
      Top             =   2400
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   286
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   285
      Top             =   1920
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   9000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   284
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   283
      Top             =   5280
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   282
      Top             =   5040
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   281
      Top             =   4800
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   280
      Top             =   4560
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   279
      Top             =   4320
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   278
      Top             =   4080
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   277
      Top             =   3840
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   276
      Top             =   3600
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   275
      Top             =   3360
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   274
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   273
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   272
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   271
      Top             =   2400
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   270
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   269
      Top             =   1920
      Width           =   255
   End
   Begin VB.PictureBox shpPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8760
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   268
      Top             =   1680
      Width           =   255
   End
   Begin VB.Frame fraAdvanced 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8535
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   31
         Left            =   8160
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   267
         Top             =   5040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   30
         Left            =   7200
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   266
         Top             =   5040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   29
         Left            =   8160
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   265
         Top             =   3840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   28
         Left            =   7200
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   264
         Top             =   3840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   27
         Left            =   6000
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   263
         Top             =   5040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   26
         Left            =   5040
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   262
         Top             =   5040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   25
         Left            =   6000
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   261
         Top             =   3840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   24
         Left            =   5040
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   260
         Top             =   3840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   23
         Left            =   3840
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   259
         Top             =   5040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   22
         Left            =   2880
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   258
         Top             =   5040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   21
         Left            =   3840
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   257
         Top             =   3840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   20
         Left            =   2880
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   256
         Top             =   3840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   19
         Left            =   1680
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   255
         Top             =   5040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   18
         Left            =   720
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   254
         Top             =   5040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   17
         Left            =   1680
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   253
         Top             =   3840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   16
         Left            =   720
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   252
         Top             =   3840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   15
         Left            =   8160
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   251
         Top             =   2040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   14
         Left            =   7200
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   250
         Top             =   2040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   13
         Left            =   8160
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   249
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   12
         Left            =   7200
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   248
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   11
         Left            =   6000
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   247
         Top             =   2040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   10
         Left            =   5040
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   246
         Top             =   2040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   9
         Left            =   6000
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   245
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   8
         Left            =   5040
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   244
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   7
         Left            =   3840
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   243
         Top             =   2040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   6
         Left            =   2880
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   242
         Top             =   2040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   5
         Left            =   3840
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   241
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   4
         Left            =   2880
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   240
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   0
         Left            =   720
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   239
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   3
         Left            =   1680
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   238
         Top             =   2040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   720
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   237
         Top             =   2040
         Width           =   255
      End
      Begin VB.PictureBox shpColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   1680
         ScaleHeight     =   705
         ScaleWidth      =   225
         TabIndex        =   236
         Top             =   840
         Width           =   255
      End
      Begin VB.OptionButton optPal 
         Caption         =   "03"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   35
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "02"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "01"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   33
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "00"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "00"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   31
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "01"
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   30
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "02"
         Height          =   255
         Index           =   6
         Left            =   2280
         TabIndex        =   29
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "03"
         Height          =   255
         Index           =   7
         Left            =   3240
         TabIndex        =   28
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "00"
         Height          =   255
         Index           =   8
         Left            =   4440
         TabIndex        =   27
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "01"
         Height          =   255
         Index           =   9
         Left            =   5400
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "02"
         Height          =   255
         Index           =   10
         Left            =   4440
         TabIndex        =   25
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "03"
         Height          =   255
         Index           =   11
         Left            =   5400
         TabIndex        =   24
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "00"
         Height          =   255
         Index           =   12
         Left            =   6600
         TabIndex        =   23
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "01"
         Height          =   255
         Index           =   13
         Left            =   7560
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "02"
         Height          =   255
         Index           =   14
         Left            =   6600
         TabIndex        =   21
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "03"
         Height          =   255
         Index           =   15
         Left            =   7560
         TabIndex        =   20
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "00"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "01"
         Height          =   255
         Index           =   17
         Left            =   1080
         TabIndex        =   18
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "02"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   17
         Top             =   4680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "03"
         Height          =   255
         Index           =   19
         Left            =   1080
         TabIndex        =   16
         Top             =   4680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "00"
         Height          =   255
         Index           =   20
         Left            =   2280
         TabIndex        =   15
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "01"
         Height          =   255
         Index           =   21
         Left            =   3240
         TabIndex        =   14
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "02"
         Height          =   255
         Index           =   22
         Left            =   2280
         TabIndex        =   13
         Top             =   4680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "03"
         Height          =   255
         Index           =   23
         Left            =   3240
         TabIndex        =   12
         Top             =   4680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "00"
         Height          =   255
         Index           =   24
         Left            =   4440
         TabIndex        =   11
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "01"
         Height          =   255
         Index           =   25
         Left            =   5400
         TabIndex        =   10
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "02"
         Height          =   255
         Index           =   26
         Left            =   4440
         TabIndex        =   9
         Top             =   4680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "03"
         Height          =   255
         Index           =   27
         Left            =   5400
         TabIndex        =   8
         Top             =   4680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "00"
         Height          =   255
         Index           =   28
         Left            =   6600
         TabIndex        =   7
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "01"
         Height          =   255
         Index           =   29
         Left            =   7560
         TabIndex        =   6
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "02"
         Height          =   255
         Index           =   30
         Left            =   6600
         TabIndex        =   5
         Top             =   4680
         Width           =   495
      End
      Begin VB.OptionButton optPal 
         Caption         =   "03"
         Height          =   255
         Index           =   31
         Left            =   7560
         TabIndex        =   4
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   235
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   234
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   233
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   232
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   231
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   230
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   229
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   228
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   227
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   226
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   225
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   224
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   11
         Left            =   1080
         TabIndex        =   223
         Top             =   2520
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   222
         Top             =   2520
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   9
         Left            =   1080
         TabIndex        =   221
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   220
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   7
         Left            =   1080
         TabIndex        =   219
         Top             =   2280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   218
         Top             =   2280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   5
         Left            =   1080
         TabIndex        =   217
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   216
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   215
         Top             =   2040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   214
         Top             =   2040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   213
         Top             =   840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   212
         Top             =   840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pal 00:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   600
         TabIndex        =   211
         Top             =   120
         Width           =   840
      End
      Begin VB.Line linDivide 
         Index           =   0
         X1              =   2040
         X2              =   2040
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Line linDivide 
         Index           =   8
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Line linDivide 
         Index           =   1
         X1              =   0
         X2              =   2040
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linDivide 
         Index           =   2
         X1              =   0
         X2              =   2040
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line linDivide 
         Index           =   3
         X1              =   2160
         X2              =   4200
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line linDivide 
         Index           =   4
         X1              =   2160
         X2              =   4200
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linDivide 
         Index           =   5
         X1              =   2160
         X2              =   2160
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Line linDivide 
         Index           =   6
         X1              =   4200
         X2              =   4200
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pal 01:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   2760
         TabIndex        =   210
         Top             =   120
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   14
         Left            =   2280
         TabIndex        =   209
         Top             =   840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   15
         Left            =   3240
         TabIndex        =   208
         Top             =   840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   16
         Left            =   2280
         TabIndex        =   207
         Top             =   2040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   17
         Left            =   3240
         TabIndex        =   206
         Top             =   2040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   18
         Left            =   2280
         TabIndex        =   205
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   19
         Left            =   3240
         TabIndex        =   204
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   20
         Left            =   2280
         TabIndex        =   203
         Top             =   2280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   21
         Left            =   3240
         TabIndex        =   202
         Top             =   2280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   22
         Left            =   2280
         TabIndex        =   201
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   23
         Left            =   3240
         TabIndex        =   200
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   24
         Left            =   2280
         TabIndex        =   199
         Top             =   2520
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   25
         Left            =   3240
         TabIndex        =   198
         Top             =   2520
         Width           =   150
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   197
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   196
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   195
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   7
         Left            =   3480
         TabIndex        =   194
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   193
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   192
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   191
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   7
         Left            =   3480
         TabIndex        =   190
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   189
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   188
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   187
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   7
         Left            =   3480
         TabIndex        =   186
         Top             =   2520
         Width           =   255
      End
      Begin VB.Line linDivide 
         Index           =   7
         X1              =   4320
         X2              =   6360
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line linDivide 
         Index           =   9
         X1              =   4320
         X2              =   6360
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linDivide 
         Index           =   10
         X1              =   4320
         X2              =   4320
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Line linDivide 
         Index           =   11
         X1              =   6360
         X2              =   6360
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pal 02:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   26
         Left            =   4920
         TabIndex        =   185
         Top             =   120
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   27
         Left            =   4440
         TabIndex        =   184
         Top             =   840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   28
         Left            =   5400
         TabIndex        =   183
         Top             =   840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   29
         Left            =   4440
         TabIndex        =   182
         Top             =   2040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   30
         Left            =   5400
         TabIndex        =   181
         Top             =   2040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   31
         Left            =   4440
         TabIndex        =   180
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   32
         Left            =   5400
         TabIndex        =   179
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   33
         Left            =   4440
         TabIndex        =   178
         Top             =   2280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   34
         Left            =   5400
         TabIndex        =   177
         Top             =   2280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   35
         Left            =   4440
         TabIndex        =   176
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   36
         Left            =   5400
         TabIndex        =   175
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   37
         Left            =   4440
         TabIndex        =   174
         Top             =   2520
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   38
         Left            =   5400
         TabIndex        =   173
         Top             =   2520
         Width           =   150
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   8
         Left            =   4680
         TabIndex        =   172
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   9
         Left            =   5640
         TabIndex        =   171
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   10
         Left            =   4680
         TabIndex        =   170
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   11
         Left            =   5640
         TabIndex        =   169
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   8
         Left            =   4680
         TabIndex        =   168
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   9
         Left            =   5640
         TabIndex        =   167
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   10
         Left            =   4680
         TabIndex        =   166
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   11
         Left            =   5640
         TabIndex        =   165
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   8
         Left            =   4680
         TabIndex        =   164
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   9
         Left            =   5640
         TabIndex        =   163
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   10
         Left            =   4680
         TabIndex        =   162
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   11
         Left            =   5640
         TabIndex        =   161
         Top             =   2520
         Width           =   255
      End
      Begin VB.Line linDivide 
         Index           =   12
         X1              =   6480
         X2              =   8520
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line linDivide 
         Index           =   13
         X1              =   6480
         X2              =   8520
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linDivide 
         Index           =   14
         X1              =   6480
         X2              =   6480
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Line linDivide 
         Index           =   15
         X1              =   8520
         X2              =   8520
         Y1              =   0
         Y2              =   2880
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pal 03:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   39
         Left            =   7080
         TabIndex        =   160
         Top             =   120
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   40
         Left            =   6600
         TabIndex        =   159
         Top             =   840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   41
         Left            =   7560
         TabIndex        =   158
         Top             =   840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   42
         Left            =   6600
         TabIndex        =   157
         Top             =   2040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   43
         Left            =   7560
         TabIndex        =   156
         Top             =   2040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   44
         Left            =   6600
         TabIndex        =   155
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   45
         Left            =   7560
         TabIndex        =   154
         Top             =   1080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   46
         Left            =   6600
         TabIndex        =   153
         Top             =   2280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   47
         Left            =   7560
         TabIndex        =   152
         Top             =   2280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   48
         Left            =   6600
         TabIndex        =   151
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   49
         Left            =   7560
         TabIndex        =   150
         Top             =   1320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   50
         Left            =   6600
         TabIndex        =   149
         Top             =   2520
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   51
         Left            =   7560
         TabIndex        =   148
         Top             =   2520
         Width           =   150
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   12
         Left            =   6840
         TabIndex        =   147
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   13
         Left            =   7800
         TabIndex        =   146
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   14
         Left            =   6840
         TabIndex        =   145
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   15
         Left            =   7800
         TabIndex        =   144
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   12
         Left            =   6840
         TabIndex        =   143
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   13
         Left            =   7800
         TabIndex        =   142
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   14
         Left            =   6840
         TabIndex        =   141
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   15
         Left            =   7800
         TabIndex        =   140
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   12
         Left            =   6840
         TabIndex        =   139
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   13
         Left            =   7800
         TabIndex        =   138
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   14
         Left            =   6840
         TabIndex        =   137
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   15
         Left            =   7800
         TabIndex        =   136
         Top             =   2520
         Width           =   255
      End
      Begin VB.Line linDivide 
         Index           =   16
         X1              =   0
         X2              =   2040
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line linDivide 
         Index           =   17
         X1              =   0
         X2              =   2040
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line linDivide 
         Index           =   18
         X1              =   0
         X2              =   0
         Y1              =   3000
         Y2              =   5880
      End
      Begin VB.Line linDivide 
         Index           =   19
         X1              =   2040
         X2              =   2040
         Y1              =   3000
         Y2              =   5880
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pal 04:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   52
         Left            =   600
         TabIndex        =   135
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   53
         Left            =   120
         TabIndex        =   134
         Top             =   3840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   54
         Left            =   1080
         TabIndex        =   133
         Top             =   3840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   55
         Left            =   120
         TabIndex        =   132
         Top             =   5040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   56
         Left            =   1080
         TabIndex        =   131
         Top             =   5040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   57
         Left            =   120
         TabIndex        =   130
         Top             =   4080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   58
         Left            =   1080
         TabIndex        =   129
         Top             =   4080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   59
         Left            =   120
         TabIndex        =   128
         Top             =   5280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   60
         Left            =   1080
         TabIndex        =   127
         Top             =   5280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   61
         Left            =   120
         TabIndex        =   126
         Top             =   4320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   62
         Left            =   1080
         TabIndex        =   125
         Top             =   4320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   63
         Left            =   120
         TabIndex        =   124
         Top             =   5520
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   64
         Left            =   1080
         TabIndex        =   123
         Top             =   5520
         Width           =   150
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   122
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   17
         Left            =   1320
         TabIndex        =   121
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   18
         Left            =   360
         TabIndex        =   120
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   19
         Left            =   1320
         TabIndex        =   119
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   118
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   17
         Left            =   1320
         TabIndex        =   117
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   18
         Left            =   360
         TabIndex        =   116
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   19
         Left            =   1320
         TabIndex        =   115
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   16
         Left            =   360
         TabIndex        =   114
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   17
         Left            =   1320
         TabIndex        =   113
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   18
         Left            =   360
         TabIndex        =   112
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   19
         Left            =   1320
         TabIndex        =   111
         Top             =   5520
         Width           =   255
      End
      Begin VB.Line linDivide 
         Index           =   20
         X1              =   2160
         X2              =   4200
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line linDivide 
         Index           =   21
         X1              =   2160
         X2              =   4200
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line linDivide 
         Index           =   22
         X1              =   2160
         X2              =   2160
         Y1              =   3000
         Y2              =   5880
      End
      Begin VB.Line linDivide 
         Index           =   23
         X1              =   4200
         X2              =   4200
         Y1              =   3000
         Y2              =   5880
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pal 05:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   65
         Left            =   2760
         TabIndex        =   110
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   66
         Left            =   2280
         TabIndex        =   109
         Top             =   3840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   67
         Left            =   3240
         TabIndex        =   108
         Top             =   3840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   68
         Left            =   2280
         TabIndex        =   107
         Top             =   5040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   69
         Left            =   3240
         TabIndex        =   106
         Top             =   5040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   70
         Left            =   2280
         TabIndex        =   105
         Top             =   4080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   71
         Left            =   3240
         TabIndex        =   104
         Top             =   4080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   72
         Left            =   2280
         TabIndex        =   103
         Top             =   5280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   73
         Left            =   3240
         TabIndex        =   102
         Top             =   5280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   74
         Left            =   2280
         TabIndex        =   101
         Top             =   4320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   75
         Left            =   3240
         TabIndex        =   100
         Top             =   4320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   76
         Left            =   2280
         TabIndex        =   99
         Top             =   5520
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   77
         Left            =   3240
         TabIndex        =   98
         Top             =   5520
         Width           =   150
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   20
         Left            =   2520
         TabIndex        =   97
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   21
         Left            =   3480
         TabIndex        =   96
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   22
         Left            =   2520
         TabIndex        =   95
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   23
         Left            =   3480
         TabIndex        =   94
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   20
         Left            =   2520
         TabIndex        =   93
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   21
         Left            =   3480
         TabIndex        =   92
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   22
         Left            =   2520
         TabIndex        =   91
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   23
         Left            =   3480
         TabIndex        =   90
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   20
         Left            =   2520
         TabIndex        =   89
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   21
         Left            =   3480
         TabIndex        =   88
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   22
         Left            =   2520
         TabIndex        =   87
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   23
         Left            =   3480
         TabIndex        =   86
         Top             =   5520
         Width           =   255
      End
      Begin VB.Line linDivide 
         Index           =   24
         X1              =   4320
         X2              =   6360
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line linDivide 
         Index           =   25
         X1              =   4320
         X2              =   6360
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line linDivide 
         Index           =   26
         X1              =   4320
         X2              =   4320
         Y1              =   3000
         Y2              =   5880
      End
      Begin VB.Line linDivide 
         Index           =   27
         X1              =   6360
         X2              =   6360
         Y1              =   3000
         Y2              =   5880
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pal 06:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   78
         Left            =   4920
         TabIndex        =   85
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   79
         Left            =   4440
         TabIndex        =   84
         Top             =   3840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   80
         Left            =   5400
         TabIndex        =   83
         Top             =   3840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   81
         Left            =   4440
         TabIndex        =   82
         Top             =   5040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   82
         Left            =   5400
         TabIndex        =   81
         Top             =   5040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   83
         Left            =   4440
         TabIndex        =   80
         Top             =   4080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   84
         Left            =   5400
         TabIndex        =   79
         Top             =   4080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   85
         Left            =   4440
         TabIndex        =   78
         Top             =   5280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   86
         Left            =   5400
         TabIndex        =   77
         Top             =   5280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   87
         Left            =   4440
         TabIndex        =   76
         Top             =   4320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   88
         Left            =   5400
         TabIndex        =   75
         Top             =   4320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   89
         Left            =   4440
         TabIndex        =   74
         Top             =   5520
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   90
         Left            =   5400
         TabIndex        =   73
         Top             =   5520
         Width           =   150
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   24
         Left            =   4680
         TabIndex        =   72
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   25
         Left            =   5640
         TabIndex        =   71
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   26
         Left            =   4680
         TabIndex        =   70
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   27
         Left            =   5640
         TabIndex        =   69
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   24
         Left            =   4680
         TabIndex        =   68
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   25
         Left            =   5640
         TabIndex        =   67
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   26
         Left            =   4680
         TabIndex        =   66
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   27
         Left            =   5640
         TabIndex        =   65
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   24
         Left            =   4680
         TabIndex        =   64
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   25
         Left            =   5640
         TabIndex        =   63
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   26
         Left            =   4680
         TabIndex        =   62
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   27
         Left            =   5640
         TabIndex        =   61
         Top             =   5520
         Width           =   255
      End
      Begin VB.Line linDivide 
         Index           =   28
         X1              =   6480
         X2              =   8520
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line linDivide 
         Index           =   29
         X1              =   6480
         X2              =   8520
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line linDivide 
         Index           =   30
         X1              =   6480
         X2              =   6480
         Y1              =   3000
         Y2              =   5880
      End
      Begin VB.Line linDivide 
         Index           =   31
         X1              =   8520
         X2              =   8520
         Y1              =   3000
         Y2              =   5880
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Pal 07:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   91
         Left            =   7080
         TabIndex        =   60
         Top             =   3120
         Width           =   840
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   92
         Left            =   6600
         TabIndex        =   59
         Top             =   3840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   93
         Left            =   7560
         TabIndex        =   58
         Top             =   3840
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   94
         Left            =   6600
         TabIndex        =   57
         Top             =   5040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Index           =   95
         Left            =   7560
         TabIndex        =   56
         Top             =   5040
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   96
         Left            =   6600
         TabIndex        =   55
         Top             =   4080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   97
         Left            =   7560
         TabIndex        =   54
         Top             =   4080
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   98
         Left            =   6600
         TabIndex        =   53
         Top             =   5280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Index           =   99
         Left            =   7560
         TabIndex        =   52
         Top             =   5280
         Width           =   165
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   100
         Left            =   6600
         TabIndex        =   51
         Top             =   4320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   101
         Left            =   7560
         TabIndex        =   50
         Top             =   4320
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   102
         Left            =   6600
         TabIndex        =   49
         Top             =   5520
         Width           =   150
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Index           =   103
         Left            =   7560
         TabIndex        =   48
         Top             =   5520
         Width           =   150
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   28
         Left            =   6840
         TabIndex        =   47
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   29
         Left            =   7800
         TabIndex        =   46
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   30
         Left            =   6840
         TabIndex        =   45
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblRed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   31
         Left            =   7800
         TabIndex        =   44
         Top             =   5040
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   28
         Left            =   6840
         TabIndex        =   43
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   29
         Left            =   7800
         TabIndex        =   42
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   30
         Left            =   6840
         TabIndex        =   41
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblGreen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   31
         Left            =   7800
         TabIndex        =   40
         Top             =   5280
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   28
         Left            =   6840
         TabIndex        =   39
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   29
         Left            =   7800
         TabIndex        =   38
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   30
         Left            =   6840
         TabIndex        =   37
         Top             =   5520
         Width           =   255
      End
      Begin VB.Label lblBlue 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         Height          =   255
         Index           =   31
         Left            =   7800
         TabIndex        =   36
         Top             =   5520
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdGetPal 
      Caption         =   "Get Palette..."
      Height          =   360
      Left            =   8760
      TabIndex        =   2
      Top             =   1080
      Width           =   1200
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save &As..."
      Height          =   360
      Left            =   8760
      TabIndex        =   1
      Top             =   600
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   8760
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "End:"
      Height          =   195
      Index           =   105
      Left            =   9360
      TabIndex        =   301
      Top             =   2280
      Width           =   330
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Start:"
      Height          =   195
      Index           =   104
      Left            =   9360
      TabIndex        =   300
      Top             =   1680
      Width           =   375
   End
End
Attribute VB_Name = "frmEditPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'   VB Interface Setup
'***************************************************************************
    
    Option Explicit
    Implements intResourceClient

'***************************************************************************
'   Form dimensions
'***************************************************************************

    Private Const DEF_WIDTH = 10170 'In Twips
    Private Const DEF_HEIGHT = 6510 'In Twips

'***************************************************************************
'   Editor specific variables
'***************************************************************************

    Private mbChanged As Boolean
    Private mbCached As Boolean
    Private msFilename As String
    Private mnSelectedPalette As Integer
    Private frm As New frmColorPicker
    Private colorIndex As Integer
    Private mbOpening As Boolean
    
'***************************************************************************
'   Resource Object
'***************************************************************************

    Private mGBPalette As New clsGBPalette
Public Property Let bOpening(bNewValue As Boolean)
    
    mbOpening = bNewValue
    
End Property

Public Property Get bOpening() As Boolean

    bOpening = mbOpening

End Property

Public Property Set GBPalette(oNewValue As clsGBPalette)

    Set mGBPalette = oNewValue

End Property

Public Property Get sFilename() As String

    sFilename = msFilename

End Property

Public Property Let sFilename(sNewValue As String)

    msFilename = sNewValue

End Property

Private Function mGetColor(CallingControl As PictureBox) As clsRGB

    On Error GoTo HandleErrors

'***************************************************************************
'   Get RGB values using a dialog box
'***************************************************************************

    Dim oRGB As New clsRGB
    Dim lVal As Long
    
'Display the dialog box
    lVal = CallingControl.BackColor
    mdiMain.Dialog.ShowColor
    
'If value doesn't change, then return -1 for each color as an error
    If lVal = mdiMain.Dialog.color Then
        oRGB.Red = -1
        oRGB.Green = -1
        oRGB.Blue = -1
        Set mGetColor = oRGB
        Exit Function
    End If
    
'Calculate the red, green, and blue value of the color
    Set oRGB = GetRGBFromLong(mdiMain.Dialog.color)

    oRGB.Red = oRGB.Red And &HF8
    oRGB.Green = oRGB.Green And &HF8
    oRGB.Blue = oRGB.Blue And &HF8

'Return the calculated values in a clsRGB class
    Set mGetColor = oRGB

Exit Function

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditPalette:mGetColor Error"
End Function


Public Property Get GBPalette() As clsGBPalette

    Set GBPalette = mGBPalette

End Property

Public Property Get SelectedPalette() As Integer

    SelectedPalette = mnSelectedPalette

End Property



Private Sub cmdGetPal_Click()

    On Error GoTo HandleErrors

    With mdiMain.Dialog
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DefaultExt = "bmp"
        .DialogTitle = "Load GB Palette from Windows Bitmap"
        .Filename = ""
        .Filter = "Windows Bitmaps (*.bmp)|*.bmp"
        .ShowOpen
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
        Dim i As Integer
        Dim Offscreen As New clsOffscreen
        Dim lColor As Long
        Dim oRGB As clsRGB
        
        Offscreen.CreateBitmapFromBMP .Filename
        
        For i = 1 To 32
            
            lColor = Offscreen.GetPixel(i - 1, 0)
            Set oRGB = GetRGBFromLong(lColor)
            
            mGBPalette.Colors(i).Red = (oRGB.Red \ 8) And &H1F
            mGBPalette.Colors(i).Green = (oRGB.Green \ 8) And &H1F
            mGBPalette.Colors(i).Blue = (oRGB.Blue \ 8) And &H1F
            
        Next i
        
    End With

    intResourceClient_Update

    If Not Offscreen Is Nothing Then
        Offscreen.Delete
    End If
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditPalette:cmdGetPal_Click Error"
End Sub

Private Sub cmdSave_Click()

'***************************************************************************
'   Save the current palette into a .pal file
'***************************************************************************

    On Error GoTo HandleErrors

    PackFile mGBPalette.intResource_ParentPath & "\Palettes\" & GetTruncFilename(msFilename), mGBPalette
    mbChanged = False

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditPalette:cmdSave_Click Error"
End Sub

Private Sub cmdSaveAs_Click()

'***************************************************************************
'   Save the current palette under a new filename
'***************************************************************************

    On Error GoTo HandleErrors

    With mdiMain.Dialog
        
    'Get filename
        .InitDir = Mid$(gsCurPath, 1, Len(gsCurPath) - Len(GetTruncFilename(gsCurPath)))
        .DialogTitle = "Save GB Palette"
        .Filename = ""
        .Filter = "GB Palettes (*.pal)|*.pal"
        .ShowSave
        If .Filename = "" Then
            Exit Sub
        End If
        gsCurPath = .Filename
        
    'Pack file
        PackFile .Filename, mGBPalette
        msFilename = .Filename
        
    'Reset parent path
        Dim i As Integer
        Dim flag As Boolean
        
        flag = False
        For i = Len(msFilename) To 1 Step -1
            If Mid$(msFilename, i, 1) = "\" Then
                If flag = True Then
                    mGBPalette.intResource_ParentPath = Mid$(msFilename, 1, i)
                    Exit For
                Else
                    flag = True
                End If
            End If
        Next i
        
    'Update resource cache
        gResourceCache.ReleaseClient Me
        gResourceCache.AddResourceToCache msFilename, mGBPalette, Me
        
    'Set flag used for saving
        mbChanged = False
        
        intResourceClient_Update
        
    End With

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditPalette:cmdSaveAs_Click Error"
End Sub

Private Sub Form_Activate()

    If gnPalCopy <> 1 Then
        tmrColorPicker.Enabled = True
    End If

End Sub

Private Sub Form_Deactivate()

    tmrColorPicker.Enabled = False

End Sub


Private Sub Form_Load()

    On Error GoTo HandleErrors

'***************************************************************************
'   Load the palette editor
'***************************************************************************
    
'Set form's dimensions
    Me.width = DEF_WIDTH
    Me.height = DEF_HEIGHT
    
'Organize forms on the screen
    CleanUpForms Me
    
'Set default colors and update screen

    If Not mbOpening Then
        mGBPalette.iStart = 0
        mGBPalette.iEnd = 7
    End If

    intResourceClient_Update

Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditPalette:Form_Load Error"
End Sub

Public Property Get bChanged() As Boolean

    bChanged = mbChanged

End Property

Public Property Let bChanged(bNewValue As Boolean)

    mbChanged = bNewValue

End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'***************************************************************************
'   Confirm file saving just before the form is closed
'***************************************************************************

    If mbChanged = True Then
        
    'Prompt the user
        Dim ret As Integer
        ret = MsgBox("Do you want to save " & GetTruncFilename(msFilename) & " before closing?", vbQuestion Or vbYesNoCancel, "Confirmation")
        
    'Save or cancel: whichever is appropriate
        If ret = vbYes Then
            cmdSave_Click
        ElseIf ret = vbCancel Then
            Cancel = True
        End If
        
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
'***************************************************************************
'   Release memory when closing form
'***************************************************************************
 
    On Error GoTo HandleErrors
       
    
    gResourceCache.ReleaseClient Me
   
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditPalette:Form_Unload Error"
End Sub



Private Sub fraAdvanced_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case gnTool
        Case GB_SETTER
            Me.MousePointer = vbUpArrow
    
    End Select

End Sub


Public Sub intResourceClient_Update(Optional tType As GB_UPDATETYPES)

    On Error GoTo HandleErrors

'***************************************************************************
'   Update the visual display of the palette editor
'***************************************************************************

'Display colors
    Dim i As Integer
    
    For i = 0 To 31
        shpColor(i).BackColor = RGB(mGBPalette.Colors(i + 1).Red * 8, mGBPalette.Colors(i + 1).Green * 8, mGBPalette.Colors(i + 1).Blue * 8)
        shpPreview(i).BackColor = RGB(mGBPalette.Colors(i + 1).Red * 8, mGBPalette.Colors(i + 1).Green * 8, mGBPalette.Colors(i + 1).Blue * 8)
        lblRed(i).Caption = Format$(CStr(mGBPalette.Colors(i + 1).Red), "00")
        lblGreen(i).Caption = Format$(CStr(mGBPalette.Colors(i + 1).Green), "00")
        lblBlue(i).Caption = Format$(CStr(mGBPalette.Colors(i + 1).Blue), "00")
    Next i

'Display the current palette's filename
    Me.Caption = GetTruncFilename(msFilename)
    
    txtStart.Text = CStr(mGBPalette.iStart)
    txtEnd.Text = CStr(mGBPalette.iEnd)
    
'Update clients of the current palette resource
    If tType = GB_ACTIVEEDITOR Then
        If Not mGBPalette Is Nothing Then
            mGBPalette.UpdateClients Me
            mGBPalette.UpdateClients Me
        End If
    End If
    
Exit Sub

HandleErrors:
    MsgBox Err.Description, vbCritical, "frmEditPalette:intResourceClient_Update Error"
End Sub



Private Sub optPal_Click(Index As Integer)

    Set gSelection.SrcForm = Me
    mnSelectedPalette = Index \ 4

End Sub


Private Sub shpColor_Click(Index As Integer)

    If gnWaitForPal = 1 Then
        If gnPalCopy = 1 Then
            gnPalRed = mGBPalette.Colors(Index + 1).Red
            gnPalGreen = mGBPalette.Colors(Index + 1).Green
            gnPalBlue = mGBPalette.Colors(Index + 1).Blue
            gnPalCopy = 2
            Screen.MousePointer = 0
        End If
    Else
        colorIndex = ((Index \ 4) * 4) + (Index Mod 4) + 1
        
        Load frm
        frm.nRed = mGBPalette.Colors(colorIndex).Red
        frm.nGreen = mGBPalette.Colors(colorIndex).Green
        frm.nBlue = mGBPalette.Colors(colorIndex).Blue
        
        gnWaitForPal = 1
        frm.Show
    End If

End Sub


Private Sub shpPreview_Click(Index As Integer)

    If gnWaitForPal = 1 Then
        If gnPalCopy = 1 Then
            gnPalRed = mGBPalette.Colors(Index + 1).Red
            gnPalGreen = mGBPalette.Colors(Index + 1).Green
            gnPalBlue = mGBPalette.Colors(Index + 1).Blue
            gnPalCopy = 2
            Screen.MousePointer = 0
        End If
    Else
        colorIndex = ((Index \ 4) * 4) + (Index Mod 4) + 1
        
        Load frm
        frm.nRed = mGBPalette.Colors(colorIndex).Red
        frm.nGreen = mGBPalette.Colors(colorIndex).Green
        frm.nBlue = mGBPalette.Colors(colorIndex).Blue
        
        gnWaitForPal = 1
        frm.Show
    End If

End Sub


Public Sub tmrColorPicker_Timer()

    If gnWaitForPal = 2 Then

        mGBPalette.Colors(colorIndex).Red = frm.nRed
        mGBPalette.Colors(colorIndex).Green = frm.nGreen
        mGBPalette.Colors(colorIndex).Blue = frm.nBlue
        Unload frm
    
        intResourceClient_Update
        mbChanged = True
    
        gnWaitForPal = 0
        
    End If

End Sub


Private Sub tmrUpdate_Timer()

    intResourceClient_Update

End Sub


Private Sub txtEnd_Change()

    mGBPalette.iEnd = val(txtEnd.Text)

End Sub

Private Sub txtStart_Change()

    mGBPalette.iStart = val(txtStart.Text)
    
End Sub


