VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8770CE95-D0D2-4A5F-BD93-E531C279B841}#1.7#0"; "VCBUTT~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Grocery 
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7080
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10905
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10905
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   6120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Grocery.frx":0000
      OLEDBString     =   $"Grocery.frx":008F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox lbltotUnitPrice 
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
      Left            =   240
      TabIndex        =   46
      Top             =   4320
      Width           =   1335
   End
   Begin vcButtonCTL.vcButton vcAdd 
      Height          =   315
      Left            =   480
      TabIndex        =   32
      ToolTipText     =   "Add"
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      ForeColor       =   255
      PictureLeftUp   =   "Grocery.frx":011E
      PictureRightUp  =   "Grocery.frx":0410
      PictureMiddleUp =   "Grocery.frx":0702
      PictureLeftDown =   "Grocery.frx":094C
      PictureRightDown=   "Grocery.frx":0C3E
      PictureMiddleDown=   "Grocery.frx":0EDC
      HoverPictureLeft=   "Grocery.frx":11CE
      HoverPictureRight=   "Grocery.frx":14C0
      HoverPictureMiddle=   "Grocery.frx":175E
   End
   Begin VB.CheckBox chkGrid1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4800
      TabIndex        =   30
      Top             =   5880
      Width           =   255
   End
   Begin VB.Frame frmEdit 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Edit List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3960
      Left            =   5040
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   6045
      Begin VB.CheckBox chkTT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         Caption         =   "Don't show Tooltips"
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   3840
         TabIndex        =   28
         Top             =   3600
         Width           =   15
      End
      Begin VB.TextBox txtNewCatName 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Height          =   285
         Left            =   2040
         TabIndex        =   23
         Top             =   2760
         Width           =   1710
      End
      Begin VB.TextBox txtEditItems 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   2610
         Left            =   3960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   360
         Width           =   1950
      End
      Begin VB.TextBox txtEditCat 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF00&
         Height          =   2130
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   360
         Width           =   1965
      End
      Begin VB.TextBox txtEditDesc 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   2580
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   390
         Width           =   1800
      End
      Begin vcButtonCTL.vcButton vcED 
         Height          =   315
         Left            =   120
         TabIndex        =   39
         Top             =   3120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Caption         =   "Edit Description"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLeftUp   =   "Grocery.frx":1A50
         PictureRightUp  =   "Grocery.frx":1D42
         PictureMiddleUp =   "Grocery.frx":2034
         PictureLeftDown =   "Grocery.frx":227E
         PictureRightDown=   "Grocery.frx":2570
         PictureMiddleDown=   "Grocery.frx":280E
         HoverPictureLeft=   "Grocery.frx":2B00
         HoverPictureRight=   "Grocery.frx":2D7A
         HoverPictureMiddle=   "Grocery.frx":3050
      End
      Begin vcButtonCTL.vcButton vcSD 
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Caption         =   "Save Description"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLeftUp   =   "Grocery.frx":33DE
         PictureRightUp  =   "Grocery.frx":36D0
         PictureMiddleUp =   "Grocery.frx":39C2
         PictureLeftDown =   "Grocery.frx":3C0C
         PictureRightDown=   "Grocery.frx":3EFE
         PictureMiddleDown=   "Grocery.frx":419C
         HoverPictureLeft=   "Grocery.frx":448E
         HoverPictureRight=   "Grocery.frx":4708
         HoverPictureMiddle=   "Grocery.frx":49DE
      End
      Begin vcButtonCTL.vcButton vcEC 
         Height          =   315
         Left            =   2040
         TabIndex        =   41
         Top             =   3120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Caption         =   "Edit Category"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLeftUp   =   "Grocery.frx":4D6C
         PictureRightUp  =   "Grocery.frx":505E
         PictureMiddleUp =   "Grocery.frx":5350
         PictureLeftDown =   "Grocery.frx":559A
         PictureRightDown=   "Grocery.frx":588C
         PictureMiddleDown=   "Grocery.frx":5B2A
         HoverPictureLeft=   "Grocery.frx":5E1C
         HoverPictureRight=   "Grocery.frx":6096
         HoverPictureMiddle=   "Grocery.frx":636C
      End
      Begin vcButtonCTL.vcButton vcSC 
         Height          =   315
         Left            =   2040
         TabIndex        =   42
         Top             =   3480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Caption         =   "Save Category"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLeftUp   =   "Grocery.frx":66FA
         PictureRightUp  =   "Grocery.frx":69EC
         PictureMiddleUp =   "Grocery.frx":6CDE
         PictureLeftDown =   "Grocery.frx":6F28
         PictureRightDown=   "Grocery.frx":721A
         PictureMiddleDown=   "Grocery.frx":74B8
         HoverPictureLeft=   "Grocery.frx":77AA
         HoverPictureRight=   "Grocery.frx":7A24
         HoverPictureMiddle=   "Grocery.frx":7CFA
      End
      Begin vcButtonCTL.vcButton vcSI 
         Height          =   315
         Left            =   3960
         TabIndex        =   43
         Top             =   3120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Caption         =   "Save Items"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLeftUp   =   "Grocery.frx":8088
         PictureRightUp  =   "Grocery.frx":837A
         PictureMiddleUp =   "Grocery.frx":866C
         PictureLeftDown =   "Grocery.frx":88B6
         PictureRightDown=   "Grocery.frx":8BA8
         PictureMiddleDown=   "Grocery.frx":8E46
         HoverPictureLeft=   "Grocery.frx":9138
         HoverPictureRight=   "Grocery.frx":93B2
         HoverPictureMiddle=   "Grocery.frx":9688
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Items and Prices"
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   3960
         TabIndex        =   27
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Categories List"
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Left            =   1920
         TabIndex        =   26
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Description List"
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   25
         Top             =   180
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NEW CATEGORY"
         ForeColor       =   &H00FFFF00&
         Height          =   210
         Left            =   2160
         TabIndex        =   24
         Top             =   2520
         Width           =   1515
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Show Tooltips"
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   4080
         TabIndex        =   29
         Top             =   3120
         Width           =   1485
      End
   End
   Begin MSComctlLib.ListView ltbItem 
      Height          =   4440
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   7832
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   65280
      BackColor       =   -2147483641
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picHiddenData 
      Height          =   435
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   17
      Top             =   5640
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtQuan 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   600
      TabIndex        =   3
      Text            =   "1"
      Top             =   5400
      Width           =   360
   End
   Begin VB.ComboBox cboCat 
      BackColor       =   &H00F6F0E0&
      Height          =   315
      ItemData        =   "Grocery.frx":9A16
      Left            =   120
      List            =   "Grocery.frx":9A18
      TabIndex        =   2
      Top             =   3150
      Width           =   1680
   End
   Begin VB.ComboBox cboUnits 
      BackColor       =   &H00F6F0E0&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3960
      Left            =   5280
      TabIndex        =   5
      Top             =   1680
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   6985
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MouseIcon       =   "Grocery.frx":9A1A
      NumItems        =   0
   End
   Begin Grocery_bhernz.DynamicPopupMenu DPM1 
      Left            =   3000
      Top             =   5640
      _ExtentX        =   529
      _ExtentY        =   476
   End
   Begin VB.PictureBox picGreenBar 
      Height          =   390
      Left            =   2400
      ScaleHeight     =   330
      ScaleWidth      =   555
      TabIndex        =   16
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chkGrid 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Show Grid"
      Height          =   225
      Left            =   10920
      TabIndex        =   0
      Top             =   6960
      Width           =   225
   End
   Begin vcButtonCTL.vcButton vcEdit 
      Height          =   315
      Left            =   3360
      TabIndex        =   33
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Edit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      ForeColor       =   255
      PictureLeftUp   =   "Grocery.frx":FCB4
      PictureRightUp  =   "Grocery.frx":FFA6
      PictureMiddleUp =   "Grocery.frx":10298
      PictureLeftDown =   "Grocery.frx":104E2
      PictureRightDown=   "Grocery.frx":107D4
      PictureMiddleDown=   "Grocery.frx":10A72
      HoverPictureLeft=   "Grocery.frx":10D64
      HoverPictureRight=   "Grocery.frx":10FDE
      HoverPictureMiddle=   "Grocery.frx":112B4
   End
   Begin vcButtonCTL.vcButton vcDel 
      Height          =   315
      Left            =   1920
      TabIndex        =   34
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      ForeColor       =   255
      PictureLeftUp   =   "Grocery.frx":11642
      PictureRightUp  =   "Grocery.frx":11934
      PictureMiddleUp =   "Grocery.frx":11C26
      PictureLeftDown =   "Grocery.frx":11E70
      PictureRightDown=   "Grocery.frx":12162
      PictureMiddleDown=   "Grocery.frx":12400
      HoverPictureLeft=   "Grocery.frx":126F2
      HoverPictureRight=   "Grocery.frx":1296C
      HoverPictureMiddle=   "Grocery.frx":12C42
   End
   Begin vcButtonCTL.vcButton vcButton2 
      Height          =   315
      Left            =   9120
      TabIndex        =   35
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      ForeColor       =   255
      PictureLeftUp   =   "Grocery.frx":12FD0
      PictureRightUp  =   "Grocery.frx":132C2
      PictureMiddleUp =   "Grocery.frx":135B4
      PictureLeftDown =   "Grocery.frx":137FE
      PictureRightDown=   "Grocery.frx":13AF0
      PictureMiddleDown=   "Grocery.frx":13D8E
      HoverPictureLeft=   "Grocery.frx":14080
      HoverPictureRight=   "Grocery.frx":142FA
      HoverPictureMiddle=   "Grocery.frx":145D0
   End
   Begin vcButtonCTL.vcButton vcClr 
      Height          =   315
      Left            =   4800
      TabIndex        =   36
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      ForeColor       =   255
      PictureLeftUp   =   "Grocery.frx":1495E
      PictureRightUp  =   "Grocery.frx":14C50
      PictureMiddleUp =   "Grocery.frx":14F42
      PictureLeftDown =   "Grocery.frx":1518C
      PictureRightDown=   "Grocery.frx":1547E
      PictureMiddleDown=   "Grocery.frx":1571C
      HoverPictureLeft=   "Grocery.frx":15A0E
      HoverPictureRight=   "Grocery.frx":15C88
      HoverPictureMiddle=   "Grocery.frx":15F5E
   End
   Begin vcButtonCTL.vcButton vcPrint 
      Height          =   315
      Left            =   7680
      TabIndex        =   37
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      ForeColor       =   255
      PictureLeftUp   =   "Grocery.frx":162EC
      PictureRightUp  =   "Grocery.frx":165DE
      PictureMiddleUp =   "Grocery.frx":168D0
      PictureLeftDown =   "Grocery.frx":16B1A
      PictureRightDown=   "Grocery.frx":16E0C
      PictureMiddleDown=   "Grocery.frx":170AA
      HoverPictureLeft=   "Grocery.frx":1739C
      HoverPictureRight=   "Grocery.frx":17616
      HoverPictureMiddle=   "Grocery.frx":178EC
   End
   Begin vcButtonCTL.vcButton vcClear 
      Height          =   315
      Left            =   240
      TabIndex        =   38
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureLeftUp   =   "Grocery.frx":17C7A
      PictureRightUp  =   "Grocery.frx":17F6C
      PictureMiddleUp =   "Grocery.frx":1825E
      PictureLeftDown =   "Grocery.frx":184A8
      PictureRightDown=   "Grocery.frx":1879A
      PictureMiddleDown=   "Grocery.frx":18A38
      HoverPictureLeft=   "Grocery.frx":18D2A
      HoverPictureRight=   "Grocery.frx":18FA4
      HoverPictureMiddle=   "Grocery.frx":1927A
   End
   Begin vcButtonCTL.vcButton vcAbout 
      Height          =   315
      Left            =   360
      TabIndex        =   44
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "About Me"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      PictureLeftUp   =   "Grocery.frx":19608
      PictureRightUp  =   "Grocery.frx":198FA
      PictureMiddleUp =   "Grocery.frx":19BEC
      PictureLeftDown =   "Grocery.frx":19E36
      PictureRightDown=   "Grocery.frx":1A128
      PictureMiddleDown=   "Grocery.frx":1A3C6
      HoverPictureLeft=   "Grocery.frx":1A6B8
      HoverPictureRight=   "Grocery.frx":1A932
      HoverPictureMiddle=   "Grocery.frx":1AC08
   End
   Begin vcButtonCTL.vcButton vcCalcu 
      Height          =   315
      Left            =   6240
      TabIndex        =   45
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "Calculator"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
      PictureLeftUp   =   "Grocery.frx":1AF96
      PictureRightUp  =   "Grocery.frx":1B288
      PictureMiddleUp =   "Grocery.frx":1B57A
      PictureLeftDown =   "Grocery.frx":1B7C4
      PictureRightDown=   "Grocery.frx":1BAB6
      PictureMiddleDown=   "Grocery.frx":1BD54
      HoverPictureLeft=   "Grocery.frx":1C046
      HoverPictureRight=   "Grocery.frx":1C2C0
      HoverPictureMiddle=   "Grocery.frx":1C596
   End
   Begin VB.Image Image8 
      Height          =   375
      Left            =   6360
      Picture         =   "Grocery.frx":1C924
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   2280
      Picture         =   "Grocery.frx":1D1EE
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   600
      Left            =   600
      Picture         =   "Grocery.frx":1DAB8
      Stretch         =   -1  'True
      Top             =   240
      Width           =   825
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   375
      Left            =   9600
      Top             =   5760
      Width           =   975
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   375
      Left            =   240
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      Height          =   375
      Left            =   5280
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   1920
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   240
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GROCERY LIST COMPUTATION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   31
      Top             =   240
      Width           =   7695
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   8040
      TabIndex        =   15
      Top             =   840
      Width           =   2865
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "Grocery.frx":1E382
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   11055
   End
   Begin VB.Label lblEditCatName 
      Height          =   240
      Left            =   0
      TabIndex        =   18
      Top             =   6840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM DESCRIPTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   1875
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "UNITS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UNIT PRICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   4020
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   11
      Top             =   4920
      Width           =   1395
   End
   Begin VB.Image ImageDown 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   960
      Picture         =   "Grocery.frx":15EB84
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image ImageUp 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   960
      Picture         =   "Grocery.frx":15ECCE
      Top             =   5400
      Width           =   300
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM NAMES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   2880
      TabIndex        =   10
      Top             =   1320
      Width           =   1620
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   8985
      TabIndex        =   8
      Top             =   5865
      Width           =   585
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "GROCERY LIST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   6960
      TabIndex        =   7
      Top             =   1320
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   9600
      TabIndex        =   6
      Top             =   5730
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   0
      Picture         =   "Grocery.frx":15EE18
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11055
   End
   Begin VB.Image Image5 
      Height          =   5415
      Left            =   4920
      Picture         =   "Grocery.frx":29F61A
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   8010
   End
   Begin VB.Image Image4 
      Height          =   5655
      Left            =   0
      Picture         =   "Grocery.frx":2DBB46
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   -120
      Picture         =   "Grocery.frx":318072
      Stretch         =   -1  'True
      Top             =   960
      Width           =   11055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   5565
      TabIndex        =   9
      Top             =   5880
      Width           =   945
   End
End
Attribute VB_Name = "Grocery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
 Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
   
   Dim TotPrice As Single
   Dim Price As String

Private Sub cboUnits_Change()
   If ListView1.ListItems.Count <> 0 Then ListView1.SelectedItem.Selected = False
End Sub
Private Sub chkGrid1_Click()
   If chkGrid1.Value = Checked Then
      ltbItem.GridLines = True
   Else
      ltbItem.GridLines = False
   End If

End Sub
Private Sub Form_Load()
Dim Region As Long
Dim strPath As String
Dim strMapName As String
Dim fStg As String
Dim fSLen As Integer
Dim firststg As String

Load welcome_bhernz
welcome_bhernz.Show
Label16.Caption = Format(Now, "Long Date")
   ListViewSetup
   LoadcboUnits
   LoadcboCat
   ListAll
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Seld As String

If Button = 1 Then               'move form
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

If Button = 2 Then               'show popup menu
   Seld = DPM1.popup("Delete Item,Clear List,Edit List,Print List,-,Exit")
   Select Case Seld
      Case "Delete Item":
          vcDel_Click
      Case "Clear List":
         vcClr_Click
      Case "Edit List":
          vcEdit_Click
          If vcEdit.Caption = "Edit (Hide)" Then
             vcEdit.Caption = "Edit (Show)"
         Else
             vcEdit.Caption = "Edit (Hide)"
         End If
      Case "Print List":
          vcPrint_Click
      Case "Exit":
         vcButton2_Click
   End Select
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub chkGrid_Click()                             'show / hide grid
   If chkGrid.Value = Checked Then
      ListView1.GridLines = True
   Else
      ListView1.GridLines = False
   End If
End Sub



Private Sub txtQuan_GotFocus()
   txtQuan.SelStart = 2                                                'set carot to right of charactor
   If ListView1.ListItems.Count <> 0 Then ListView1.SelectedItem.Selected = False     'hide the selection bar (highlight)
End Sub

Private Sub txtQuan_KeyPress(KeyAscii As Integer)
    ' accept numbers and backspace only
    If InStr(1, "8 48 49 50 51 52 53 54 55 56 57", CStr(KeyAscii)) = 0 Then
         KeyAscii = 0
    End If
End Sub
Private Sub cboCat_Click()
   If ListView1.ListItems.Count <> 0 Then ListView1.SelectedItem.Selected = False   'hide the selection bar (highlight)
   If cboCat.Text = "" Or cboCat.Text = "All" Then
      ListAll
   Else
      ltbItem.ListItems.Clear
      ltbItem.Sorted = False    'turn sorted off and load items into list
      LoadCombo cboCat.Text, cboCat
      ltbItem.Sorted = True     'with items loaded, we can now sort the items.Not all prices show if you don't do this
   End If
End Sub


Private Sub ListAll()   'loads items from all catogories into one list
   Dim xxx As Integer
   'load all items in lstItemName
   ltbItem.ListItems.Clear
   ltbItem.Sorted = False
   For xxx = 1 To cboCat.ListCount - 1
      LoadCombo cboCat.List(xxx), cboCat
   Next xxx
   ltbItem.Sorted = True
   cboCat.Text = "All"
End Sub

Private Sub LoadCombo(textfile As String, cboname As ComboBox)  'loads items into listbox
   Dim strArray() As String
   Dim i As Integer
   Dim iFile As Integer
   Dim y As Integer
   
      iFile = FreeFile
      If textfile = "" Or textfile = "All" Then Exit Sub
      Open App.Path & "\ListData\" & textfile & ".txt" For Input As #iFile
      Do While Not EOF(iFile)
         Line Input #iFile, textfile
         strArray = Split(textfile, ",")
         y = ltbItem.ListItems.Count + 1
       For i = 0 To UBound(strArray) Step 2
         If Not Trim$(strArray(i)) = "" Then
            ltbItem.ListItems.Add (y), , strArray(i)
            ltbItem.ListItems(y).SubItems(1) = strArray(i + 1)
         End If
      Next i
         Loop
         Close #iFile
End Sub

Private Sub LoadcboUnits()   ' load the descriptors
   Dim textfile As String
   Dim strArrayUnit() As String
   Dim i As Integer
   Dim iFile As Integer
   
   cboUnits.Clear
   iFile = FreeFile
   Open App.Path & "\ListData\UnitDes.udr" For Input As #iFile
   Do While Not EOF(iFile)
      Line Input #iFile, textfile
      strArrayUnit = Split(textfile, ",")
      For i = 0 To UBound(strArrayUnit)
         If Not Trim$(strArrayUnit(i)) = "" Then
            cboUnits.AddItem strArrayUnit(i)
         End If
      Next i
      Loop
      Close #iFile
End Sub

Private Sub LoadcboCat()
   Dim strPath As String
   Dim fStg As String
   Dim fSLen As Integer
   'Load categories into combo box
   cboCat.AddItem "All"
   strPath = Dir$(App.Path & "\ListData\" & "*.txt")
   If Not strPath = "" Then                                    'yes, there are files here so
   Do                                                                  'go get them
      fSLen = Len(strPath) - 4                                'filename length minus extension
      fStg = Mid$(strPath, 1, fSLen)                        'filename without extension
      cboCat.AddItem fStg                                     'put filename into combobox
      strPath = Dir$
   Loop Until strPath = ""
Else
   MsgBox "No files found!", vbCritical + vbOKOnly, "File - Error"
End If
End Sub

Private Sub ImageUp_Click()
     txtQuan.Text = Val(txtQuan.Text) + 1
     If Val(txtQuan.Text) > 25 Then txtQuan.Text = "25"
     txtQuan.SelStart = 1                                     'keep carot of the right of number
End Sub

Private Sub ImageDown_Click()
    txtQuan.Text = Val(txtQuan.Text) - 1
    If Val(txtQuan.Text) < 1 Then txtQuan.Text = "1"
    txtQuan.SelStart = 1                                      'keep carot of the right of number
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim Seld As String
   
   If Button = 2 Then                                           'show popup menu
   Seld = DPM1.popup("Delete Item,Clear List,Edit List,Print List,-,Exit")
   Select Case Seld
      Case "Delete Item":
          vcDel_Click
      Case "Clear List":
          vcClr_Click
      Case "Edit List":
          vcEdit_Click
          If vcEdit.Caption = "Edit (Hide)" Then
             vcEdit.Caption = "Edit (Show)"
         Else
             vcEdit.Caption = "Edit (Hide)"
         End If
      Case "Print List":
          vcPrint_Click
      Case "Exit":
         vcButton2_Click
   End Select
End If
End Sub

Private Sub ltbItem_Click()
    If txtQuan.Text = "" Then txtQuan.Text = "1"
    lbltotUnitPrice.Text = Val(txtQuan.Text) * ltbItem.SelectedItem.SubItems(1)
    lbltotUnitPrice.Text = Format(lbltotUnitPrice, "##0.00")
End Sub

Private Sub ltbItem_DblClick()
   vcAdd_Click
End Sub

Private Sub txtEditCat_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If chkTT.Value = 1 Then MsgBox "Click on a category to show current items.", "Categorys"
End Sub

Private Sub txtEditDesc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If chkTT.Value = 1 Then MsgBox "Add or Delete descriptors and Save." & vbCrLf & "Don't forget the comma.", "Edit Descriptors"
End Sub

Private Sub txtEditItems_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If chkTT.Value = 1 Then MsgBox "Add or Delete items and price and press Save." & vbCrLf & "Format: Item,Price, (ex.item,0.00,)" & vbCrLf & "Don't forget the commas", "Edit Items and Price"
End Sub

Private Sub txtNewCatName_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If chkTT.Value = 1 Then MsgBox "Add a New Category name here and press Save." & vbCrLf & "To delete, Select a category above and press Delete.", "Edit Category"
End Sub




Private Sub txtEditCat_Click()
'highlight category that was clicked on
SendKeys "{HOME}+{END}"
DoEvents
lblEditCatName.Caption = txtEditCat.SelText
EditLoadItems                                                           'load categories into textbox
End Sub

Private Sub ListViewSetup()
    Dim iFontHeight As Long
    Dim iBarHeight As Integer
  
    'Set up a few listview properties, these can be set on the property page instead of being listed here.
    ListView1.View = lvwReport
    ListView1.ColumnHeaders.Add 1, , "Qty"
    ListView1.ColumnHeaders(1).Width = 530
    ListView1.ColumnHeaders.Add 2, , "Descriptor"
    ListView1.ColumnHeaders(2).Width = 1080
    ListView1.ColumnHeaders.Add 3, , "Item"
    ListView1.ColumnHeaders(3).Width = 1765
    ListView1.ColumnHeaders.Add 4, , "Unit Price"
    ListView1.ColumnHeaders(4).Width = 960
    ListView1.ColumnHeaders.Add 5, , "Total Price"
    ListView1.ColumnHeaders(5).Width = 965
    ListView1.Width = ListView1.ColumnHeaders(1).Width + ListView1.ColumnHeaders(2).Width + ListView1.ColumnHeaders(3).Width + ListView1.ColumnHeaders(4).Width + ListView1.ColumnHeaders(5).Width
    ltbItem.ColumnHeaders.Add 1, , "Item"
    ltbItem.ColumnHeaders(1).Width = 2045
    ltbItem.ColumnHeaders.Add 2, , "Price"
    ltbItem.ColumnHeaders(2).Width = 1050
    ltbItem.Width = ltbItem.ColumnHeaders(1).Width + ltbItem.ColumnHeaders(2).Width + 20
    Me.ScaleMode = vbTwips 'make sure our form is In twips
    'Paints the green and white bars
    picGreenBar.ScaleMode = vbTwips
    picGreenBar.BorderStyle = vbBSNone 'this is important - we don't want To measure the border In our calcs.
    picGreenBar.AutoRedraw = True
    picGreenBar.Visible = False
    picGreenBar.Font = ListView1.Font
    picGreenBar.FontSize = ListView1.Font.Size
    iFontHeight = picGreenBar.TextHeight("b") + Screen.TwipsPerPixelY
    iBarHeight = (iFontHeight)
    picGreenBar.Width = ListView1.Width
    picGreenBar.Height = iBarHeight * 2
    picGreenBar.ScaleMode = vbUser
    picGreenBar.ScaleHeight = 2
    picGreenBar.ScaleWidth = 1
    picGreenBar.Line (0, 0)-(1, 1), vbWhite, BF
    picGreenBar.Line (0, 1)-(1, 2), &HF6F0E0, BF
   
   ListView1.PictureAlignment = lvwTile
    ListView1.Picture = picGreenBar.Image
   
   
   
    
End Sub
Private Sub EditLoadCat()
   Dim strPath As String
   Dim fStg As String
   Dim fSLen As Integer
   
   txtEditCat.Text = ""
   strPath = Dir(App.Path & "\ListData\" & "*.txt")
   If Not strPath = "" Then
      Do
         fSLen = Len(strPath) - 4
         fStg = Mid$(strPath, 1, fSLen)
         txtEditCat.Text = txtEditCat.Text & fStg & vbCrLf
    
         strPath = Dir$
      Loop Until strPath = ""
   Else
      MsgBox "No files found!", vbCritical + vbOKOnly, "File - Error"
   End If
End Sub

Private Sub EditLoadItems()
   If lblEditCatName.Caption = "" Then
      MsgBox "Select a category to load", vbOKOnly, "No Category selected"
      txtEditItems.Text = ""
      Exit Sub
   End If
    LoadText (App.Path & "\ListData\" & lblEditCatName.Caption & ".txt"), txtEditItems
End Sub
Private Sub List_Remove(TheList As ListBox)
   On Error Resume Next
   If TheList.ListCount < 0 Then Exit Sub
   TheList.RemoveItem TheList.ListIndex
End Sub

Public Function SaveText(strText As String, FileName As String) As Boolean
Dim iFile As Integer

On Error GoTo handle
    iFile = FreeFile
    Open FileName For Output As #iFile          'Opening the file to SaveText
        Print #iFile, strText                               'Printing  the text to the file
    Close #iFile                                              'Closing
    If FileExists(FileName) = False Then          'Check whether the file created
        MsgBox "Unexpectd error occured. File could not be saved", vbCritical, "Sorry"
        SaveText = False                                  'Returns 'False'
    Else
        SaveText = True                                    'Returns 'True'
    End If
Exit Function
handle:
    SaveText = False
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Function

Private Sub LoadText(textfile As String, txtname As TextBox)
   Dim iFile As Integer
   
   txtname.Text = ""
    iFile = FreeFile
    Open textfile For Input As #iFile
        textfile = Input(LOF(iFile), iFile)
        txtname.Text = textfile
    Close #iFile
End Sub
Private Sub vcAbout_Click()
Splash.Show
Unload Me
End Sub
Private Sub vcAdd_Click()
   Dim x As Integer
    If lbltotUnitPrice.Text = "" Then Exit Sub

   If ltbItem.SelectedItem.Text = "" Then
      txtQuan.Text = "1"
      cboUnits.Text = ""
      Exit Sub
   End If

    If txtQuan.Text = "" Then txtQuan.Text = "1"
    txtQuan.Text = Format(txtQuan.Text, "##0")
    x = ListView1.ListItems.Count + 1
    ListView1.ListItems.Add x, , txtQuan.Text
    ListView1.ListItems(x).SubItems(1) = cboUnits.Text
    ListView1.ListItems(x).SubItems(2) = ltbItem.SelectedItem.Text
    ListView1.ListItems(x).SubItems(3) = ltbItem.SelectedItem.SubItems(1)
    Price = Val(txtQuan.Text) * ltbItem.SelectedItem.SubItems(1)
    ListView1.ListItems(x).SubItems(4) = Format(Price, "P0.00")
   
   TotPrice = TotPrice + Val(Price)
   Label1.Caption = Format(TotPrice, "P0.00")
   

   ltbItem.SelectedItem.Selected = False
   txtQuan.Text = "1"
   txtQuan.SetFocus
   lbltotUnitPrice.Text = ""
   ListView1.SelectedItem.Selected = False
   PcSpeakerBeep 400, 50
   PcSpeakerBeep 600, 50

End Sub

Private Sub vcButton2_Click()
Unload Me

End Sub

Private Sub vcCalcu_Click()
On Error GoTo errHandle
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\calc.exe", vbNormalFocus)
    Exit Sub
errHandle:
    MsgBox "Unable to run Calculator Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next

End Sub

Private Sub vcClear_Click()
   txtQuan.Text = "1"
   cboUnits.Text = ""
   lbltotUnitPrice.Caption = ""
   ltbItem.SelectedItem.Selected = False
   PcSpeakerBeep 500, 50

End Sub

Private Sub vcClr_Click()
    ListView1.ListItems.Clear
    Label1.Caption = "P0.00"
    TotPrice = 0
   PcSpeakerBeep 400, 50
   PcSpeakerBeep 600, 50

End Sub

Private Sub vcDel_Click()
   On Error Resume Next
   If ListView1.SelectedItem.Selected = False Then Exit Sub
   If txtQuan.Text = "" Then txtQuan.Text = "1"
   If ListView1.ListItems.Count = 0 Then Exit Sub
   Price = Val(txtQuan.Text) * ListView1.SelectedItem.SubItems(4)
   ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
   TotPrice = TotPrice - Val(Price)
   Label1.Caption = Format(TotPrice, "P0.00")
   PcSpeakerBeep 600, 50
   PcSpeakerBeep 400, 50

End Sub

Private Sub vcEC_Click()
    DeleteFile App.Path & "\ListData\" & txtEditCat.SelText & ".txt"
    txtEditItems.Text = ""
    txtNewCatName.Text = ""
    EditLoadCat
    cboCat.Clear
    LoadcboCat


End Sub

Private Sub vcED_Click()
    LoadText (App.Path & "\ListData\UnitDes.udr"), txtEditDesc

End Sub

Private Sub vcEdit_Click()
On Error Resume Next
     frmEdit.Visible = Not frmEdit.Visible
     txtEditCat.Text = ""
     txtEditDesc.Text = ""
     txtEditItems.Text = ""
     txtEditItems.Text = "Click on a category from the list on the left to show current items."
     cboCat_Click
     txtQuan.SetFocus
     EditLoadCat
     vcED_Click
     PcSpeakerBeep 400, 50
     PcSpeakerBeep 600, 50
   


End Sub

Private Sub vcPrint_Click()
   Dim subit As String
   Dim i As Integer
    
    If ListView1.ListItems.Count = 0 Then Exit Sub
   On Error Resume Next
   Printer.ScaleMode = 3
   Printer.FontSize = 12
   Printer.Print Tab(25); "Grocery List"
   Printer.Print
   Printer.FontUnderline = True
   Printer.Print "Qty  Descriptor         Item                                                Price     Total Price"
   Printer.FontUnderline = False
    For i = 1 To ListView1.ListItems.Count
       subit = ListView1.ListItems(i).Text
       Printer.Print subit; Tab(6); ListView1.ListItems(i).SubItems(1); Tab(22); ListView1.ListItems(i).SubItems(2); Tab(59); ListView1.ListItems(i).SubItems(3); Tab(70); ListView1.ListItems(i).SubItems(4)
       subit = ""
    Next i
     Printer.Print
     Printer.Print Tab(62); "Total:  " & Label1.Caption
     Printer.NewPage
     Printer.EndDoc
     PcSpeakerBeep 400, 50
     PcSpeakerBeep 600, 50
   End Sub

Private Sub vcSC_Click()
   If txtNewCatName.Text = "" Then
      MsgBox "Nothing to Save"
      Exit Sub
   End If
   SaveText "", App.Path & "\ListData\" & txtNewCatName.Text & ".txt"
   txtNewCatName.Text = ""
   txtEditCat.Text = ""
   cboCat.Clear
   EditLoadCat
   LoadcboCat


End Sub

Private Sub vcSD_Click()
   SaveText txtEditDesc.Text, App.Path & "\ListData\UnitDes.udr"
   cboUnits.Clear
   LoadcboUnits
   txtEditDesc.Text = ""

End Sub

Private Sub vcSI_Click()
      If lblEditCatName.Caption = "" Then
         MsgBox "No category to save to"
         txtEditItems.Text = ""
         Exit Sub
      End If
      If MsgBox("Are you sure?", vbYesNo, "Save") = vbNo Then Exit Sub
    SaveText txtEditItems.Text, App.Path & "\ListData\" & lblEditCatName.Caption & ".txt"
    cboCat.Text = lblEditCatName.Caption
    cboCat_Click


End Sub
