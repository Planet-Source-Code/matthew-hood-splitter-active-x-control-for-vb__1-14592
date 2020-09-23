VERSION 5.00
Object = "{D362A7ED-D769-43C2-8F58-E0BF62E9AAF5}#1.0#0"; "VBSPLITTER.OCX"
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vbSplitter Demo Application"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin vbSplitterControl.vbSplitter vbSplitter1 
      Height          =   5055
      Left            =   3480
      TabIndex        =   24
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8916
      BorderStyle     =   1
      MaxSize         =   0
      MinSize         =   0
      AutoResize      =   -1  'True
      Orientation     =   1
      SelectedColor   =   -2147483628
      SplitterColor   =   -2147483633
      SplitterWidth   =   60
      Begin vbSplitterControl.vbSplitter vbSplitter2 
         Height          =   2415
         Left            =   60
         TabIndex        =   26
         Top             =   2520
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   4260
         BorderStyle     =   1
         MaxSize         =   0
         MinSize         =   0
         AutoResize      =   -1  'True
         Orientation     =   1
         SelectedColor   =   -2147483628
         SplitterColor   =   -2147483633
         SplitterWidth   =   60
         Begin VB.CommandButton cmdDummy 
            Caption         =   "Dummy Button"
            Height          =   615
            Index           =   0
            Left            =   180
            TabIndex        =   29
            Top             =   1560
            Width           =   2295
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FF0000&
            Height          =   915
            Left            =   60
            ScaleHeight     =   855
            ScaleWidth      =   2595
            TabIndex        =   27
            Top             =   420
            Width           =   2655
            Begin VB.CommandButton cmdDummy 
               Caption         =   "Dummy Button"
               Height          =   555
               Index           =   1
               Left            =   240
               TabIndex        =   28
               Top             =   180
               Width           =   1815
            End
         End
      End
      Begin VB.PictureBox picPanel1 
         BackColor       =   &H000000FF&
         Height          =   2115
         Left            =   180
         ScaleHeight     =   2055
         ScaleWidth      =   2655
         TabIndex        =   25
         Top             =   60
         Width           =   2715
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Size Properties"
      Height          =   2535
      Left            =   60
      TabIndex        =   1
      Top             =   2640
      Width           =   3315
      Begin VB.TextBox txtNewSplitterWidth 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1260
         TabIndex        =   21
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txtSplitterWidth 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txtPanelSize 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtMaxSize 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtMinSize 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtNewPanelSize 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1260
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtNewMaxSize 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1260
         TabIndex        =   15
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtNewMinSize 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1260
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "ForceResize"
         Height          =   255
         Index           =   1
         Left            =   1260
         TabIndex        =   13
         Top             =   2100
         Width           =   1935
      End
      Begin VB.Label Labels 
         Caption         =   "SplitterWidth:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Labels 
         Caption         =   "PanelSize:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Labels 
         Caption         =   "MaxSize:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2220
         X2              =   2220
         Y1              =   420
         Y2              =   1980
      End
      Begin VB.Label Labels 
         Caption         =   "MinSize:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   3315
      Begin VB.CheckBox chkAutoSize 
         Alignment       =   1  'Right Justify
         Caption         =   "Auto Resize:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.ComboBox cboSelectedColor 
         Height          =   315
         ItemData        =   "frmDemo.frx":0E42
         Left            =   1260
         List            =   "frmDemo.frx":0FBF
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1620
         Width           =   1935
      End
      Begin VB.ComboBox cboSplitterColor 
         Height          =   315
         ItemData        =   "frmDemo.frx":118B
         Left            =   1260
         List            =   "frmDemo.frx":1308
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cboBorderStyle 
         Height          =   315
         ItemData        =   "frmDemo.frx":14D4
         Left            =   1260
         List            =   "frmDemo.frx":14DE
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cboOrientation 
         Height          =   315
         ItemData        =   "frmDemo.frx":1503
         Left            =   1260
         List            =   "frmDemo.frx":150D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   1935
      End
      Begin VB.Label Labels 
         Caption         =   "Selected Color:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Labels 
         Caption         =   "Splitter Color:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Labels 
         Caption         =   "Border Style:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Labels 
         Caption         =   "Orientation:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************************************
'vbSplitter Control Demo
'Copyright: ©2001 Dragon Weyr Development.
'Author(s): Matthew Hood Email: DragonWeyrDev@Yahoo.com
'***********************************************************************************************
'Revision History:
'[Matthew Hood]
'   01/18/01 - New
'***********************************************************************************************
'***********************************************************************************************
'Private Variables/Constants
'***********************************************************************************************

'***********************************************************************************************
'API Declarations
'***********************************************************************************************
'***********************************************************************************************
'Public Properties/Constants
'***********************************************************************************************
'***********************************************************************************************
'Public Methods
'***********************************************************************************************
'***********************************************************************************************
'Private Methods
'***********************************************************************************************
'***********************************************************************************************
'Load/Unload Events
'***********************************************************************************************
Private Sub Form_Load()

    'Set the splitter panel objects.
    Set vbSplitter1.Child1 = picPanel1
    Set vbSplitter1.Child2 = vbSplitter2
    
    Set vbSplitter2.Child1 = cmdDummy(0)
    Set vbSplitter2.Child2 = Picture1
    
    'Set the second splitter borderstyle.
    vbSplitter2.BorderStyle = vbStyleNone

    'Setup display value defaults.
    cboBorderStyle.ListIndex = 1
    cboOrientation.ListIndex = 1
    cboSplitterColor.ListIndex = 9
    cboSelectedColor.ListIndex = 2
    
    'Get current size properties.
    txtSplitterWidth.Text = vbSplitter1.SplitterWidth
    txtMinSize.Text = vbSplitter1.MinSize
    txtMaxSize.Text = vbSplitter1.MaxSize
    txtPanelSize.Text = vbSplitter1.PanelSize
End Sub
'***********************************************************************************************
'Resize Events
'***********************************************************************************************
Private Sub Picture1_Resize()
On Error Resume Next
    cmdDummy(1).Move 120, 120, Picture1.ScaleWidth - 240, Picture1.ScaleHeight - 240
End Sub
'***********************************************************************************************
'Focus Events
'***********************************************************************************************
'***********************************************************************************************
'Click Events
'***********************************************************************************************
'Change the splitter control border style.
Private Sub cboBorderStyle_Click()
    vbSplitter1.BorderStyle = cboBorderStyle.ItemData(cboBorderStyle.ListIndex)
End Sub

'Change the splitter orientation.
Private Sub cboOrientation_Click()
    vbSplitter1.Orientation = cboOrientation.ItemData(cboOrientation.ListIndex)
End Sub

'Change the splitter color.
Private Sub cboSplitterColor_Click()
    vbSplitter1.SplitterColor = cboSplitterColor.ItemData(cboSplitterColor.ListIndex)
End Sub

'Change the splitter selected color.
Private Sub cboSelectedColor_Click()
    vbSplitter1.SelectedColor = cboSelectedColor.ItemData(cboSelectedColor.ListIndex)
End Sub

'Enable the panels to resize automatically while the splitter is being moved.
Private Sub chkAutoSize_Click()
    vbSplitter1.AutoResize = chkAutoSize.Value
End Sub

'Force the panels to resize.
Private Sub cmdForceResize_Click()
       vbSplitter1.ForceResize
End Sub
'***********************************************************************************************
'Keyboard/Mouse Events
'***********************************************************************************************
'***********************************************************************************************
'Change/Validation Events
'***********************************************************************************************
'Change the Child1 panel Maximum size.
Private Sub txtNewMaxSize_Validate(Cancel As Boolean)
On Error Resume Next
    vbSplitter1.MaxSize = txtNewMaxSize.Text
    txtMaxSize.Text = vbSplitter1.MaxSize
End Sub

'Change the Child1 panel Minimum size.
Private Sub txtNewMinSize_Validate(Cancel As Boolean)
On Error Resume Next
    vbSplitter1.MinSize = txtNewMinSize.Text
    txtMinSize = vbSplitter1.MinSize
End Sub

'Change the control's splitter width.
Private Sub txtNewSplitterWidth_Validate(Cancel As Boolean)
On Error Resume Next
    vbSplitter1.SplitterWidth = txtNewSplitterWidth.Text
    txtSplitterWidth.Text = vbSplitter1.Width
End Sub

'Change the Child1 panel size.
Private Sub txtNewPanelSize_Validate(Cancel As Boolean)
On Error Resume Next
    vbSplitter1.PanelSize = txtNewPanelSize.Text
    txtPanelSize.Text = vbSplitter1.PanelSize
End Sub
'***********************************************************************************************
'Control Events
'***********************************************************************************************
'Event fired after panels are resized.
Private Sub vbSplitter1_Resize()
On Error Resume Next

    'Update the current size properties.
    With vbSplitter1
        txtSplitterWidth.Text = .SplitterWidth
        txtMinSize.Text = .MinSize
        txtMaxSize.Text = .MaxSize
        txtPanelSize.Text = .PanelSize
    End With
End Sub
