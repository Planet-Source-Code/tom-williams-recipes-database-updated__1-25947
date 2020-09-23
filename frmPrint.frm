VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrint 
   BackColor       =   &H00FF8080&
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11160
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   10575
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":08CA
            Key             =   "W95MBX01"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":0D1C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":125E
            Key             =   "ARW06LT"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":16B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":1B02
            Key             =   "LITENING"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":1F54
            Key             =   "TRFFC14"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":23A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":27F8
            Key             =   "MISC36"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":2C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":309C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":34EE
            Key             =   "Camera"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":3600
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":3A52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   5775
      TabIndex        =   4
      Top             =   5925
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "Exit Print View"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copyrecipe"
            Object.ToolTipText     =   "Copy Highlighted Text to Recipe Database Wizard"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh Page"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            Object.ToolTipText     =   "Stop Page Load"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchepicurious"
            Object.ToolTipText     =   "Search Epicurious Recipes"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchfoodtv"
            Object.ToolTipText     =   "Search Food TV Recipes"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchqvc"
            Object.ToolTipText     =   "Search QVC Recipes"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchhungrymonster"
            Object.ToolTipText     =   "Search Hungry Monster Recipes"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchcdkitchen"
            Object.ToolTipText     =   "Search CDKitchen Recipes"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchcopykat"
            Object.ToolTipText     =   "Search CopyKat Recipes"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5910
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   75
      TabIndex        =   1
      Top             =   390
      Width           =   11025
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   5175
      Left            =   75
      TabIndex        =   0
      Top             =   675
      Width           =   11010
      ExtentX         =   19420
      ExtentY         =   9128
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "FrmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "copykat"
            'ToDo: Add 'copykat' button code.
            MsgBox "Add 'copykat' button code."
        Case "cdkitchen"
            'ToDo: Add 'cdkitchen' button code.
            MsgBox "Add 'cdkitchen' button code."
        Case "copyrecipe"
            cmdNewRecipe
        Case "exit"
            cmdExit
        Case "print"
            cmdPrint
        Case "back"
            cmdBack
        Case "forward"
            cmdForward
        Case "refresh"
            cmdRefresh
        Case "stop"
            cmdStop
        Case "searchepicurious"
            SearchEQ
        Case "searchfoodtv"
            SearchFoodTV
        Case "searchqvc"
            SearchQVC
        Case "searchhungrymonster"
            SearchHungryMonster
        Case "searchcopykat"
            SearchCopyKat
        Case "searchcdkitchen"
            SearchCDKitchen
    End Select
End Sub

Private Sub cmdPrint()
    Dim eQuery As OLECMDF
    On Error Resume Next
    eQuery = WB.QueryStatusWB(OLECMDID_PRINT)
    If Err.Number = 0 Then
        If eQuery And OLECMDF_ENABLED Then
            WB.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, "", ""
        Else
            MsgBox "The Print command is currently disabled."
        End If
    Else
        MsgBox "Print command Error: " & Err.Description
    End If
End Sub

Private Sub cmdNewRecipe()
    Dim eQuery As OLECMDF
    On Error Resume Next
    eQuery = WB.QueryStatusWB(OLECMDID_COPY)
    If Err.Number = 0 Then
        If eQuery And OLECMDF_ENABLED Then
            WB.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT, "", ""
        Else
            MsgBox "The Copy command is currently disabled."
        End If
    Else
        MsgBox "Print command Error: " & Err.Description
    End If
    frmNewWebRecipe.Show vbModal, Me
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        WB.Navigate txtAddress
    End If
End Sub

Private Sub WB_DocumentComplete(ByVal pDisp As Object, Url As Variant)
    StatusBar1.Panels(1).Text = "Document Finished."
End Sub

Private Sub WB_DownloadBegin()
    StatusBar1.Panels(1).Text = "Opening Page....."
End Sub

Private Sub WB_DownloadComplete()
    StatusBar1.Panels(1).Text = "Download Finished..."
End Sub

Private Sub WB_FileDownload(Cancel As Boolean)
    StatusBar1.Panels(1).Text = "Beginning Download....."
End Sub

Private Sub WB_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
    StatusBar1.Panels(1).Text = WB.LocationURL
    txtAddress.Text = WB.LocationURL
End Sub

Private Sub WB_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
    ProgressBar1.Max = ProgressMax
    ProgressBar1.Value = Progress
        If Progress = 0 Then
            ProgressBar1.Visible = False
        Else:
            ProgressBar1.Visible = True
        End If
End Sub

Private Sub Web_StatusTextChange(ByVal Text As String)
StatusBar1.Panels(1).Text = Text
End Sub

Private Sub cmdBack()
    On Error Resume Next
    WB.GoBack
End Sub

Private Sub cmdExit()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cmdForward()
    On Error Resume Next
    WB.GoForward
End Sub

Private Sub cmdRefresh()
    On Error Resume Next
    WB.Refresh
End Sub

Private Sub cmdStop()
    On Error Resume Next
    WB.Stop
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        WB.Width = (Width - 240)
        WB.Height = (Height - 500 - txtAddress.Height - StatusBar1.Height - Toolbar1.Height)
        txtAddress.Width = Width - 240
        ProgressBar1.Top = Height - 625
        ProgressBar1.Left = Width - 5505
    End If
End Sub
