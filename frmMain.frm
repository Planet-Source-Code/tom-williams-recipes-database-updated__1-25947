VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FF8080&
   ClientHeight    =   6585
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11010
   DrawMode        =   11  'Not Xor Pen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "save"
            Description     =   "Save Recipe"
            Object.ToolTipText     =   "Save Recipe"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deleterecipe"
            Description     =   "Delete Recipe"
            Object.ToolTipText     =   "Delete Recipe"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deletecategory"
            Description     =   "Delete Category"
            Object.ToolTipText     =   "Delete Category"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertrecipe"
            Description     =   "Insert Recipe"
            Object.ToolTipText     =   "Insert Recipe"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertcategory"
            Description     =   "Insert Category"
            Object.ToolTipText     =   "Insert Category"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "importrecipe"
            Object.ToolTipText     =   "Import Recipe"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exportrecipe"
            Object.ToolTipText     =   "Export Recipe"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "backup"
            Description     =   "Backup Database"
            Object.ToolTipText     =   "Backup Database"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "restore"
            Description     =   "Restore Database"
            Object.ToolTipText     =   "Restore Database"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchepecurious"
            Description     =   "Search Epecurious"
            Object.ToolTipText     =   "Search Epecurious Recipes"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchfoodtv"
            Description     =   "Search FoodTv"
            Object.ToolTipText     =   "Search FoodTv Recipes"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchqvcrecipes"
            Description     =   "Search QVC Recipes"
            Object.ToolTipText     =   "Search QVC Recipes"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchhungrymonster"
            Object.ToolTipText     =   "Search Hungry Monster Recipes"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchcdkitchen"
            Object.ToolTipText     =   "Search CDKitchen Recipes"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "searchcopykat"
            Object.ToolTipText     =   "Search CopyKat Recipes"
            ImageIndex      =   21
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10875
      Top             =   1200
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   6330
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "Current Time"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.ToolTipText     =   "Current Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Total Recipes in Current Category"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4145
            MinWidth        =   4145
            Object.ToolTipText     =   "Total Recipes in Recipes Database"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4145
            MinWidth        =   4145
            Object.ToolTipText     =   "Total Categories"
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
   Begin VB.TextBox txtCategory 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   735
      Width           =   3015
   End
   Begin VB.CommandButton cmdNormalize 
      Caption         =   "&Normalize DB"
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   450
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9450
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtInstructions 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   4425
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2535
      Width           =   6495
   End
   Begin VB.TextBox txtAuthor 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Top             =   1935
      Width           =   3015
   End
   Begin VB.TextBox txtRecipe 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   2
      Top             =   1335
      Width           =   3015
   End
   Begin MSComctlLib.TreeView Tv 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   495
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   10186
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   88
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   10350
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22EC
            Key             =   "DISKS04"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":273E
            Key             =   "TRAFFIC"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B90
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30D2
            Key             =   "TRASH02A"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3524
            Key             =   "TRASH02B"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3976
            Key             =   "MISC33"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DC8
            Key             =   "MISC29"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":421A
            Key             =   "ARW05DN"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":466C
            Key             =   "ARW05UP"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4ABE
            Key             =   "FILES03B"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F10
            Key             =   "FILES04"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5362
            Key             =   "CIRC1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":567C
            Key             =   "CIRC2"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5996
            Key             =   "CIRC3"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5CB0
            Key             =   "MISC34"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6102
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6554
            Key             =   "MISC36"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69A6
            Key             =   "MISC37"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DF8
            Key             =   "W95MBX01"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":724A
            Key             =   "MISC01"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":769C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCategory 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4440
      TabIndex        =   9
      Top             =   495
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   7500
      Picture         =   "frmMain.frx":7AEE
      Stretch         =   -1  'True
      Top             =   450
      Width           =   3495
   End
   Begin VB.Label lblInstructions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preparation Instructions"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4440
      TabIndex        =   8
      Top             =   2295
      Width           =   2040
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4425
      TabIndex        =   7
      Top             =   1695
      Width           =   570
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipe Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4440
      TabIndex        =   6
      Top             =   1095
      Width           =   1155
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Recipe Main Menu"
      Begin VB.Menu mnuSave 
         Caption         =   "Save Recipe"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Out Recipe"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add New Recipe"
      End
      Begin VB.Menu mnuDelRec 
         Caption         =   "Delete Checked Recipes"
      End
      Begin VB.Menu mnuDelCat 
         Caption         =   "Delete Catagory From DataBase"
      End
      Begin VB.Menu mnuAddCat 
         Caption         =   "Add Recipe Catagory to DataBase"
      End
      Begin VB.Menu mnuSeparator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBkUp 
         Caption         =   "BackUp Recipe DataBase"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore DataBase from BackUp"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import Recipes"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Checked Recipes"
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit Application"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchEQ 
         Caption         =   "Search Epicurious Recipes"
      End
      Begin VB.Menu mnuSearchFoodTV 
         Caption         =   "Search FoodTv Recipes"
      End
      Begin VB.Menu mnuSearchQVC 
         Caption         =   "Search QVC Recipes"
      End
      Begin VB.Menu mnuSearchHungryMonster 
         Caption         =   "Search HungryMonster"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdNormalize_Click()
    Normalize
    MsgBox "Done."
End Sub

Private Sub Form_Load()
    'Compact Database
    Dim strSource As String
    Dim strTarget As String
    
    strSource = App.Path & "\Recipe.mdb"
    strTarget = App.Path & "\Compact.mdb"
    DBEngine.CompactDatabase strSource, strTarget
    
    'Delete Old Database
    Kill (strSource)
    
    'Copy the Compact.mdb DataBase back to Recipes.mdb
    FileCopy strTarget, strSource
    
    'Kill Old Compact
    Kill (strTarget)
    
    'Call function to load the DataBase catagories and Records
    'into the Treeview
    StartIt
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'just incase user has one last change to add
    CloseRs
End Sub

Private Sub Form_Resize()
'form withd at startup 11130
    If Me.WindowState = vbMinimized Then Exit Sub
    If (Width - 7050) > 4080 Then
        If (Width - 7050) > 5190 Then
            Tv.Width = 5190
            Me.txtInstructions.Width = Me.Width - Tv.Width - 420
        Else
            Tv.Width = (Width - 7050) + 120
            Me.txtInstructions.Width = Me.Width - Tv.Width - 420
        End If
    Else
        Tv.Width = 4215
    End If
    
    
    If Me.Height - Toolbar1.Height - 900 - StatusBar1.Height >= 500 Then
        Tv.Height = Me.Height - Toolbar1.Height - 900 - StatusBar1.Height
    Else
        Tv.Height = 500
    End If
    
    'If Me.Height > 3825 Then
    If Me.Height - Toolbar1.Height - Image1.Height - 880 - StatusBar1.Height > 285 Then
        Me.txtInstructions.Height = Me.Height - Toolbar1.Height - Image1.Height - 880 - StatusBar1.Height
    Else
        Me.txtInstructions.Height = 285
    End If
    
    Me.txtAuthor.Left = Me.Tv.Width + 240
    Me.txtCategory.Left = Me.Tv.Width + 240
    Me.txtRecipe.Left = Me.Tv.Width + 240
    
    Me.txtAuthor.Width = Me.txtInstructions.Width - 3480
    Me.txtCategory.Width = Me.txtInstructions.Width - 3480
    Me.txtRecipe.Width = Me.txtInstructions.Width - 3480
    
    If Me.Width > 11130 Then
        Me.Image1.Left = Me.txtAuthor.Left + Me.txtAuthor.Width ' + 100
    Else
        Me.Image1.Left = 7500
    End If
    
    Me.txtInstructions.Left = Me.Tv.Width + 240
    
    Me.lblCategory.Left = Me.Tv.Width + 240
    Me.lblAuthor.Left = Me.Tv.Width + 240
    Me.lblInstructions.Left = Me.Tv.Width + 240
    Me.lblName.Left = Me.Tv.Width + 240
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub mnuAdd_Click()
    On Error Resume Next
    Dim I As Integer
    
    'User wants to add a Record
    'Makes sure user has choosen a Catagory to add a Record to
    If Mid(Tv.SelectedItem.Key, 1, 3) <> "Cat" Then
        MsgBox "Please Select a Category before choosing to Add a Recipe"
        Exit Sub
    End If

    rstRecipes.AddNew
    mnuMain.Enabled = False
    mnuSearch.Enabled = False
    For I = 2 To Toolbar1.Buttons.Count Step 1
        Toolbar1.Buttons(I).Enabled = False
    Next I
    Tv.Enabled = False
    Toolbar1.Buttons(1).Enabled = True
    
    txtCategory = Tv.SelectedItem.Text

End Sub


Private Sub mnuAddCat_Click()
    'Call Function to add a Catagory (Table)
    AddCat ("")
End Sub

Private Sub mnuBkUp_Click()
    'User wants to BackUp the DataBase
    
    'Call Function to Create the Directory
    DirMk

    'Change the mouse Pointer to an HourGlass
    FrmMain.MousePointer = 11
    
    'Make sure all Open Files are Closed
    Close
    If Dir(App.Path & "\BackUpDB.bat") > "" Then Kill (App.Path & "\BackUpDB.bat")
    If Dir(App.Path & "\DataBackup\BackUp.mdb") > "" Then Kill (App.Path & "\DataBackup\BackUp.mdb")
    
    'Copy the DataBase to the DataBackup dir
    FileCopy App.Path & "\Recipe.mdb", App.Path & "\DataBackup\BackUp.mdb"

    'Change the Mouse Pointer back to standard
    FrmMain.MousePointer = 0

    'let user know DataBase has been Backed Up
    MsgBox "DataBase Successfully Backed Up to " & App.Path & "\DataBackup\BackUp.mdb"
End Sub

Private Sub mnuDelCat_Click()
    'user wants to delete a Catagory
    DelCat
End Sub

Private Sub mnuDelRec_Click()
    'User wants to delete the shown Record
    Dim intResult As String
    
    intResult = MsgBox("Are you sure you want to delete all Checked Recipes? " _
                      , vbQuestion + vbOKCancel, "Confirm Recipe Delete")
    If intResult = 0 Then Exit Sub
    
    DelRec
End Sub

Private Sub mnuExit_Click()
    'User wants exit the Program
    Unload Me
End Sub

Private Sub mnuExport_Click()

    Dim lngExportCount As Long
    'Create a database to hold the Info
    CreateExportDB
    
    'Establish Connection String
    strXCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & _
             "\Export.mdb"";Mode=ReadWrite;Persist Security Info=False;" & _
             "Jet OLEDB:Compact Without Replica Repair=False"
    
    Set cnnExport = New ADODB.Connection
    Set rstExport = New ADODB.Recordset
    
    'Test if connection not open.  If not open it.
    If cnnExport.State = adStateClosed Then
        cnnExport.CursorLocation = adUseClient
        cnnExport.Open strXCnn
    End If
    
    'Populate Recordset.
    rstExport.Open "Export", cnnExport, adOpenDynamic, adLockBatchOptimistic, adCmdTable

    
    'Put the info into the database
    For I = 1 To (intTtlRecipeCount + intTtlCategories)
        If Tv.Nodes.Item(I).Checked = True Then
            If InStr(1, Tv.Nodes.Item(I).Key, "_") Then
                rstRecipes.AbsolutePosition = ((Replace(Tv.Nodes.Item(I).Key, "_", "")) + 1)
                rstExport.AddNew
                rstExport.Fields(0).Value = 1
                rstExport.Fields(1).Value = txtCategory & " "
                rstExport.Fields(2).Value = txtRecipe & " "
                rstExport.Fields(3).Value = txtAuthor & " "
                rstExport.Fields(4).Value = txtInstructions & " "
                lngExportCount = lngExportCount + 1
            End If
        End If
    Next
    rstExport.UpdateBatch
    Set rstExport.ActiveConnection = Nothing
    Set cnnExport = Nothing
    
    If lngExportCount <= 0 Then
        MsgBox "You have not exported any recipes.  You may want to try again and be sure " & _
               "that you have placed a Check Mark next to the recipes that you have chosen to export." _
               , vbQuestion + vbOKOnly, "No Recipes Exported"
    Else
        MsgBox "You have successfully exported " & lngExportCount & " records to " & _
               " the " & App.Path & "\Export.mdb file." & _
               vbCrLf & vbCrLf & vbCrLf & "Be sure to tell the person that you " & _
               "are sending this to that they must place the Export.mdb file " & _
               "into the folder that they installed the application to in order " & _
               "successfully be able to import it into their Recipes database." _
               , vbCritical + vbOKOnly, "Export Completed"
    End If
End Sub

Private Sub mnuImport_Click()
    On Error Resume Next
    If Dir(App.Path & "\Export.mdb") <= "" Then
        MsgBox "You must place your Export.mdb file in the application path " & _
               "before you can continue. " & vbCrLf & vbCrLf & vbCrLf & App.Path _
               , vbCritical + vbOKOnly, "Import Error"
    End If
    cnnRecipe.Open strCnn                           'reopen connection
    Set rstRecipes.ActiveConnection = cnnRecipe     'reconnect recordset
    rstRecipes.UpdateBatch                          'Update Recipes
    Set rstRecipes.ActiveConnection = Nothing
    cnnRecipe.Close
    
    frmImport.Show vbModal, Me
End Sub

Private Sub mnuPrint_Click()
'User wants to Print the Record Shown

'Make sure has choosen a record to Print
If Trim(txtRecipe) = "" Then
    MsgBox "Please first select a recipe to print.", vbOKOnly + vbInformation, "What Recipe?"
    Exit Sub
End If

'Make sure all Open Files are Closed
Close

'Create a File to Hold the Info
Open App.Path & "\print.txt" For Output As #1

'Put the info into the File
Print #1, "Name: " & Trim(txtRecipe)
Print #1, ""
Print #1, "Author: " & Trim(txtAuthor)
Print #1, ""
Print #1, ""
Print #1, "Instructions:"
Print #1, Trim(txtInstructions)
'Close the File
Close #1

'Make sure all open files are Closed
Close

'Open the Htm file used by the FrmPrint Form
Open App.Path & "\print.htm" For Output As #1

'Open the Created File to get the Info from
Open App.Path & "\print.txt" For Input As #2

'Put the Header into #1 File
Print #1, "<HTML>"
Print #1, "<HEAD>"
Print #1, "    <TITLE>My Recipes by Tom Wiliams Jr.</TITLE>"
Print #1, "</HEAD>"

Print #1, "<BODY><FONT SIZE=" & Chr(34) & "3" & Chr(34) & " FACE=" & Chr(34) & "Comic Sans MS" & Chr(34) & ">"
Print #1, "<Marquee Behavior=ALTERNATE SCROLLDELAY=5 SCROLLAMOUNT=5 " & _
          "TITLE=""Tom's Killer Visual Basic Coding""> Printed " & _
          "from the ""My Recipes Application"" coded by <B><I>Tom Williams Jr.</B></I></U></MARQUEE><br><br><B>"

'Get the Info from File #2
Do Until EOF(2)
Line Input #2, a$

'Put the Info from File #2 into File #1
Print #1, a$ & "<br>"
Loop

'Put in the Tail for File #1
Print #1, "</B></FONT></BODY>"
Print #1, "</HTML>"

'Close Opened Files
Close #1
Close #2

'set the caption for FrmPrint Form
FrmPrint.Caption = "Print Current Recipe or Search For Other Recipes..."

'Call the Print Form
FrmPrint.Show
'Open the Htm file Created from the FrmMain
FrmPrint.WB.Navigate App.Path & "\print.htm"

End Sub

Private Sub mnuRestore_Click()
    'User wants to Restore the DataBase from the BackUp DataBase
    
    'Call Function to Create the Directory Just incase
    'so Program want error out
    DirMk
    
    'make sure file is not already opened
    Close
    
    'Check to be sure user has backed up the DataBase
    Open App.Path & "\DataBkUp\BackUp.mdb" For Binary As #1
    If LOF(1) = 0 Then
        Close #1
        Kill App.Path & "\DataBkUp\BackUp.mdb"
        MsgBox "Sorry, No Backup Exists To Restore From"
        Exit Sub
    End If

    'Make sure all Open Files are Closed
    Close
    'Let the User know that this action will Over Write The Present
    'DataBase
    Dim Msg, Style, Title, Response
    Msg = "WARNNING !!!" & vbCrLf & _
        "This Action will OVER WRITE Your Recipe DataBase" & vbCrLf & _
        "With Your Latest BackUp DataBase" & vbCrLf & _
        "DO YOU WANT TO CONTINUE ?"
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "Restore DataBase"
    Response = MsgBox(Msg, Style, Title)

    If Response = vbYes Then
        
        'Call Function To close the DataBase just in case it
        'is Open somwhere
        CloseRs
        
        'Change the Mouse Pointer to an HourGlass
        FrmMain.MousePointer = 11
        
        'Copy the BackUp DataBase to the app.Path as the Recipe.mdb
        FileCopy App.Path & "\DataBkUp\BackUp.mdb", App.Path & "\Recipe.mdb"
        
        'Change the Mouse Pointer back to standard
        FrmMain.MousePointer = 0
        
        'Call Function to Load the New DataBase into the Treeview
        StartIt
    End If

End Sub

Private Sub mnuSave_Click()
    On Error Resume Next
    If txtCategory = "" _
     Or txtRecipe = "" _
      Or txtAuthor = "" _
       Or txtInstructions = "" Then
            MsgBox "Please fill in all of the recipe information."
            Exit Sub
    End If
    
    Tv.Nodes.Add Tv.SelectedItem, tvwChild, intTtlRecipeCount & "_", txtRecipe, 4, 3
    intTtlRecipeCount = intTtlRecipeCount + 1
    rstRecipes.Update
    
    Me.mnuMain.Enabled = True
    Me.mnuSearch.Enabled = True
    Me.Tv.Enabled = True
    Me.Toolbar1.Buttons(1).Enabled = False
    For I = 2 To Toolbar1.Buttons.Count Step 1
        Toolbar1.Buttons(I).Enabled = True
    Next I
    
    'let user know Record has been added
    MsgBox "Recipe Added"
End Sub

Private Sub Timer1_Timer()
    StatusBar1.Panels(1) = Time
    If StatusBar1.Panels(2) <> Date Then
        StatusBar1.Panels(2) = Date
    End If
    If StatusBar1.Panels(3) <> intCatRecipeCount & " Current Category Recipes" Then
        StatusBar1.Panels(3) = intCatRecipeCount & " Current Category Recipes"
    End If
    If StatusBar1.Panels(4) <> intTtlRecipeCount & " Total Recipes" Then
        StatusBar1.Panels(4) = intTtlRecipeCount & " Total Recipes"
    End If
    If StatusBar1.Panels(5) <> intTtlCategories & " Total Categories" Then
        StatusBar1.Panels(5) = intTtlCategories & " Total Categories"
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "save"
            mnuSave_Click
        Case "exit"
            mnuExit_Click
        Case "print"
            mnuPrint_Click
        Case "deleterecipe"
            mnuDelRec_Click
        Case "deletecategory"
            mnuDelCat_Click
        Case "insertrecipe"
            mnuAdd_Click
        Case "insertcategory"
            mnuAddCat_Click
        Case "importrecipe"
            mnuImport_Click
        Case "exportrecipe"
            mnuExport_Click
        Case "backup"
            mnuBkUp_Click
        Case "restore"
            mnuRestore_Click
        Case "searchepecurious"
            SearchEQ
        Case "searchfoodtv"
            SearchFoodTV
        Case "searchqvcrecipes"
            SearchQVC
        Case "searchhungrymonster"
            SearchHungryMonster
        Case "searchcopykat"
            SearchCopyKat
        Case "searchcdkitchen"
            SearchCDKitchen
    End Select
End Sub

Private Sub Tv_DragDrop(Source As Control, X As Single, y As Single)
    On Error Resume Next
    Dim intRec As Integer

    If Mid(SourceNode.Key, 1, 3) = "Cat" Then Exit Sub
    If SourceNode.Key = "BOOK" Then Exit Sub

    ' If it's the same as last time, do nothing.
    If Tv.SelectedItem Is Tv.DropHighlight Then Exit Sub
    
    intRec = Replace(SourceNode.Key, "_", "") + 1

    If Not (Tv.DropHighlight Is Nothing) Then
        ' It's a valid drop. Set source node's
        ' parent to be the target node.
        If Mid(Tv.DropHighlight.Key, 1, 3) = "Cat" Then
            Set SourceNode.Parent = Tv.DropHighlight
            rstRecipes.AbsolutePosition = intRec
            rstRecipes.Fields(0) = Tv.DropHighlight.Text
            txtCategory = Tv.DropHighlight.Text
        Else
            Tv.Nodes.Remove (SourceNode.Index)
            Tv.Nodes.Add Tv.DropHighlight, tvwNext, _
            SourceNode.Key, SourceNode.Text, SourceNode.Image, _
            SourceNode.SelectedImage
            rstRecipes.AbsolutePosition = intRec
            rstRecipes.Fields(0) = Tv.DropHighlight.Parent.Text
            txtCategory = Tv.DropHighlight.Parent.Text
        End If
        rstRecipes.Update
        Set Tv.DropHighlight = Nothing
    End If

    Set SourceNode = Nothing

End Sub

Private Sub Tv_DragOver(Source As Control, X As Single, y As Single, State As Integer)
    On Error Resume Next

    Dim target As Node
    Dim highlight As Boolean

    ' See what node we're above.
    Set target = Tv.HitTest(X, y)
    
    ' If it's the same as last time, do nothing.
    If target Is TargetNode Then Exit Sub
    Set TargetNode = target
    
    highlight = False
    If Not (TargetNode Is Nothing) Then
        ' See what kind of node were above.
        highlight = True
    End If
    
    If highlight Then
        Set Tv.DropHighlight = TargetNode
    Else
        Set Tv.DropHighlight = Nothing
    End If
End Sub

Private Sub Tv_Expand(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
End Sub

Private Sub Tv_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    ' Set the item being dragged.
    Set SourceNode = Tv.HitTest(X, y)
End Sub

Private Sub Tv_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then
        ' Start a new drag.
        If y = Tv.Height Then
            
        End If
        
        ' Select this node. When no node is highlighted,
        ' this node will be displayed as selected. That
        ' shows where it will land if dropped.
        Set Tv.SelectedItem = SourceNode

        ' Fire the Begin Drag
        Tv.Drag vbBeginDrag
    End If
End Sub

Private Sub Tv_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim I As Integer
    Select Case Mid(Node.Key, 1, 3)
    Case "Cat"
        If Node.Checked = True Then
            For I = 1 To Node.Children
                Tv.Nodes(Node.Index + I).Checked = True
            Next
        Else
            For I = 1 To Node.Children
                Tv.Nodes(Node.Index + I).Checked = False
            Next
        End If
    Case "BOO"
        If Node.Checked = True Then
            For I = 1 To intTtlRecipeCount + intTtlCategories
                Tv.Nodes(Node.Index + I).Checked = True
            Next
        Else
            For I = 1 To intTtlRecipeCount + intTtlCategories
                Tv.Nodes(Node.Index + I).Checked = False
            Next
        End If
    End Select
        
End Sub

Private Sub Tv_NodeClick(ByVal Node As MSComctlLib.Node)
    'this is what does the stuff when a user click on
    'a Itemin the TreeView
        Select Case Mid(Node.Key, 1, 3)
            Case "BOO"
                'we do nothing here
            Case "Cat"
                'If user clicks Catagory then populate category count Variable
                intCatRecipeCount = Node.Children
            Case Else
                If rstRecipes.AbsolutePosition < 0 Then
                    rstRecipes.MoveFirst
                End If
                If Node.Key = "0_" Then
                    rstRecipes.MoveFirst
                Else
                    rstRecipes.AbsolutePosition = ((Replace(Node.Key, "_", "")) + 1)
                End If
                intCatRecipeCount = Node.Parent.Children
        End Select
End Sub

Private Sub txtAuthor_Change()
    If Len(txtAuthor) > 75 Then
        txtAuthor = Mid(txtAuthor, 1, 75)
        txtAuthor.SelStart = 75
        Beep
    End If
End Sub

Private Sub txtAuthor_LostFocus()
    txtAuthor = Proper(txtAuthor)
    If Not rstRecipes.EOF And Not rstRecipes.BOF Then
        rstRecipes.Update
    End If
End Sub

Private Sub txtCategory_Change()
    If Len(txtCategory) > 75 Then
        txtCategory = Mid(txtCategory, 1, 75)
        txtCategory.SelStart = 75
        Beep
    End If
End Sub

Private Sub txtInstructions_LostFocus()
    If Not rstRecipes.EOF And Not rstRecipes.BOF Then
        rstRecipes.Update
    End If
End Sub

Private Sub txtRecipe_Change()
    If Len(txtRecipe) > 75 Then
        txtRecipe = Mid(txtRecipe, 1, 75)
        txtRecipe.SelStart = 75
        Beep
    End If
End Sub

Private Sub txtRecipe_LostFocus()
    txtRecipe = Proper(txtRecipe)
    If Not rstRecipes.EOF And Not rstRecipes.BOF Then
        rstRecipes.Update
    End If
End Sub
