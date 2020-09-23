VERSION 5.00
Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDATREP.OCX"
Begin VB.Form frmImport 
   BackColor       =   &H00FF8080&
   Caption         =   "Import Recipes"
   ClientHeight    =   7050
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   9165
   Icon            =   "frmImport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin MSDataRepeaterLib.DataRepeater DataRepeater1 
      Height          =   5925
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10451
      _StreamID       =   -1412567295
      _Version        =   393216
      BackColor       =   16744576
      RowDividerStyle =   3
      CaptionStyle    =   0
      Caption         =   "DataRepeater1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IntegralHeight  =   0   'False
      BeginProperty RepeatedControlName {21FC0FC0-1E5C-11D1-A327-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         Name            =   "RecipesImport.RecipesImportCtl"
      EndProperty
      RepeaterBindings=   5
      BeginProperty RepeaterBinding(0) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Author"
         DataField       =   "Author"
         Key             =   "Author"
      EndProperty
      BeginProperty RepeaterBinding(1) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Import"
         DataField       =   "Import"
         Key             =   "Import"
      EndProperty
      BeginProperty RepeaterBinding(2) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Recipe"
         DataField       =   "Name"
         Key             =   "Recipe"
      EndProperty
      BeginProperty RepeaterBinding(3) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Instructions"
         DataField       =   "Instructions"
         Key             =   "Instructions"
      EndProperty
      BeginProperty RepeaterBinding(4) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "Category"
         DataField       =   "Category"
         Key             =   "Category"
      EndProperty
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9165
      TabIndex        =   0
      Top             =   6315
      Width           =   9165
      Begin VB.CommandButton cmdFilter 
         BackColor       =   &H00FF8080&
         Caption         =   "&View Only Recipes Marked for Import"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2850
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "OFF"
         ToolTipText     =   "Click this button to turn on/off the view only recipes marked for import.  This Button Will Not Import Your Recipes."
         Top             =   0
         Width           =   1440
      End
      Begin VB.CommandButton cmdImport 
         BackColor       =   &H00FF00FF&
         Caption         =   "&Import Recipes Into Recipes Database"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   7350
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1770
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Sa&ve Current Recipe Changes"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1515
      End
      Begin VB.CommandButton cmdUnCheckAll 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&UnSelect All Recipes For Import"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Click this button to UNMARK all recipes for import into Recipes Database"
         Top             =   0
         Width           =   1440
      End
      Begin VB.CommandButton cmdCheckAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Select All Recipes For Import"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Click this button to MARK all recipes for import into Recipes Database"
         Top             =   0
         Width           =   1440
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Close Without Importing Recipes"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5775
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Click this button to Close your Import Preparation Database.  This Button Will Not Import Your Recipes."
         Top             =   0
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim WithEvents rstExport As Recordset
Dim cnnExport As ADODB.Connection
Dim rstExport As ADODB.Recordset
Attribute rstExport.VB_VarHelpID = -1
Dim booImportedAnything As Boolean

Private Sub cmdCheckAll_Click()
    rstExport.MoveFirst
    While Not rstExport.EOF
        rstExport.Fields("Import") = 1
        rstExport.MoveNext
    Wend
    rstExport.MoveFirst
End Sub

Private Sub cmdFilter_Click()
    If cmdFilter.Tag = "OFF" Then
        cmdFilter.Tag = "ON"
        cmdFilter.Caption = "&View All Recipes"
        'Set the Filter so that all of the records don't have to be processed
        rstExport.Filter = "Import = 1"
    Else
        cmdFilter.Tag = "OFF"
        cmdFilter.Caption = "&View Only Recipes Marked for Import"
        'Set the Filter so that all of the records don't have to be processed
        rstExport.Filter = "Import = 1 or Import = 0"
    End If
End Sub

Private Sub cmdImport_Click()

    Dim strSQL As String
    'Set the Filter so that all of the records don't have to be processed
    rstExport.Filter = "Import = 1"
    
    Set cmdChange = New ADODB.Command
    Set rstCategory = New ADODB.Recordset
    rstCategory.CursorType = adOpenStatic
    cnnRecipe.Open
    Set cmdChange.ActiveConnection = cnnRecipe

    'Move to the first Record
    rstExport.MoveFirst
    'Cycle through all records
    While Not rstExport.EOF
        'If Record marked for Import then Import it
        If rstExport.Fields("Import") = 1 Then
            'Set Command Text
            cmdChange.CommandText = "INSERT INTO Recipes " & _
                                    "( Category, Name, Author, Instructions ) " & _
                                    "SELECT """ & Trim(LCase(rstExport(1))) & """ AS Expr1, """ & _
                                                 Trim(Proper(rstExport(2))) & """ AS Expr2, """ & _
                                                 Trim(Proper(rstExport(3))) & """ AS Expr3, """ & _
                                                 rstExport(4) & """ AS Expr4; "
                                    
            'Execute Command Text
            cmdChange.Execute
            'Set flag for use during From_QueryUnload.
            booImportedAnything = True
            
            'Now test to see if the Category exists in the Category Table
            
            'Open connection.
            strSQL = "SELECT Category FROM Category WHERE Category = """ & Trim(LCase(rstExport(1))) & """"
            'Open Category Recordset
            rstCategory.Open strSQL, cnnRecipe, , , adCmdText
            'Test if Category exists or not
            If rstCategory.RecordCount < 1 Then
                With cmdChange
                    'Set Command Text
                    .CommandText = "INSERT INTO Category (Category) SELECT """ & _
                                   LCase(rstExport(1)) & """ AS Expr1;"
                    'Execute Command Text
                    .Execute
                End With
            End If
            'Close the Recordset for next loop.
            Set rstCategory.ActiveConnection = Nothing
        End If
        
        'Move to the next Export Record
        rstExport.MoveNext
    Wend
    
    'Close Connections
    Set rstExport.ActiveConnection = Nothing
    cnnExport.Close
    Set cmdChange = Nothing
    
    'Close Import form when finished
    Unload Me
End Sub

Private Sub cmdUnCheckAll_Click()
    rstExport.MoveFirst
    While Not rstExport.EOF
        rstExport.Fields("Import") = 0
        rstExport.MoveNext
    Wend
    rstExport.MoveFirst
End Sub

Private Sub cmdSave_Click()
    rstExport.Move 0
End Sub

Private Sub Form_Load()
    On Error GoTo Error_Form_Load
    Icon = FrmMain.Icon
    Set cnnExport = New ADODB.Connection
    Set rstExport = New ADODB.Recordset
    
    strXCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & _
              "\Export.mdb"";Mode=ReadWrite;Persist Security Info=False;" & _
              "Jet OLEDB:Compact Without Replica Repair=False"

    cnnExport.Open strXCnn
    
    With rstExport
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .ActiveConnection = cnnExport
        .LockType = adLockOptimistic
        .Open "Select * From Export Order By Import DESC, Category, Name", , , , adCmdText
    End With
    
    Set DataRepeater1.DataSource = rstExport
Error_Form_Load:
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If booImportedAnything = True Then
        'Re-open the Recipes Recordset so that you can see new recipes.
        StartIt
    End If
    'Establish Connection String
    Set rstExport.ActiveConnection = Nothing
    Set rstRecipes.ActiveConnection = Nothing
    Set cnnExport = Nothing
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  DataRepeater1.Height = Me.ScaleHeight - 30 - picButtons.Height
  DataRepeater1.Width = Me.ScaleWidth - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub
