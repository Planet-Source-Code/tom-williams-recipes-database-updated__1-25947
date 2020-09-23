VERSION 5.00
Begin VB.Form frmNewWebRecipe 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Web Recipe Wizard"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6630
   Icon            =   "frmNewWebRecipe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbCategory 
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
      Height          =   360
      Left            =   1125
      TabIndex        =   4
      Top             =   375
      Width           =   3765
   End
   Begin VB.CommandButton cmdCanel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   75
      TabIndex        =   3
      Top             =   4725
      Width           =   1665
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   390
      Left            =   4875
      TabIndex        =   2
      Top             =   4725
      Width           =   1665
   End
   Begin VB.TextBox txtInstructions 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   750
      Width           =   6465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select a Recipe Category or type in the name of a new Recipe Category"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   6435
   End
End
Attribute VB_Name = "frmNewWebRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbCategory_Change()
    If Len(cmbCategory) > 75 Then
        cmbCategory = Mid(cmbCategory, 1, 75)
        cmbCategory.SelStart = 75
        Beep
    End If
End Sub

Private Sub cmbCategory_LostFocus()
    cmbCategory.Text = LCase(cmbCategory.Text)
End Sub

Private Sub cmdCanel_Click()
    Dim strRetVal As String
    strRetVal = MsgBox("Are you sure you want to cancel " & _
                       "the Recipe Wizard?", vbQuestion + vbOKCancel, "Cancel Confirmation")
    If strRetVal = "1" Then
        Unload Me
    End If
End Sub

Private Sub cmdFinish_Click()
    Dim I As Integer, booFound As Boolean
    If cmbCategory.Text <= "" Then
        MsgBox "You must either select an existing Recipe Category " & _
               "or type in a new Recipe Category.", vbOKOnly + vbExclamation, _
               "Category Error"
        Exit Sub
    End If
    
    'Search for Category
    With FrmMain.Tv
        For I = 1 To .Nodes.Count
            If .Nodes.Item(I).Text = cmbCategory.Text Then
                'Category found so select it
                .Nodes(I).Selected = True
                booFound = True
                Exit For
            End If
        Next I
    End With
    'Category not found so Add one for it
    If booFound = False Then
        AddCat cmbCategory.Text
    End If
    
    rstRecipes.AddNew
    FrmMain.mnuMain.Enabled = False
    FrmMain.mnuSearch.Enabled = False
    For I = 2 To FrmMain.Toolbar1.Buttons.Count Step 1
        FrmMain.Toolbar1.Buttons(I).Enabled = False
    Next I
    FrmMain.Tv.Enabled = False
    FrmMain.Toolbar1.Buttons(1).Enabled = True
    
    FrmMain.txtCategory = cmbCategory.Text
    FrmMain.txtAuthor = "Unknown"
    FrmMain.txtRecipe = Mid(Clipboard.GetText, 1, (InStr(1, Clipboard.GetText, vbCrLf) - 1))
    FrmMain.txtInstructions = Mid(Clipboard.GetText, (InStr(1, Clipboard.GetText, vbCrLf) + 1))
    
    Unload Me
    FrmMain.Show

End Sub

Private Sub Form_Load()
    Dim strSQL As String, booCancel As Boolean
    
    If FrmMain.Toolbar1.Buttons(1).Enabled = True Then
        MsgBox "You cannot add another recipe until you save the current recipe."
        Unload Me
        booCancel = True
        FrmMain.Show
    End If
    
    If booCancel <> True Then
        Me.Icon = FrmMain.Icon
        Me.Caption = "Web Based Recipe Wizard"
        
        ' Open connection.
        strSQL = "SELECT Category From Category"
        strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & _
                 "\Recipe.mdb"";Mode=ReadWrite;Persist Security Info=False;" & _
                 "Jet OLEDB:Compact Without Replica Repair=False"
        
        Set cnnRecipe = New ADODB.Connection
        
        If cnnRecipe.State = adStateClosed Then
            cnnRecipe.CursorLocation = adUseClient
            cnnRecipe.Open strCnn
        End If
                 
        'Set cnnRecipe = New ADODB.Connection
        Set rstCategory = New ADODB.Recordset
        'cnnRecipe.Open strCnn
        rstCategory.CursorType = adOpenStatic
        rstCategory.Open strSQL, cnnRecipe, , , adCmdText
        If rstCategory.RecordCount > 0 Then
            With rstCategory
                .MoveFirst
                While Not rstCategory.EOF
                    Me.cmbCategory.AddItem (.Fields(0))
                    .MoveNext
                Wend
            End With
        End If
        
        Set rstCategory.ActiveConnection = Nothing
        cnnRecipe.Close
    End If
Exit_Form_Load:

End Sub
