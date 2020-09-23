VERSION 5.00
Begin VB.UserControl RecipesImportCtl 
   BackColor       =   &H00FF8080&
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8475
   DataBindingBehavior=   1  'vbSimpleBound
   ScaleHeight     =   2325
   ScaleWidth      =   8475
   Begin VB.CheckBox chkImport 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Import   this recipe?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   75
      TabIndex        =   8
      Top             =   1500
      Value           =   1  'Checked
      Width           =   930
   End
   Begin VB.TextBox txtRecipe 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1125
      TabIndex        =   2
      Top             =   675
      Width           =   3525
   End
   Begin VB.TextBox txtInstructions 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1125
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   975
      Width           =   7290
   End
   Begin VB.TextBox txtAuthor 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1125
      TabIndex        =   1
      Top             =   375
      Width           =   3525
   End
   Begin VB.TextBox txtCategory 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1125
      TabIndex        =   0
      Top             =   75
      Width           =   3525
   End
   Begin VB.Label lblCarriageReturnInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   "To Add a Carriage Return in your Instructions Field, simultaneously hit [CTL]+[ENTER]"
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
      Height          =   240
      Left            =   1125
      TabIndex        =   10
      Top             =   2100
      Width           =   7140
   End
   Begin VB.Label lblSave 
      BackStyle       =   0  'Transparent
      Caption         =   $"RecipesImport.ctx":0000
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   990
      Left            =   4725
      TabIndex        =   9
      Top             =   0
      Width           =   3795
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   "Instructions"
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
      Height          =   255
      Left            =   75
      TabIndex        =   7
      Top             =   975
      Width           =   975
   End
   Begin VB.Label lblRecipe 
      BackStyle       =   0  'Transparent
      Caption         =   "Recipe"
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
      Height          =   255
      Left            =   75
      TabIndex        =   6
      Top             =   675
      Width           =   975
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
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
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Top             =   375
      Width           =   975
   End
   Begin VB.Label lblCategory 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Height          =   255
      Left            =   75
      TabIndex        =   4
      Top             =   75
      Width           =   975
   End
End
Attribute VB_Name = "RecipesImportCTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Category() As String
Attribute Category.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Category.VB_MemberFlags = "101c"
    Category = txtCategory.Text
End Property

Public Property Let Category(ByVal strCategory As String)
   'If CanPropertyChange("Category") Then
      txtCategory.Text = strCategory
      ' The following line tells Visual Basic the
      ' property has changed--if you omit this line,
      ' the data source will not be updated!
      PropertyChanged "Category"
   'End If
End Property

Private Sub txtCategory_LostFocus()
    txtCategory = LCase(txtCategory)
End Sub

Private Sub txtCategory_Change()
    PropertyChanged "Category"
End Sub

Public Property Get Author() As String
Attribute Author.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Author.VB_MemberFlags = "101c"
    Author = txtAuthor.Text
End Property

Public Property Let Author(ByVal strAuthor As String)
   'If CanPropertyChange("Author") Then
      txtAuthor.Text = strAuthor
      ' The following line tells Visual Basic the
      ' property has changed--if you omit this line,
      ' the data source will not be updated!
      PropertyChanged "Author"
   'End If
End Property

Private Sub txtAuthor_Change()
    PropertyChanged "Author"
End Sub

Public Property Get Recipe() As String
Attribute Recipe.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Recipe.VB_MemberFlags = "101c"
    Recipe = txtRecipe.Text
End Property

Public Property Let Recipe(ByVal strRecipe As String)
   'If CanPropertyChange("Recipe") Then
      txtRecipe.Text = strRecipe
      ' The following line tells Visual Basic the
      ' property has changed--if you omit this line,
      ' the data source will not be updated!
      PropertyChanged "Recipe"
   'End If
End Property

Private Sub txtRecipe_Change()
    PropertyChanged "Recipe"
End Sub

Public Property Get Instructions() As String
Attribute Instructions.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Instructions.VB_MemberFlags = "101c"
    Instructions = txtInstructions.Text
End Property

Public Property Let Instructions(ByVal strInstructions As String)
   'If CanPropertyChange("Instructions") Then
      txtInstructions.Text = strInstructions
      ' The following line tells Visual Basic the
      ' property has changed--if you omit this line,
      ' the data source will not be updated!
      PropertyChanged "Instructions"
   'End If
End Property

Private Sub txtInstructions_Change()
    PropertyChanged "Instructions"
End Sub

Public Property Get Import() As Integer
Attribute Import.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Import.VB_MemberFlags = "101c"
    Import = chkImport.Value
End Property

Public Property Let Import(ByVal intImport As Integer)
   'If CanPropertyChange("Import") Then
      chkImport.Value = intImport
      ' The following line tells Visual Basic the
      ' property has changed--if you omit this line,
      ' the data source will not be updated!
      PropertyChanged "Import"
   'End If
End Property

Private Sub chkImport_Click()
    PropertyChanged "Import"
End Sub
