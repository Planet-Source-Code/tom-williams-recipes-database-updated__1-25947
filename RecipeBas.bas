Attribute VB_Name = "RecipeBas"
Option Explicit
' Set up global Virables
'Connections
Public cnnRecipe As ADODB.Connection
Public cnnExport As ADODB.Connection

'Execute Commands
Public cmdChange As ADODB.Command

'RecordSets
Public rstExport As ADODB.Recordset
Public rstRecipes As ADODB.Recordset
Public rstCategory As ADODB.Recordset

' Set Category Count Variable
Public intCatCount As Integer
Public intCatRecipeCount As Integer
Public intTtlRecipeCount As Integer
Public intTtlCategories As Integer
Public strCnn As String
Public strXCnn As String

' Set Drag and Drop variables
Public SourceNode As Object
Public TargetNode As Object

Public Function Proper(X)
    Dim temp$, C$, OldC$, I As Integer
    If IsNull(X) Then
        Exit Function
    Else
        temp$ = CStr(LCase(X))
        ' Initialize OldC$ to a single space because first
        ' letter needs to be capitalized but has no preceding letter.
        OldC$ = " "
        For I = 1 To Len(temp$)
            C$ = Mid$(temp$, I, 1)
            If C$ >= "a" And C$ <= "z" And (OldC$ < "a" Or OldC$ > "z") Then
                Mid$(temp$, I, 1) = UCase$(C$)
            End If
            OldC$ = C$
        Next I
        Proper = temp$
    End If
End Function

Public Function StartIt()
    'Function To Load the Data into TreeView (tv)
    'Loads All The Catagories and Recipes
    'On Error Resume Next

    'Set Virables
    Dim tvNode As Node
    Dim strCategory As String
    Dim I As Integer
    Dim booExists As Boolean
    Dim strSQL As String
    
    strCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" & App.Path & _
             "\Recipe.mdb"";Mode=ReadWrite;Persist Security Info=False;" & _
             "Jet OLEDB:Compact Without Replica Repair=False"
    
    Set cnnRecipe = New ADODB.Connection
    Set rstRecipes = New ADODB.Recordset
    
    If cnnRecipe.State = adStateClosed Then
        cnnRecipe.CursorLocation = adUseClient
        cnnRecipe.Open strCnn
    End If
    
    rstRecipes.Open "Select * From Recipes Order By Category, Name", cnnRecipe, adOpenStatic, adLockBatchOptimistic, adCmdText
    

    'Reset Field Owners to Speed Up Refresh
    FrmMain.txtCategory.DataField = ""
    FrmMain.txtRecipe.DataField = ""
    FrmMain.txtAuthor.DataField = ""
    FrmMain.txtInstructions.DataField = ""
    
    'Clear TreeView
    FrmMain.Tv.Nodes.Clear

    'Put in the Main Node
    Set tvNode = FrmMain.Tv.Nodes.Add(, , "BOOK", "My Recipe Book", 1)
    
    'Make sure Main Node is Visible
    intCatCount = 1
    With rstRecipes
        If .RecordCount = 0 Then Exit Function
        'Open The RecordSet for each Child Node (Catagory)
        'start at the first Record
        .MoveFirst
        strCategory = "Cat" & intCatCount
        Set tvNode = FrmMain.Tv.Nodes.Add("BOOK", tvwChild, _
                    strCategory, .Fields(0), 2, 3)
        While Not .EOF
            If FrmMain.Tv.Nodes(strCategory).Text <> .Fields(0) Then
                'Put the Child Nodes into the TreeView (Catagories)
                intCatCount = intCatCount + 1
                strCategory = "Cat" & intCatCount
                Set tvNode = FrmMain.Tv.Nodes.Add("BOOK", tvwChild, _
                            strCategory, .Fields(0), 2, 3)
                'Set strCategory Variable
            End If
            'Put in the children Node
            'Place "_" after absolute position so that it in not interprited as a true Index Number
            Set tvNode = FrmMain.Tv.Nodes.Add(strCategory, tvwChild, (.AbsolutePosition - 1) & "_", _
                         .Fields(1) & "", 4, 3)
            intTtlRecipeCount = intTtlRecipeCount + 1
            'Make sure the Child Nodes are Visible
            .MoveNext
        Wend
    End With
    intTtlCategories = intCatCount
    
    FrmMain.Tv.Nodes.Item("BOOK").Expanded = True
    
    ' Open connection.
    strSQL = "SELECT Category.Category From Category WHERE (((Category.Category) " & _
              "Not In (Select Distinct Category from Recipes where Category.Category = Recipes.Category)))"
             
             
    'Set cnnRecipe = New ADODB.Connection
    Set rstCategory = New ADODB.Recordset
    'cnnRecipe.Open strCnn
    rstCategory.CursorType = adOpenStatic
    rstCategory.Open strSQL, cnnRecipe, , , adCmdText
    ' Add Categories that do not exist in Recipes table but do exist in Category table
    If rstCategory.RecordCount > 0 Then
        With rstCategory
            .MoveFirst
            While Not rstCategory.EOF
                intCatCount = intCatCount + 1
                Set tvNode = FrmMain.Tv.Nodes.Add("BOOK", tvwChild, _
                            "Cat" & intCatCount, .Fields(0), 2, 3)
                .MoveNext
            Wend
        End With
    End If
    
    'Set Fields Back To Original DataSource
    Set FrmMain.txtCategory.DataSource = rstRecipes
    Set FrmMain.txtRecipe.DataSource = rstRecipes
    Set FrmMain.txtAuthor.DataSource = rstRecipes
    Set FrmMain.txtInstructions.DataSource = rstRecipes

    'Set Fields Back To Original DB Owner
    FrmMain.txtCategory.DataField = "Category"
    FrmMain.txtRecipe.DataField = "Name"
    FrmMain.txtAuthor.DataField = "Author"
    FrmMain.txtInstructions.DataField = "Instructions"
    
    Set rstRecipes.ActiveConnection = Nothing
    Set rstCategory.ActiveConnection = Nothing
    cnnRecipe.Close
    
    'Select Top Node
    FrmMain.Tv.Nodes(1).Selected = True

End Function

Public Sub AddCat(strCategory)
    'Function to add a Catagory to the DataBase
    On Error Resume Next
          
    Dim strMessage As String, strTitle As String, strSQLChange As String
    
    ' Show the Input Box for the Name of Catagory to add to DataBase
    If Trim(strCategory) <= "" Then
Again:
        strMessage = "Enter a Name For The New Catagory"
        strTitle = "Create New Catagory"
        strCategory = LCase(InputBox(strMessage, strTitle))
        If Len(strCategory) = 0 Then Exit Sub
        If Len(strCategory) > 75 Then
            MsgBox "You have entered in a value that is greater than 75 characters. " & _
                   Chr(13) & Chr(13) & "Please try again.", vbOKOnly + vbCritical _
                   , "Input Error"
            GoTo Again
        End If
    End If
    If Trim(strCategory) = "" Then Exit Sub
    
    ' Open connection.
    'Set cnnRecipe = New ADODB.Connection
    cnnRecipe.Open strCnn

    strSQLChange = "INSERT INTO Category ( Category )" & _
                   "SELECT '" & strCategory & "' AS Expr1;"

    ' Create command object.
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnRecipe
    cmdChange.CommandText = strSQLChange
    cmdChange.Execute
    
    If Err.Number <> 0 Then
        ' Notify User of Error
        MsgBox "Cannot Add Category [" & strCategory & "] Becuase it Already Exists." & _
               " Please Choose Another Category.", vbCritical + vbOKOnly, "New Category Error"
    Else
        intCatCount = intCatCount + 1
        FrmMain.Tv.Nodes.Add "BOOK", tvwChild, "Cat" & (intCatCount), _
                             strCategory, 2, 3
        FrmMain.Tv.Nodes("Cat" & intCatCount).Selected = True
        intTtlCategories = intTtlCategories + 1
    End If
    
    cnnRecipe.Close
    Set cmdChange = Nothing
End Sub

Public Function DelRec()
    'Function to Delete a Record from the DataBase
    'Just in case No Records are in Catagory Program want error out
    On Error Resume Next
    Dim I As Integer
    Dim intDelCount As Integer
    'For I = 1 To FrmMain.Tv.Nodes.Count
        ' Move recordset to delete position
    For I = FrmMain.Tv.Nodes.Count To 1 Step -1
        If FrmMain.Tv.Nodes.Item(I).Checked = True _
         And InStr(1, FrmMain.Tv.Nodes.Item(I).Key, "_") Then
            rstRecipes.AbsolutePosition = ((Replace(FrmMain.Tv.Nodes.Item(I).Key, "_", "")) + 1)
            rstRecipes.Delete
            'Delete Recipe Node from TreeView
            FrmMain.Tv.Nodes.Remove FrmMain.Tv.Nodes.Item(I).Key
            'Update Counters
            intTtlRecipeCount = intTtlRecipeCount - 1
            intCatRecipeCount = intCatRecipeCount - 1
            intDelCount = intDelCount + 1
        End If
    Next I
    
    MsgBox "You have successfully deleted " & intDelCount & _
           " recipes.", vbInformation + vbOKOnly, "Delete Confirmation"
    
End Function

Public Function DelCat()
    'Function to Delete a Catagory
    Dim strSQLChange As String
    Dim intIndex As Integer
    Dim Msg, Style, Title, Response
   
    'Make user has choosen a Catagory to Delete
    If Mid(FrmMain.Tv.SelectedItem.Key, 1, 3) <> "Cat" Then
       MsgBox "Please First Choose a Catagory to Delete"
       Exit Function
    End If

    'Warn User He/She is about to Delete the Catagory
    Msg = "This will Delete Catagory " & Chr$(34) & FrmMain.Tv.SelectedItem & Chr$(34) _
          & vbCrLf & "All of It's Recipes" & vbCrLf & "Do you want to continue ?"
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "Warnning"
    Response = MsgBox(Msg, Style, Title)
    
    If Response = vbNo Then Exit Function

    'Delete The Catagory (Table)
    ' Open connection.
    'Set cnnRecipe = New ADODB.Connection
    cnnRecipe.Open strCnn

    ' Process Delete from Recipes Table
    strSQLChange = "DELETE From Recipes Where Category = '" & FrmMain.Tv.SelectedItem & "'"
    Set cmdChange = New ADODB.Command
    Set cmdChange.ActiveConnection = cnnRecipe
    cmdChange.CommandText = strSQLChange
    cmdChange.Execute
    
    ' Process Delete from Category Table
    strSQLChange = "DELETE From Category Where Category = '" & FrmMain.Tv.SelectedItem & "'"
    cmdChange.CommandText = strSQLChange
    cmdChange.Execute
    
    cnnRecipe.Close
    Set cmdChange = Nothing
   
    'Delete Category Node from TreeView
    FrmMain.Tv.Nodes.Remove (FrmMain.Tv.SelectedItem.Index)
    intTtlCategories = intTtlCategories - 1

    'Reset Recordset Cursor to Beginning of File
    FrmMain.Tv.Nodes.Item("BOOK").Selected = True
    
End Function

Public Function DirMk()
    'Function to Create BackUp Directory
    On Error Resume Next

    'Make the Directory
    If Trim(Dir(App.Path & "\DataBackup")) <= "" Then MkDir App.Path & "\DataBackup"
End Function


Public Function CloseRs()
    'Function to Close RecordSets and DataBase
    On Error Resume Next
    
    ' In Case user has made an update movenext to commit update to recordset
    ' For some reason the Update command does not do this when the form is
    ' unloaded so I used Move 0 to get around this issue.
    If Not rstRecipes.EOF And Not rstRecipes.BOF Then
        rstRecipes.Move 0
    End If
    cnnRecipe.Open strCnn                           'reopen connection
    Set rstRecipes.ActiveConnection = cnnRecipe     'reconnect recordset
    rstRecipes.UpdateBatch                          'Update Recipes
    
    Set rstRecipes.ActiveConnection = Nothing
    cnnRecipe.Close

End Function

Public Function Normalize()
    'On Error Resume Next

    'Change Mouse Pointer to HourGlass
    FrmMain.MousePointer = 11
    
    'Set Virables
    Dim rs As Recordset
    Dim DB As Database
    Dim rsRecipes As Recordset
    Dim Table As TableDef
    Dim I As Integer
    

    'Open DataBase
    Set DB = OpenDatabase(App.Path & "\Recipe.mdb")

    With DB
        Set rsRecipes = DB.OpenRecordset("Recipes")
        'Move through the DataBase and get the Catagories (Tables)
        For Each Table In .TableDefs
            'We don't Want the Next Selven Ms Access Stuff in the Tables
            If Table.Name = "MSysACEs" Then GoTo TheNextTable
            If Table.Name = "MSysModules" Then GoTo TheNextTable
            If Table.Name = "MSysModules2" Then GoTo TheNextTable
            If Table.Name = "MSysObjects" Then GoTo TheNextTable
            If Table.Name = "MSysQueries" Then GoTo TheNextTable
            If Table.Name = "MSysRelationships" Then GoTo TheNextTable
            If Table.Name = "MSysAccessObjects" Then GoTo TheNextTable
            If Table.Name = "Recipes" Then GoTo TheNextTable
            If Table.Name = "Category" Then GoTo TheNextTable
            If Table.Name = "zzzNothing" Then GoTo TheNextTable
            'Open The RecordSet for each Child Node (Catagory)
            Set rs = DB.OpenRecordset(Table.Name)
            rs.MoveFirst
            'Move through the Records
            For I = 1 To rs.RecordCount
                rsRecipes.AddNew
                rsRecipes.Fields(0) = LCase(Table.Name)
                rsRecipes.Fields(1) = Proper(rs.Fields(0).Value)
                rsRecipes.Fields(2) = Proper(rs.Fields(1).Value)
                rsRecipes.Fields(3) = rs.Fields(2).Value & vbCrLf & _
                                      vbCrLf & vbCrLf & rs.Fields(3).Value
                'Move to the Next record
                rsRecipes.Update
                rs.MoveNext
            Next I
TheNextTable:
        Next Table
    End With
    
    'Call function to close the DataBase
    rs.Close
    rsRecipes.Close
    'Close DataBase
    DB.Close
    'Set Virables to Nothing
    Set rs = Nothing
    Set rsRecipes = Nothing
    Set DB = Nothing
    'Change Mouse Pointer to Standard
    FrmMain.MousePointer = 0
End Function

Public Sub CreateExportDB()
    'Function to Create a New DataBase

    'Change Mouse Pointer to HourGlass
    FrmMain.MousePointer = 11
    
    'Setup Virables
    Dim NewDB As Database
    Dim NewTable As TableDef
    Dim DBName As String
    Dim strReturn As String
    
    strReturn = MsgBox("This action will replace your last Export." _
                       , vbOKCancel + vbQuestion, "Backup Confirmation")
    If strReturn = "2" Then GoTo CloseRoutine
    
    'DataBase Name and Path
    DBName = App.Path + "\Export.mdb"
      
    'If Exist Delete It
    If Dir(DBName) <> "" Then
        Kill DBName
    End If
    
    'Create the New DataBase
    Set NewDB = CreateDatabase(DBName, dbLangGeneral)
                
    'Add the Catagories (Tables) to the Newly Created DataBase
    Set NewTable = NewDB.CreateTableDef("Export")
    With NewTable
        .Fields.Append .CreateField("Import", dbInteger)
        .Fields.Append .CreateField("Category", dbText, 50)
        .Fields.Append .CreateField("Name", dbText, 50)
        .Fields.Append .CreateField("Author", dbText, 50)
        .Fields.Append .CreateField("Instructions", dbMemo)
    End With
    
    ' Commit the append
    NewDB.TableDefs.Append NewTable

    'Close the Newly Created DataBase
    NewDB.Close
CloseRoutine:
    'Change Mouse Pointer to Standard
    FrmMain.MousePointer = 0
   
End Sub

Public Sub Search4Recipes(ByVal strConnectStr As String, ByVal strMessage As String _
                          , ByVal booCanSearch As Boolean, ByVal strCaption As String)
    'Show the Input Box for the Name of Catagory to Addto DataBase
    Dim strTitle As String
    Dim strMyValue As String
    Dim strSearch As String
    
    If booCanSearch = True Then
        strTitle = "Recipe Search"
        strSearch = InputBox(strMessage, strTitle)
        If strSearch = "" Then
            MsgBox "Nothing to search for", vbInformation, "Search Aborted"
            Exit Sub
        Else
            strSearch = Replace(strSearch, " ", "+")
        End If
    Else
        strSearch = ""
    End If
    
    'Call the Print Form
    FrmPrint.Show
    'Open the Htm file Created from the FrmMain
    FrmPrint.WB.Navigate strConnectStr & strSearch
    
    FrmPrint.txtAddress.Text = FrmPrint.WB.LocationURL
    FrmPrint.txtAddress.ToolTipText = FrmPrint.WB.LocationURL
    FrmPrint.Caption = strCaption
End Sub

Public Sub SearchEQ()
    Dim strConnectStr As String
    Dim strMessage As String
    
    strConnectStr = "http://www.epicurious.com/s97is.vts?action=" & _
                    "filtersearch&filter=recipe-filter.hts&collection" & _
                    "=Recipes&ResultTemplate=recipe-results.hts&queryType" & _
                    "=and&keyword="
                    
    strMessage = "Enter a food item to search Epicurious's recipe database for." & _
              Chr(13) & Chr(13) & Chr(13) & _
              "          ***** Please Note *****" & _
              Chr(13) & Chr(13) & "Separate you search items with one space only."
    
    Search4Recipes strConnectStr, strMessage, True, "Search Epicurious For Your Favorite Recipe"
End Sub

Public Sub SearchFoodTV()
    Search4Recipes "www.foodtv.com", "", False, "FoodTV Recipes Manual Search"
End Sub


Public Sub SearchHungryMonster()
    Search4Recipes "http://www.hungrymonster.com/recipe/recipe-search.cfm" _
                   , "", False, "HungryMonster Recipes Manual Search"
End Sub

Public Sub SearchQVC()
    Search4Recipes "http://www.qvc.com/asp/frameset.asp?mhproduct=mastheadp0200.gif&mhtitle=" & _
                   "mastheadt0200.gif&dd=/frames/drillframe0200.html&nest=/qvcrecip.html" _
                   , "", False, "QVC Recipes Manual Search"
End Sub

Public Sub SearchCopyKat()
    Search4Recipes "http://www.cdkitchen.com/search/allsearch.shtml", "", False, "QVC Recipes Manual Search"
End Sub

Public Sub SearchCDKitchen()
    Search4Recipes "http://www.copykat.com/asp/recipes.asp", "", False, "QVC Recipes Manual Search"
End Sub


