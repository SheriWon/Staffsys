Attribute VB_Name = "Module1"
'create submit process
Sub submit()
          
          Dim i As Integer
          
          'if no data selected
          If select_row = 0 Then
                lastrow = Worksheets("data").Cells(Rows.Count, 1).End(xlUp).Row + 1
          Else
                lastrow = select_row + 1  'new row number will be select row +1
          End If
          
            'table format
          With Range(Cells(lastrow, 1), Cells(lastrow, 9)).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
           End With
           
           'using with clause to insert data
           With Worksheets("data")
                        .Cells(lastrow, 1).Value = lastrow - 2
                        .Cells(lastrow, 2).Value = UserForm1.TextBox1.Text  'id B
                        .Cells(lastrow, 3).Value = UserForm1.TextBox2.Text  'name C
                        
                        'Gender D
                        If UserForm1.OptionButton1.Value = True Then
                                .Cells(lastrow, 4).Value = "M"
                        End If
                        
                        If UserForm1.OptionButton2.Value = True Then
                                .Cells(lastrow, 4).Value = "F"
                        End If
                        
                        .Cells(lastrow, 5).Value = UserForm1.ComboBox1.Value   ' department E
                        
                        'Hobby E-H
                        If UserForm1.CheckBox1.Value = True Then
                                .Cells(lastrow, 6).Value = "Yes"
                        End If
                        
                         If UserForm1.CheckBox2.Value = True Then
                                .Cells(lastrow, 7).Value = "Yes"
                        End If
                        
                        If UserForm1.CheckBox3.Value = True Then
                                .Cells(lastrow, 8).Value = "Yes"
                        End If
                        
                        .Cells(lastrow, 9).Value = UserForm1.ListBox1.Value         'state I
                End With
                
            
            
End Sub

'search
Sub search()
    Application.ScreenUpdating = False 'cancel screen update
    Dim shData As Worksheet
    Dim shSearch As Worksheet
    Dim iColumn As Integer
    Dim iDataRow As Long
    Dim iSearchRow As Long
    Dim sColumn As String
    Dim sValue As String
    
    Set shData = Worksheets("data")
    Set shSearch = Worksheets("search")
    
    iDataRow = shData.Cells(Rows.Count, 1).End(xlUp).Row
    sColumn = UserForm1.ComboBox2.Value         'method
    sValue = UserForm1.TextBox3.Value                  'category
    iColumn = Application.WorksheetFunction.Match(sColumn, shData.Range("A2:I2"), 0) 'return result
    
    If shData.FilterMode = True Then         'no filter
         shData.AutoFilterMode = False
    End If
    
    ' add filter
    If UserForm1.ComboBox2.Value = "ID" Then
         shData.Range("A2:I" & iDataRow).AutoFilter field:=iColumn, Criteria1:=sValue
    Else
         shData.Range("A2:I" & iDataRow).AutoFilter field:=iColumn, Criteria1:="*" & sValue & "*"
    End If
     
    ' start search
    If Application.WorksheetFunction.Subtotal(3, shData.Range("C:C")) >= 2 Then
        shSearch.Cells.Clear
        ' Copy the column titles to the first row of the "search" worksheet
        shData.Range("A1:I1").Copy shSearch.Cells(1, 1)
        ' Copy the filtered data (including headers) to the "search" worksheet starting from row 2
        shData.AutoFilter.Range.Copy shSearch.Cells(2, 1)
        Application.CutCopyMode = False
        
        iSearchRow = shSearch.Cells(Rows.Count, 1).End(xlUp).Row
        UserForm1.ListBox2.ColumnCount = 9
        UserForm1.ListBox2.ColumnWidths = "40, 60, 60, 50, 60, 60, 60, 60"
        If iSearchRow > 1 Then
            UserForm1.ListBox2.RowSource = "search!A2:I" & iSearchRow
            MsgBox "Find result"
        End If
    Else
        MsgBox "Sorry, no result"
    End If
    
    shData.AutoFilterMode = False
    Application.ScreenUpdating = True
End Sub


'-------------------------------------------------------------------------------------------------------------------------------------
'define function to return select row
Function select_row() As Long
                Dim i As Integer
                select_row = 0
                ' for loop from fiirst row to last row
                For i = 0 To UserForm1.ListBox2.ListCount - 1  'the last row number is listcount - 1
                    If UserForm1.ListBox2.Selected(i) = True Then   ' select row you want to edit
                          select_row = i + 1
                          Exit For
                    End If
                 Next
                 
End Function

'Edit process you want to provide additional details about the selected item from a ListBox to the user by displaying those details in a TextBox.
Sub Edit()
        Dim gender As String
        Dim us1 As UserForm1
        
        
        Set us1 = UserForm1   'an object refers to the UderForm
        
        'us1 is likely the name of the UserForm or an object that refers to the UserForm where the TextBox (TextBox1) and ListBox (ListBox2) controls are located.
        'List is a property of the ListBox control that returns a two-dimensional array representing the data displayed in the ListBox.
        'us1.ListBox2.ListIndex is the index of the selected item in the ListBox. ListIndex is a property of the ListBox control that represents the currently selected item's index. It will be -1 if no item is selected.
        
        us1.TextBox1.Value = us1.ListBox2.List(us1.ListBox2.ListIndex, 1)   'ID (second column is 1 first column id 0)
        us1.TextBox2.Value = us1.ListBox2.List(us1.ListBox2.ListIndex, 2)   'name
        
        gender = us1.ListBox2.List(us1.ListBox2.ListIndex, 3) 'gender
        If gender = "M" Then
                us1.OptionButton1.Value = True
        Else
                us1.OptionButton2.Value = True
        End If
        
        us1.ComboBox1.Value = us1.ListBox2.List(us1.ListBox2.ListIndex, 4) 'department
        us1.CheckBox1.Value = us1.ListBox2.List(us1.ListBox2.ListIndex, 5) 'sport
        us1.CheckBox2.Value = us1.ListBox2.List(us1.ListBox2.ListIndex, 6) 'singing
        us1.CheckBox3.Value = us1.ListBox2.List(us1.ListBox2.ListIndex, 7) 'reading
        us1.ListBox1.Value = us1.ListBox2.List(us1.ListBox2.ListIndex, 8) 'state
        
        MsgBox "Start to edit and input data?", vbOKOnly + vbInformation, "Edit"
        
        
        
End Sub



















