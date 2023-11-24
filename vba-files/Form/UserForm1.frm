VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6552
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9312.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EnableEvents As Boolean
' delete process
Sub delete()
        TextBox1.Value = ""
        TextBox2.Value = ""
        OptionButton1.Value = False
        OptionButton2.Value = False
        CheckBox1.Value = False
        CheckBox2.Value = False
        CheckBox3.Value = False
        ListBox1.Value = ""
        
        'delete department data
        ComboBox1.Clear
        ComboBox1.AddItem "HR"
        ComboBox1.AddItem "IT"
        ComboBox1.AddItem "MARKETING"
        
        'add searchdata process
        Call searchdata
        Worksheets("data").AutoFilterMode = False
        Worksheets("search").AutoFilterMode = False
        Worksheets("search").Cells.Clear
        
        
        'input data process
        irow = Worksheets("data").Cells(Rows.Count, 1).End(xlUp).Row 'return max row
        
        With ListBox2
                .ColumnCount = 9                    'there has 9 columns
                .ColumnHeads = False           'title will show on second row
                .ColumnWidths = "40, 60, 60, 50, 60, 60, 60, 60"
                
                If irow > 2 Then
                      .RowSource = "data!A2:I" & irow  'display all data
                      
                Else
                      .RowSource = "data!A2:I2"           'display title
                End If
         End With
         
                      
                
        
        
End Sub

Private Sub ComboBox2_Change()
        If Me.EnableEvents = False Then Exit Sub
        If Me.ComboBox2.Value = "all" Then
                Call delete
        Else
                Me.TextBox3.Value = ""
                Me.TextBox3.Enabled = True
                Me.CommandButton3.Enabled = True
        End If

End Sub

Private Sub CommandButton1_Click()
        'connect with input data button(with notice window)
        Dim msgValue As VbMsgBoxResult
        msgValue = MsgBox("Do you want to input data?", vbYesNo + vbInformation, "Yes")
        If msgValue = vbNo Then Exit Sub
             Call submit
             Call delete
             
End Sub

Private Sub CommandButton2_Click()
        'connect with delete button(with notice window)
        Dim msgValue As VbMsgBoxResult
        msgValue = MsgBox("Do you want to delete data?", vbYesNo + vbInformation, "Yes")
        If msgValue = vbNo Then Exit Sub
             
             Call delete

End Sub

Private Sub CommandButton3_Click()
        If Me.TextBox3.Value = "" Then              ' Me is userform
             msgValue = MsgBox("Please input your data", vbOKOnly + vbInformation, "search")
              Exit Sub
        End If
        Call search

End Sub

Private Sub CommandButton4_Click()
'Edit data
        If select_row = 0 Then
            MsgBox "Sorry, no data selected", vbOKOnly + vbInformation, "Yes"
            Exit Sub
        End If
        Call Edit
End Sub

Private Sub CommandButton5_Click()
'remove data
    Dim irow As Long
    Dim i As VbMsgBoxResult
    
    If select_row = 0 Then
            MsgBox "no data selected", vbOKOnly + vbInformation, "remove"
    End If
    
    irow = select_row + 1
    i = MsgBox("Do you want to remove the data?", vbYesNo + vbInformation, "remove")
    If i = vbNo Then Exit Sub
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("data").Rows(irow).delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Call delete
    MsgBox "Your data has been removed!", vbOKOnly + vbInformation, "remove"
    
    
    
End Sub

Private Sub UserForm_Initialize()
         Dim i As Integer
         irow = Worksheets("state").Cells(Rows.Count, 1).End(xlUp).Row
         
         For i = 1 To irow
                ListBox1.AddItem Worksheets("state").Cells(i, 1).Value
         Next
         Call delete

End Sub


'----------------------------------------------------------------------------------------------------------------------------------------------------------
'search data process
Sub searchdata()
        EnableEvents = False   ' only run once
        
        With UserForm1.ComboBox2
                 .Clear
                 .AddItem "All"
                 .AddItem "ID"
                 .AddItem "name"
                 .AddItem "gender"
                 .AddItem "department"
                 .AddItem "state"
                 
                 .Value = "All"    '  display all data as defualt
         End With
         
         EnableEvents = True
         TextBox3.Value = ""
         TextBox3.Enabled = False
         CommandButton3.Enabled = False
        
End Sub























