Dim chChart             As Chart
    Dim sFileLocation       As String
    Dim sFileName           As String
    Dim dteStart As Date, dteFinish As Date
    Dim dteStopped As Date, dteElapsed As Date
    Dim boolStopPressed As Boolean, boolResetPressed As Boolean

Private Sub btnreset_Click()
dteStopped = 0
    dteStart = 0
    dteElapsed = 0
    Label1 = "00:00:00"
    boolResetPressed = True
End Sub

Private Sub Btnstart_Click()
UserForm1.MultiPage1.Value = 1
MultiPage1.Visible = True
Image1.Picture = LoadPicture(ThisWorkbook.Path & "\image\0.gif")
UserForm1.TextBox2 = 0
Start_timer:
    dteStart = Time
    boolStopPressed = False
    boolResetPressed = False
Timer_Loop:
    DoEvents
    dteFinish = Time
    dteElapsed = dteFinish - dteStart + dteStopped
    lama = Minute(dteElapsed) * 60 + Second(dteElapsed)
    If Not boolStopPressed = True Then
        Label1.Caption = Format(lama, "##")
           UserForm1.TextBox1 = lama
        If boolResetPressed = True Then GoTo Start_timer
        GoTo Timer_Loop
    Else
        Exit Sub
    End If
End Sub

Private Sub Btnstop_Click()
 boolStopPressed = True
    dteStopped = dteElapsed
End Sub

    Private Sub CommandButton1_Click()
UserForm1.MultiPage1.Value = 0
MultiPage1.Visible = True
UserForm1.ListBox1.ColumnWidths = "100,55,55,55,55,55,55,55,55,360"
UserForm1.ListBox2.ColumnWidths = "90,60,60,60,60,60,60,60,60"
UserForm1.ListBox3.ColumnWidths = "90,60,60,60,60,60,60,60,60"
End Sub

Private Sub CommandButton2_Click()

For i = 0 To 13

UserForm1.TextBox1.Text = awal
If Val(awal / 5) = 0 Then
UserForm1.TextBox2 = awal
Loop
 
  
End Sub

Private Sub CommandButton3_Click()
Dim chChart             As Chart
    Dim sFileLocation       As String
    Dim sFileName           As String
    
    On Error Resume Next
    UserForm1.MultiPage1.Value = 1
    MultiPage1.Visible = True
    
    
    sFileName = "target"
    Set chChart = Sheet7.ChartObjects(1).Chart
    sFileLocation = ThisWorkbook.Path & "\" & sFileName & ".gif"
    chChart.Export Filename:=sFileLocation, FilterName:="GIF"
    Image1.Picture = LoadPicture(ThisWorkbook.Path & "\" & sFileName & ".gif")
    On Error GoTo 0
End Sub

Private Sub CommandButton4_Click()
UserForm1.Hide
 Sheet5.Activate
End Sub

Private Sub CommandButton5_Click()
    Dim chChart             As Chart
    Dim sFileLocation       As String
    Dim sFileName           As String
    On Error Resume Next
    For i = 1 To 10
    sFileName = i + 2
    Set chChart = Sheet7.ChartObjects("chart " & i).Chart
    sFileLocation = ThisWorkbook.Path & "\image\" & sFileName & ".gif"
    chChart.Export Filename:=sFileLocation
    ', FilterName:="GIF"
    Next i
   ' On Error GoTo 0
   Unload Me
End Sub

Private Sub CommandButton6_Click()
UserForm1.Hide
 
Sheet1.Activate

End Sub

Private Sub CommandButton7_Click()
Unload Me
End Sub


Private Sub CommandButton9_Click()
sfile = ThisWorkbook.Path & "\grade2021.xlsm"
Workbooks.Open Filename:=sfile
Me.Hide
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Image2_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()
If Val(TextBox2) > 13 Then
TextBox2 = 0
End If
End Sub

Private Sub TextBox1_Change()
k = Val(UserForm1.TextBox1) Mod 10
If k = 0 And Val(TextBox1) > 0 Then
TextBox2 = Val(TextBox2) + 1
h = Val(TextBox2)
UserForm1.Image1.Picture = LoadPicture(ThisWorkbook.Path & "\image\" & h & ".gif")
End If
j = Val(TextBox1.Text)
If j Mod 2 = 0 Then
UserForm1.TextBox24.ForeColor = vbRed
End If
If j Mod 2 <> 0 Then
UserForm1.TextBox24.ForeColor = vbBlack
End If




End Sub

Private Sub TextBox10_Change()

End Sub

Private Sub TextBox13_Change()

End Sub

Private Sub TextBox14_Change()

End Sub

Private Sub TextBox2_Change()
If Val(TextBox2) > 9 Then
TextBox2 = 0
End If
If Val(TextBox2) Mod 5 <> 0 Then
UserForm1.Image3.Picture = LoadPicture(ThisWorkbook.Path & "\image\11.gif")
Else
    UserForm1.Image3.Picture = LoadPicture(ThisWorkbook.Path & "\image\12.gif")
End If
End Sub

Private Sub TextBox22_Change()

End Sub

Private Sub TextBox24_Change()

End Sub

Private Sub TextBox24_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub TextBox29_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Activate()
Application.WindowState = xlMaximized
    'Me.Left = Application.Width - Me.Width
    Me.Top = 0
    
    
    
    
With Application
.WindowState = xlMaximized
Zoom = Int(.Width / Me.Width * 100)
Width = .Width
Height = .Height
End With
    
    



tg1 = Format(Sheet9.Cells(4, 9), "dd")
bl1 = Sheet9.Cells(4, 10)

hr = 0
lbr = 0
        
    For i = 3 To 300
    bl = Val(Sheet9.Cells(i, 2))
    tg = Format(Sheet9.Cells(i, 1), "dd")
    hri = Sheet9.Cells(i, 6)
   
    
    If bl = bl1 And tg >= tg1 Then
        hr = hr + hri
             If Sheet9.Cells(i, 5) = 0 Then
                lbr = lbr + 1
                End If
       
    End If
    Next i
    Sheet9.Cells(4, 14) = hr - lbr
    'Sheet9.Cells(4, 13) = lbr
    
    
UserForm1.TextBox5.Text = Format(Now(), "mm")
UserForm1.TextBox10.Text = Sheet9.Cells(4, 14) & "  Dari " & Sheet9.Cells(4, 12)
UserForm1.TextBox6.Text = Format(Sheet9.Cells(6, 9), "#.0")
UserForm1.TextBox7.Text = Format(Sheet9.Cells(6, 10), "#.0")
UserForm1.TextBox8.Text = Format(Sheet9.Cells(6, 11), "#.0")
UserForm1.TextBox9.Text = Format(Sheet9.Cells(6, 12), "#.0")
UserForm1.TextBox11.Text = Format(Sheet9.Cells(6, 13), "#.0")

UserForm1.TextBox17.Text = Format(Sheet9.Cells(7, 9), "#.0")
UserForm1.TextBox18.Text = Format(Sheet9.Cells(7, 10), "#.0")
UserForm1.TextBox19.Text = Format(Sheet9.Cells(7, 11), "#.0")
UserForm1.TextBox20.Text = Format(Sheet9.Cells(7, 12), "#.0")
UserForm1.TextBox21.Text = Format(Sheet9.Cells(7, 13), "#.0")
UserForm1.TextBox23.Text = Format(Sheet9.Cells(8, 9), "#.0")
UserForm1.TextBox25.Text = Format(Sheet9.Cells(8, 10), "#.0")
UserForm1.TextBox26.Text = Format(Sheet9.Cells(8, 11), "#.0")
UserForm1.TextBox27.Text = Format(Sheet9.Cells(8, 12), "#.0")
UserForm1.TextBox28.Text = Format(Sheet9.Cells(8, 13), "#.0")




UserForm1.Image1.Picture = LoadPicture(ThisWorkbook.Path & "\image\0.gif")
UserForm1.Image2.Picture = LoadPicture(ThisWorkbook.Path & "\image\9.gif")
UserForm1.Image3.Picture = LoadPicture(ThisWorkbook.Path & "\image\11.gif")
'====

For i = 6 To 330
Data = Sheet5.Cells(i, 4)
Data2 = Sheet5.Cells(i + 1, 4)
Data3 = Sheet5.Cells(i + 2, 4)
lbr = Sheet5.Cells(i, 18)
If Data = 0 And Data2 = 0 And Data3 = 0 Then
tgl = Sheet5.Cells(i - 1, 1)
UserForm1.TextBox24 = "--LAST UPDATE!!--  " & Format(tgl, "dd-mmm")
Exit For
End If
Next i



UserForm1.MultiPage1.Value = 1
'UserForm1.MultiPage1.Pages(0).Visible = True



End Sub

Private Sub UserForm_Click()

End Sub


