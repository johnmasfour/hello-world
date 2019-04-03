VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClients 
   Caption         =   "SALES DETAILS INPUT"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16095
   OleObjectBlob   =   "frmClients.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddData_Click()

'Macro create form to input sales data onto Sales Transfers'

'Confirm Client Route entry'

Dim RowCount As Long
Dim benefits, total As Single
If Me.ListBoxRoute.Value = "" Then
MsgBox "Please enter Client Route", vbExclamation, "Client Route"
Me.ListBoxRoute.SetFocus
Exit Sub
End If

'Determine last empty row and check for Customer Routing

Dim lastrow
lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
lastrow = lastrow + 1
Cells(lastrow, 1) = TextBoxMonth

'Transfer information Week 1 & 2

Cells(lastrow, 1).Value = TextBoxMonth.Value
Cells(lastrow, 2).Value = ListBoxRoute.Value
Cells(lastrow, 3).Value = ListBoxCustomerNames.Value
Cells(lastrow, 4).Value = TextBox1001.Value
Cells(lastrow, 5).Value = TextBox1002.Value
Cells(lastrow, 6).Value = TextBox1003.Value
Cells(lastrow, 7).Value = TextBox1004.Value
Cells(lastrow, 8).Value = TextBox1005.Value
Cells(lastrow, 9).Value = TextBox1006.Value
Cells(lastrow, 10).Value = TextBox1007.Value
Cells(lastrow, 11).Value = TextBox1008.Value
Cells(lastrow, 12).Value = TextBox1009.Value
Cells(lastrow, 13).Value = TextBox1010.Value
Cells(lastrow, 14).Value = TextBox1011.Value
Cells(lastrow, 15).Value = TextBox1012.Value
Cells(lastrow, 16).Value = TextBox1013.Value
Cells(lastrow, 17).Value = TextBox1014.Value
Cells(lastrow, 18).Value = TextBox1015.Value
Cells(lastrow, 19).Value = TextBox1016.Value
Cells(lastrow, 20).Value = TextBox1017.Value
Cells(lastrow, 21).Value = TextBox1018.Value
Cells(lastrow, 22).Value = TextBox1019.Value
Cells(lastrow, 23).Value = TextBox1020.Value
Cells(lastrow, 24).Value = TextBox1021.Value
Cells(lastrow, 25).Value = TextBox1022.Value
Cells(lastrow, 26).Value = TextBox1023.Value
Cells(lastrow, 27).Value = TextBox1024.Value
Cells(lastrow, 28).Value = TextBox1025.Value
Cells(lastrow, 29).Value = TextBox1026.Value
Cells(lastrow, 30).Value = TextBox1027.Value
Cells(lastrow, 31).Value = TextBox1028.Value
Cells(lastrow, 32).Value = TextBox1029.Value
Cells(lastrow, 33).Value = TextBox1030.Value
Cells(lastrow, 34).Value = TextBox1031.Value
Cells(lastrow, 35).Value = TextBox1032.Value
Cells(lastrow, 36).Value = TextBox1033.Value
Cells(lastrow, 37).Value = TextBox1034.Value
Cells(lastrow, 38).Value = TextBox1035.Value
Cells(lastrow, 39).Value = TextBox1036.Value
Cells(lastrow, 40).Value = TextBox1037.Value
Cells(lastrow, 41).Value = TextBox1038.Value
Cells(lastrow, 42).Value = TextBox1039.Value
Cells(lastrow, 43).Value = TextBox1040.Value
Cells(lastrow, 44).Value = TextBox1041.Value
Cells(lastrow, 45).Value = TextBox1042.Value
Cells(lastrow, 46).Value = TextBox1043.Value
Cells(lastrow, 47).Value = TextBox1044.Value
Cells(lastrow, 48).Value = TextBox1045.Value
Cells(lastrow, 49).Value = TextBox1046.Value
Cells(lastrow, 50).Value = TextBox1047.Value
Cells(lastrow, 51).Value = TextBox1048.Value
Cells(lastrow, 52).Value = TextBox2001.Value
Cells(lastrow, 53).Value = TextBox2002.Value
Cells(lastrow, 54).Value = TextBox2003.Value
Cells(lastrow, 55).Value = TextBox2004.Value
Cells(lastrow, 56).Value = TextBox2005.Value
Cells(lastrow, 57).Value = TextBox2006.Value
Cells(lastrow, 58).Value = TextBox2007.Value
Cells(lastrow, 59).Value = TextBox2008.Value
Cells(lastrow, 60).Value = TextBox2009.Value
Cells(lastrow, 61).Value = TextBox2010.Value
Cells(lastrow, 62).Value = TextBox2011.Value
Cells(lastrow, 63).Value = TextBox2012.Value
Cells(lastrow, 64).Value = TextBox2013.Value
Cells(lastrow, 65).Value = TextBox2014.Value
Cells(lastrow, 66).Value = TextBox2015.Value
Cells(lastrow, 67).Value = TextBox2016.Value
Cells(lastrow, 68).Value = TextBox2017.Value
Cells(lastrow, 69).Value = TextBox2018.Value
Cells(lastrow, 70).Value = TextBox2019.Value
Cells(lastrow, 71).Value = TextBox2020.Value
Cells(lastrow, 72).Value = TextBox2021.Value
Cells(lastrow, 73).Value = TextBox2022.Value
Cells(lastrow, 74).Value = TextBox2023.Value
Cells(lastrow, 75).Value = TextBox2024.Value
Cells(lastrow, 76).Value = TextBox2025.Value
Cells(lastrow, 77).Value = TextBox2026.Value
Cells(lastrow, 78).Value = TextBox2027.Value
Cells(lastrow, 79).Value = TextBox2028.Value
Cells(lastrow, 80).Value = TextBox2029.Value
Cells(lastrow, 81).Value = TextBox2030.Value
Cells(lastrow, 82).Value = TextBox2031.Value
Cells(lastrow, 83).Value = TextBox2032.Value
Cells(lastrow, 84).Value = TextBox2033.Value
Cells(lastrow, 85).Value = TextBox2034.Value
Cells(lastrow, 86).Value = TextBox2035.Value
Cells(lastrow, 87).Value = TextBox2036.Value
Cells(lastrow, 88).Value = TextBox1037.Value
Cells(lastrow, 89).Value = TextBox2038.Value
Cells(lastrow, 90).Value = TextBox2039.Value
Cells(lastrow, 91).Value = TextBox2040.Value
Cells(lastrow, 92).Value = TextBox2041.Value
Cells(lastrow, 93).Value = TextBox2042.Value
Cells(lastrow, 94).Value = TextBox2043.Value
Cells(lastrow, 95).Value = TextBox2044.Value
Cells(lastrow, 96).Value = TextBox2045.Value
Cells(lastrow, 97).Value = TextBox2046.Value
Cells(lastrow, 98).Value = TextBox2047.Value
Cells(lastrow, 99).Value = TextBox2048.Value
Cells(lastrow, 100).Value = ListBoxStaff.Value
Cells(lastrow, 101).Value = TextBoxDateEntered.Value

'Clearinformation Week 1 & 2
  Me.TextBoxMonth.Value = ""
  Me.ListBoxRoute.Value = ""
  Me.ListBoxCustomerNames.Value = ""
  Me.TextBox1001.Value = ""
  Me.TextBox1002.Value = ""
  Me.TextBox1003.Value = ""
  Me.TextBox1004.Value = ""
  Me.TextBox1005.Value = ""
  Me.TextBox1006.Value = ""
  Me.TextBox1007.Value = ""
  Me.TextBox1008.Value = ""
  Me.TextBox1009.Value = ""
  Me.TextBox1010.Value = ""
  Me.TextBox1011.Value = ""
  Me.TextBox1012.Value = ""
  Me.TextBox1013.Value = ""
  Me.TextBox1014.Value = ""
  Me.TextBox1015.Value = ""
  Me.TextBox1016.Value = ""
  Me.TextBox1017.Value = ""
  Me.TextBox1018.Value = ""
  Me.TextBox1019.Value = ""
  Me.TextBox1020.Value = ""
  Me.TextBox1021.Value = ""
  Me.TextBox1022.Value = ""
  Me.TextBox1023.Value = ""
  Me.TextBox1024.Value = ""
  Me.TextBox1025.Value = ""
  Me.TextBox1026.Value = ""
  Me.TextBox1027.Value = ""
  Me.TextBox1028.Value = ""
  Me.TextBox1029.Value = ""
  Me.TextBox1030.Value = ""
  Me.TextBox1031.Value = ""
  Me.TextBox1032.Value = ""
  Me.TextBox1033.Value = ""
  Me.TextBox1034.Value = ""
  Me.TextBox1035.Value = ""
  Me.TextBox1036.Value = ""
  Me.TextBox1037.Value = ""
  Me.TextBox1038.Value = ""
  Me.TextBox1039.Value = ""
  Me.TextBox1040.Value = ""
  Me.TextBox1041.Value = ""
  Me.TextBox1042.Value = ""
  Me.TextBox1043.Value = ""
  Me.TextBox1044.Value = ""
  Me.TextBox1045.Value = ""
  Me.TextBox1046.Value = ""
  Me.TextBox1047.Value = ""
  Me.TextBox1048.Value = ""
  Me.TextBox2001.Value = ""
  Me.TextBox2002.Value = ""
  Me.TextBox2003.Value = ""
  Me.TextBox2004.Value = ""
  Me.TextBox2005.Value = ""
  Me.TextBox2006.Value = ""
  Me.TextBox2007.Value = ""
  Me.TextBox2008.Value = ""
  Me.TextBox2009.Value = ""
  Me.TextBox2010.Value = ""
  Me.TextBox2011.Value = ""
  Me.TextBox2012.Value = ""
  Me.TextBox2013.Value = ""
  Me.TextBox2014.Value = ""
  Me.TextBox2015.Value = ""
  Me.TextBox2016.Value = ""
  Me.TextBox2017.Value = ""
  Me.TextBox2018.Value = ""
  Me.TextBox2019.Value = ""
  Me.TextBox2020.Value = ""
  Me.TextBox2021.Value = ""
  Me.TextBox2022.Value = ""
  Me.TextBox2023.Value = ""
  Me.TextBox2024.Value = ""
  Me.TextBox2025.Value = ""
  Me.TextBox2026.Value = ""
  Me.TextBox2027.Value = ""
  Me.TextBox2028.Value = ""
  Me.TextBox2029.Value = ""
  Me.TextBox2030.Value = ""
  Me.TextBox2031.Value = ""
  Me.TextBox2032.Value = ""
  Me.TextBox2033.Value = ""
  Me.TextBox2034.Value = ""
  Me.TextBox2035.Value = ""
  Me.TextBox2036.Value = ""
  Me.TextBox1037.Value = ""
  Me.TextBox2038.Value = ""
  Me.TextBox2039.Value = ""
  Me.TextBox2040.Value = ""
  Me.TextBox2041.Value = ""
  Me.TextBox2042.Value = ""
  Me.TextBox2043.Value = ""
  Me.TextBox2044.Value = ""
  Me.TextBox2045.Value = ""
  Me.TextBox2046.Value = ""
  Me.TextBox2047.Value = ""
  Me.TextBox2048.Value = ""
  Me.ListBoxStaff.Value = ""
  Me.TextBoxDateEntered.Value = ""
  Me.TextBoxMonth.SetFocus

End Sub

Private Sub cmdClose_Click()
'Close the form
  Unload Me
End Sub

Private Sub CommandButton1_Click()
    Me.lstProduct.AddItem cboProduct.Value
    Me.lstWeek12.AddItem txtWeek12.Value
    Me.lstWeek34.AddItem txtWeek34.Value
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then
    Cancel = True
    MsgBox "Please use the Close Form button!"
  End If
End Sub

