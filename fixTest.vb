'Dim oFileName As String = ThisDoc.FileName(False)
Dim oFileName As String
Dim NazwaDzielona = Split(oFileName, "_")
Dim Test, Test2 As Integer

oFileName = "2999.00.00_test_druga_trzecia"
Test = UBound(NazwaDzielona)
Test2 = LBound(NazwaDzielona)

MsgBox(Test)
'MsgBox(Test2)

'If Test > 1 Then
	
	
'Else	

MsgBox(NazwaDzielona(0) & vbCrLf & NazwaDzielona(1) & vbCrLf & NazwaDzielona(2) & vbCrLf & NazwaDzielona(3))

Test = UBound(NazwaDzielona)
'Test2 = LBound(NazwaDzielona)

MsgBox(Test)
'MsgBox(Test2)


'iProperties.Value("Project", "Part Number") = NazwaDzielona(0)
'iProperties.Value("Project", "Description") = NazwaDzielona(1)
'iProperties.Value("Summary", "Title") = NazwaDzielona(1)


