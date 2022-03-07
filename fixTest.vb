Dim oFileName As String = "2999.00.00_test_druga_trzecia"
		'Dim oFileName As String = "2999.00.00_testdrugatrzecia"
    Dim NazwaDzielona = Split(oFileName, "_")
    Dim iloscNazw As Integer
    Dim resztaNazwy as String = NazwaDzielona(1)
    Dim index As Integer = 2
    
    iloscNazw = UBound(NazwaDzielona)

    If iloscNazw > 2 Then
      
      Do
          resztaNazwy = resztaNazwy + "_" + NazwaDzielona(index)        
          index += 1
      Loop Until index > iloscNazw
      

      iProperties.Value("Project", "Part Number") = NazwaDzielona(0)
      iProperties.Value("Project", "Description") = resztaNazwy
      iProperties.Value("Summary", "Title") = resztaNazwy
      
Else
      
      
      iProperties.Value("Project", "Part Number") = NazwaDzielona(0)
      iProperties.Value("Project", "Description") = NazwaDzielona(1)
      iProperties.Value("Summary", "Title") = NazwaDzielona(1)
    
End If