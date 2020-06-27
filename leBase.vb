Function LerBase()
    Dim L As Long, LF As Long
    Dim AA As String, BB As String
    Dim sT As String, sB As String
    
    Dim LI As Long, LC As Long
    
    LG = Sheets("NotasDeposito").Range("A1048576").End(xlUp).Row
    
    LI = 18
    LC = 1
    For L = 3 To LG
      'If Sheets("NotasDeposito").Cells(L, "Z") > 0 Then
         'AA = Sheets("NotasDeposito").Cells(L, "X")
         
         If Sheets("NotasDeposito").Cells(L, "AA") > 0 Then
         AA = Sheets("NotasDeposito").Cells(L, "Y")
         
         

         'If Sheets("NotasDeposito").Cells(L, "P") = "NAO" Then
         '   If InStr(sB, AA) = 0 Then
         '      sB = sB & AA & ","
         '   End If
         'End If
         
         
         If Sheets("NotasDeposito").Cells(L, "P") = "NAO" Then
            If InStr(sT, AA) = 0 Then
               If LC <= 3 Then
                  sT = sT & AA
                  'Sheets("Resumo").Cells(LI, "B") = Sheets("NotasDeposito").Cells(L, "X")
                    Sheets("Resumo").Cells(LI, "B") = Sheets("NotasDeposito").Cells(L, "Y")
               End If
               '--
               LI = LI + 1
               LC = LC + 1
            End If
         End If
         
        End If
    Next L
    
    
    
    'MsgBox sB
    'MsgBox Split(sB, ",")(0)
   'Sheets("Resumo").Cells(18, "B") = isP[0]
   'Sheets("Resumo").Cells(19, "B") = iSP(1)
   'Sheets("Resumo").Cells(20, "B") = iSP(2)
End Function