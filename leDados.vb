Sub LerDados()
   Dim L As Long, LG As Long, II As Long, R As Long, RF As Long
   Dim LI As String, LF As String
   '---
   Application.ScreenUpdating = False
   Application.Calculation = xlManual
   '--
   Call RemoveShets
   '--
   LG = Sheets("Documentos").Range("A1048576").End(xlUp).Row
      With Sheets("Documentos").Sort
           .SortFields.Clear
           .SortFields.Add Key:=Range("Q3:Q" & LG & ""), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
           .SortFields.Add Key:=Range("E3:E" & LG & ""), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
           .SetRange Rows("3:" & LG & "")
           .Header = xlGuess
           .MatchCase = False
           .Orientation = xlTopToBottom
           .SortMethod = xlPinYin
           .Apply
      End With
   
   R = 3: II = 3
   For L = 3 To LG
      LI = UCase(Sheets("Documentos").Cells(L, "Q"))
      If LI <> LF Then
         If LF <> "" Then
            Sheets("Documentos").Rows(II & ":" & (L - 1)).Copy
            Sheets(LF).Rows("3").PasteSpecial
            Application.CutCopyMode = False
            Sheets(LF).Cells.EntireColumn.AutoFit
         End If
         '--
         Call CreateSheet(LI)
         Sheets("Documentos").Rows("1:2").Copy
         Sheets(LI).Rows("1:2").PasteSpecial
         Application.CutCopyMode = False
         '--
         R = 3: II = L
         LF = LI
      End If
      '--
      R = R + 1
      '--
      Application.StatusBar = "Processando Registros.: " & L & " de " & LG & "  :  " & Format(L / LG, "Percent")
      DoEvents
   Next L
   '--
   If LF <> "" Then
      Sheets("Documentos").Rows(II & ":" & (L - 1)).Copy
      Sheets(LF).Rows("3:" & R).PasteSpecial
      Application.CutCopyMode = False
      Sheets(LF).Cells.EntireColumn.AutoFit
   End If
         '--
   Sheets("Documentos").Select
   Sheets("Documentos").Range("A2").Select
   '--
   
   Application.Calculation = xlAutomatic
   Application.ScreenUpdating = True
   Application.StatusBar = False



End Sub


Private Function CreateSheet(NameSheet As String)
   With ThisWorkbook
      .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = NameSheet
   End With
   '--
   Sheets(NameSheet).Select
   ActiveWindow.DisplayGridlines = False
   '--
   Sheets(NameSheet).Cells.Select
   With Selection
      .HorizontalAlignment = xlGeneral
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = False
      '-- Font
      .Font.Name = "Calibri"
      .Font.Size = 8
      .Font.Strikethrough = False
      .Font.Superscript = False
      .Font.Subscript = False
      .Font.OutlineFont = False
      .Font.Shadow = False
      .Font.Underline = xlUnderlineStyleNone
      .Font.Color = -16777216
      .Font.TintAndShade = 0
      .Font.ThemeFont = xlThemeFontNone
   End With
   Sheets(NameSheet).Cells(1, 1).Select
End Function

Private Function RemoveShets()
   Dim PF As Boolean, L As Long, SH()
   '--
   Application.DisplayAlerts = False
   '--
   ReDim SH(0)
   For Each asheet In ActiveWorkbook.Sheets
      PF = True
      If asheet.Name = "Documentos" Then PF = False
      If asheet.Name = "Resumo" Then PF = False
      '--
      If PF = True Then
         SH(UBound(SH)) = asheet.Name
         ReDim Preserve SH(UBound(SH) + 1)
      End If
   Next
   '--
   If UBound(SH) > 0 Then
      ReDim Preserve SH(UBound(SH) - 1)
      Sheets(SH).Delete
   End If
   ''--
   Application.DisplayAlerts = True
End Function

Sub CopySheets()
   Dim PF As Boolean, L As Long
   Dim nPlan As String, nnPlan As Variant
   '--
   Application.ScreenUpdating = False
   Application.Calculation = xlManual
   
   For Each asheet In ActiveWorkbook.Sheets
      PF = True
      If asheet.Name = "Documentos" Then PF = False
      If asheet.Name = "Resumo" Then PF = False
      '--
      If PF = True Then
         nPlan = nPlan & asheet.Name & "|"
      End If
   Next
   
   'MsgBox nPlan
   
   '--
   Application.Calculation = xlAutomatic
   nnPlan = Split(nPlan, "|")
   If UBound(nnPlan) > 0 Then
      For L = LBound(nnPlan) To (UBound(nnPlan) - 1)
         Application.StatusBar = "Arquivo criado.: " & CStr(nnPlan(L))
         Call exportCreateSheet(CStr(nnPlan(L)))
         'MsgBox CStr(nnPlan(L))
         DoEvents
      Next L
   End If
   '--
   
   Application.ScreenUpdating = True
   Application.StatusBar = False
   
End Sub

Private Function exportCreateSheet(nPlan As String)
   Dim wkbk As Workbook
   Application.DisplayAlerts = False
   Set wkbk = Application.Workbooks.Add
   With wkbk
      ThisWorkbook.Sheets(Array("Resumo", nPlan)).Copy After:=.Sheets(.Sheets.Count)
      Sheets(nPlan).Name = "NotasDeposito"
      '--
      .Sheets("Plan1").Delete
      '.Sheets("Plan2").Delete
     ' .Sheets("Plan3").Delete
      
      Call LerBase
      '.SaveAs Replace(ThisWorkbook.FullName, ThisWorkbook.Name, vbNullString) & nPlan & ".xlsx"
      .SaveAs Replace(ThisWorkbook.FullName, ThisWorkbook.Name, vbNullString) & PMaiuscula(nPlan) & " - Mercadorias em Depósito" & ".xlsx"
      .Close
   End With
   'zera as variáveis
   Set wkbk = Nothing
   Application.DisplayAlerts = True
   
  
End Function

Private Function PMaiuscula(sText As String) As String
    PMaiuscula = Application.Proper(sText)
End Function