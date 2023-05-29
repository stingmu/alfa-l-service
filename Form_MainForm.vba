Option Compare Database

Private Sub AssortmentAddButton_Click()
 
 DoCmd.OpenForm "AddProductToAssortmentForm"
 
End Sub

Private Sub AssortmentChangeButton_Click()

 DoCmd.OpenForm "ChangeProductInAssortmentForm"
 
End Sub

Private Sub AssortmentClearAllButton_Click()

 Dim AssortmentListControl As ListBox
 
 Dim Response As Integer
 
 Set AssortmentListControl = Forms("MainForm").Controls("TabControlMainForm").Pages("AssortmentTab").Controls("AssortmentproductList")
 
 Set dbsCurrent = CurrentDb
 Set qdfQuery = dbsCurrent.CreateQueryDef("")
 
 Response = vbCancel
 
 If (AssortmentListControl.ListCount <> 0) Then
  
  Response = MsgBox("Óäàëèòü âñå ïðîäóêòû èç àññîðòèìåíòà?", vbOKCancel)
  
  If (Response = vbOK) Then
    
   qdfQuery.SQL = "DELETE FROM AssortmentProductTable;"
   qdfQuery.Execute
      
   ResetCalculationResult
   PagesVisiability
    
   ControlView_AssortmentPage
   
   If (AssortmentListControl.ListCount <> 0) Then

    MsgBox "Îïåðàöèÿ î÷èñòêè àññîðòèìåíòà âûïîëíåíà ñ îøèáêîé."
 
   End If
   
  Else
  
   MsgBox "Îïåðàöèÿ î÷èñòêè àññîðòèìåíòà ïðåðâàíà ïîëüçîâàòåëåì."
  
  End If
  
 Else
 
  MsgBox "Â àññîðòèìåíòå íåò ïðîäóêòîâ. Îïåðàöèÿ î÷èñòêè àññîðòèìåíòà ïðåðâàíà."
  
 End If
 
 qdfQuery.Close
 dbsCurrent.Close
  
End Sub

Private Sub AssortmentDeleteButton_Click()
  
 Dim AssortmentListControl As ListBox
 
 Dim TempInt As Integer
 Dim Response As Integer
 
 Set AssortmentListControl = Forms("MainForm").Controls("TabControlMainForm").Pages("AssortmentTab").Controls("AssortmentproductList")
     
 Set dbsCurrent = CurrentDb
 Set qdfQuery = dbsCurrent.CreateQueryDef("")
 
 Response = vbCancel
 
 If (AssortmentListControl.ListCount <> 0) Then
 
  If (AssortmentListControl.ListIndex > -1) Then
  
   Response = MsgBox("Óäàëèòü ïðîäóêò èç àññîðòèìåíòà?", vbOKCancel)
   
   If (Response = vbOK) Then
    
    TempInt = AssortmentListControl.ListCount
    
    qdfQuery.SQL = "DELETE FROM AssortmentProductTable WHERE AssortmentProductID IN (SELECT " & _
                   "SystemProductListTable.SystemProductID FROM SystemProductListTable WHERE " & _
                   "SystemProductListTable.SystemProductName = '" & _
                    AssortmentListControl.Column(0, AssortmentListControl.ListIndex) & "');"
    
    qdfQuery.Execute
    
    ResetCalculationResult
    PagesVisiability
    
    ControlView_AssortmentPage
    
    If ((TempInt - AssortmentListControl.ListCount) <> 1) Then
    
     MsgBox "Ïðè óäàëåíèè ïðîäóêòà èç àññîðòèìåíòà ïðîèçîøëà îøèáêà."
    
    End If
     
   Else
   
    MsgBox "Îïåðàöèÿ óäàëåíèÿ ïðåðâàíà ïîëüçîâàòåëåì."
   
   End If
   
  Else
  
   MsgBox "Íå âûáðàí ïðîäóêò äëÿ óäàëåíèÿ. Îïåàðöèÿ óäàëåíèÿ ïðåðâàíà."
  
  End If
  
 Else
 
  MsgBox "Àññîðòèìåíò ïóñò. Îïåðàöèÿ óäàëåíèÿ ïðåðâàíà."
 
 End If
 
 qdfQuery.Close
 dbsCurrent.Close
   
End Sub

Private Sub AssortmentproductList_Click()
 
 ControlView_AssortmentPage
  
End Sub

Private Sub ButterMilkFatField_AfterUpdate()

 Dim ButterMilkFat As TextBox
 
 Set ButterMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("ButterMilkFatField")
 
 If (ButterMilkFatAcc <> ButterMilkFat.Value) Then
 
  ButterMilkFatAcc = ButterMilkFat.Value
  
  ResetCalculationResult
  
 End If
 
 ButterMilkFat.SetFocus
  
 PagesVisiability
 ControlView_SourceBalancePage

End Sub

Private Sub CalculateButton_Click()
 
 Dim CalculationButton As CommandButton
 Dim ReportButton As CommandButton
 
 Dim RawMilkQuantity As TextBox
 Dim RawMilkFat As TextBox
 Dim CreamFat As TextBox
 Dim FreeFatMilkFat As TextBox
 Dim ButterMilkFat As TextBox
 
 Dim RawMilkNeededQuantity As TextBox
 Dim RawMilkQuantityDifference As TextBox
 Dim SeparatedCreamQuantity As TextBox
 Dim SeparatedCreamQuantityDifference As TextBox
 Dim FreeFatMilkQuantity As TextBox
 Dim FreeFatMilkQuantityDifference As TextBox
 Dim CottageCheeseWhey As TextBox
 Dim CheeseWhey As TextBox
 Dim ButterMilk As TextBox
 
 Dim Index As Integer
 Dim EndIndex As Integer
  
 Dim LowFatMilkVolumeForLowFatProducts As Double
 Dim LowFatMilkVolumeForMeduimFatProducts As Double
 Dim LowFatMilkVolumeForAllProducts As Double
 
 Dim LowFatMilkVolumeForMeduimFatProductsIteration As Double
 
 Dim DeltaLowFatMilkVolumeForLowFatProducts As Double
 
 Dim CreamWeightIteration As Double
 Dim RawMilkWeightIteration As Double
 
 Dim CreamWeightFromLowFatProducts As Double
 Dim RawMilkWeightForLowFatProducts As Double
 Dim ButterMilkWeightIteration As Double
 
 Dim CreamWeightForMediumFatProducts As Double
 Dim RawMilkWeightForMediumFatProducts As Double
 
 Dim CreamWeightForHiFatProducts As Double
 Dim ButterMilkWeightFromHiFatProducts As Double
 Dim DeltaRawMilkNeedForLowFatProducts As Double
 Dim DeltaCreamFromLowFatProducts As Double
 
 Dim RawMilkNeedWeight As Double
 Dim RawMilkAdditionalSeparationNeedWeight As Double
 Dim CreamOutputWeight As Double
 Dim CreamCurrentWeight As Double
 Dim CreamAdditionalSeparationWeight As Double
 Dim FreeFatMilkAdditionalSeparationWeight As Double
 Dim FreeFatMilkCurrentWeight As Double
 Dim FreeFatMilkOutputWeight As Double
 Dim CottageCheeseWheyWeight As Double
 Dim CheeseWheyWeight As Double
 
 Dim AssortmentListControl As ListBox
 
 Dim CurrentFat As Single
 Dim CurrentWeight As Double
 Dim CurrentProductType As Long
 
 Set CalculationButton = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("CalculateButton")
 Set ReportButton = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("ReportButton")
 
 Set RawMilkQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkQuantityField")
 Set RawMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkFatField")
 Set CreamFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("CreamFatField")
 Set FreeFatMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("FreeFatMilkFatField")
 Set ButterMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("ButterMilkFatField")
 
 Set RawMilkNeededQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkNeededQuantityField")
 Set RawMilkQuantityDifference = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkQuantityDifferenceField")
 Set SeparatedCreamQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("SeparatedCreamQuantityField")
 Set SeparatedCreamQuantityDifference = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("SeparatedCreamQuantityDifferenceField")
 Set FreeFatMilkQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("FreeFatMilkQuantityField")
 Set FreeFatMilkQuantityDifference = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("FreeFatMilkQuantityDifferenceField")
 Set CottageCheeseWhey = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("CottageCheeseWheyField")
 Set CheeseWhey = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("CheeseWheyField")
 Set ButterMilk = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("ButterMilkField")
 
 Set AssortmentListControl = Forms("MainForm").Controls("TabControlMainForm").Pages("AssortmentTab").Controls("AssortmentproductList")
 
 ResetCalculationResult
 
 RawMilkQuantity.SetFocus
 
 CalculationButton.Enabled = False
 ReportButton.Enabled = False
 
 If ((RawMilkQuantity.Value = 0) Or (RawMilkFat.Value = 0) Or (CreamFat.Value = 0)) Then
 
  MsgBox "Íåäîïóñòèìûå çíà÷åíèÿ ïàðàìåòðîâ ñûðîãî ìîëîêà è ñëèâîê"
  GoTo ExitLabel
 
 End If
   
 If (FreeFatMilkFat.Enabled) Then
 
  If (ButterMilkFat.Enabled) Then
  
   If ((FreeFatMilkFat.Value = 0) Or (ButterMilkFat.Value = 0)) Then
   
    MsgBox "Íåäîïóñòèìûå çíà÷åíèÿ ïàðàìåòðîâ îáðàòà èëè ïàõòû"
    GoTo ExitLabel
   
   End If
  
  Else
   
   If (FreeFatMilkFat.Value = 0) Then
      
    MsgBox "Íåäîïóñòèìûå çíà÷åíèÿ ïàðàìåòðîâ îáðàòà"
    GoTo ExitLabel
       
   End If
 
  End If
       
 End If
       
 If (AssortmentListControl.ListCount = 0) Then
   
   AssortmentListControl.Enabled = False
   MsgBox "Ïåðåä ðàñ÷åòîì íåîáõîäèìî çàïîëíèòü àññîðòèìåíò"
   GoTo ExitLabel
  
 End If
 
 CalculationButton.Enabled = True
 CalculationButton.SetFocus
  
 'íà÷àëî ðàñ÷åòà
 
 CreamWeightIteration = 0
 RawMilkWeightIteration = 0
 ButterMilkWeightIteration = 0
 
 CreamWeightFromLowFatProducts = 0
 RawMilkWeightForLowFatProducts = 0
 
 CreamWeightForMediumFatProducts = 0
 RawMilkWeightForMediumFatProducts = 0
 
 CreamWeightForHiFatProducts = 0
 ButterMilkWeightFromHiFatProducts = 0
   
 LowFatMilkVolumeFromLowFatProducts = 0
 LowFatMilkVolumeForMeduimFatProducts = 0
 LowFatMilkVolumeForAllProducts = 0
 
 LowFatMilkVolumeForMeduimFatProductsIteration = 0
 
 DeltaRawMilkNeedForLowFatProducts = 0
 DeltaCreamFromLowFatProducts = 0
 DeltaLowFatMilkVolumeForLowFatProducts = 0
 
 CottageCheeseWheyWeight = 0
 CheeseWheyWeight = 0
     
 Index = 0
 EndIndex = AssortmentListControl.ListCount - 1
   
 For Index = 0 To EndIndex
    
  CurrentProductType = CLng(AssortmentListControl.Column(4, Index)) 'òèï ïðîäóêòà
  
  'MsgBox "Òèï ïðîäóêòà " & CurrentProductType
  'MsgBox "Íàçâàíèå ïðîäóêòà: " & AssortmentListControl.Column(3, Index)
  
  If ((CurrentProductType = 4) Or (CurrentProductType = 5)) Then '4 - òâîðîã, 5 - ñûð
  
   CurrentFat = CSng(AssortmentListControl.Column(5, Index)) 'æèð ñìåñè
   CurrentWeight = CLng(AssortmentListControl.Column(1, Index)) * CLng(AssortmentListControl.Column(6, Index)) / 1000 'êîëè÷åñòâî ñìåñè
   
   'MsgBox "Ðàñõîä ñìåñè " & CLng(AssortmentListControl.Column(6, Index))
   'MsgBox "Ìàññà ïðîäóêòà " & CLng(AssortmentListControl.Column(1, Index))
      
  Else
  
   CurrentFat = CSng(AssortmentListControl.Column(2, Index)) 'æèð ïðîäóêòà
   CurrentWeight = CLng(AssortmentListControl.Column(1, Index)) 'êîë-âî ïðîäóêòà, êã
   
  End If
  
  'MsgBox "Æèð ñìåñè " & CurrentFat
  'MsgBox "Ìàññà ñìåñè " & CurrentWeight
  
  If (CurrentFat <= RawMilkFat.Value) Then 'Æèð ïðîäóêòà ìåíüøå æèðà ñûðîãî ìîëîêà
     
   If ((CurrentProductType = 4) And (CurrentFat <= FreeFatMilkFat.Value)) Then
   
    LowFatMilkVolumeForLowFatProducts = LowFatMilkVolumeFromLowFatProducts + CurrentWeight 'òðåáóåìîå êîëè÷åñòâî îáðàòà äëÿ íèçêîæèðíîãî òâîðîãà
   
   Else
    'íåîáõîäèìîå êîëè÷åñòâî ñûðîãî ìîëîêà äëÿ ïîëó÷åíèÿ ñìåñè
    RawMilkWeightIteration = CurrentWeight * (CreamFat.Value - CurrentFat) / _
                                  (CreamFat.Value - RawMilkFat.Value)
    'ìàññà ñëèâîê, ïîëó÷åííûõ â ðåçóëüòàòå ñåïàðèðîâàíèÿ
    CreamWeightIteration = RawMilkWeightIteration - CurrentWeight
     
    CreamWeightFromLowFatProducts = CreamWeightFromLowFatProducts + CreamWeightIteration 'ñëèâêè îò ïðîèçâîäñòâà íèçêîæèðíîé ïðîäóêöèè è ñìåñåé
    RawMilkWeightForLowFatProducts = RawMilkWeightForLowFatProducts + RawMilkWeightIteration 'ñûðîå ìîëîêî, íåîáõîäèìîå äëÿ ïðîèçâîäñòâà íèçêîæèðíîé ïðîäóêöèè è ñìåñåé
     
   End If
   
   If (CurrentProductType = 4) Then
    'Ìàññà òâîðîæíîé ñûâîðîòêè
    CottageCheeseWheyWeight = CottageCheeseWheyWeight + (CurrentWeight - CLng(AssortmentListControl.Column(1, Index)))
    
   End If
   
   If (CurrentProductType = 5) Then
    'Ìàññû ïîäñûðíîé ñûâîðîòêè
    CheeseWheyWeight = CheeseWheyWeight + (CurrentWeight - CLng(AssortmentListControl.Column(1, Index)))
    
   End If

  Else
  
   If ((CurrentFat > RawMilkFat.Value) And (CurrentFat <= CreamFat.Value)) Then 'Æèð ïðîäóêòà áîëüøå æèðà ñûðîãî ìîëîêà, íî ìåíüøå æèðà ñëèâîê
      
    'MsgBox "Æèð ïðîäóêòà áîëüøå æèðà ñûðîãî ìîëîêà, íî ìåíüøå æèðà ñëèâîê"
    
    '13/09/2022 - was
    'RawMilkWeightIteration = CurrentWeight * (CreamFat.Value - RawMilkFat.Value) / _
    '                                         (RawMilkFat.Value - FreeFatMilkFat.Value)
    'CreamWeightIteration = CurrentWeight - RawMilkWeightIteration
    
    '13/09/2022 - new version
    RawMilkWeightIteration = 0 'ñûðîå ìîëîêî íå èñïîëüçóåì
    'ðàçâîäèì ñëèâêè èç ñåïàðàòîðà îáðàòîì
    CreamWeightIteration = CurrentWeight * (CurrentFat - FreeFatMilkFat.Value) / (CreamFat.Value - FreeFatMilkFat.Value)
    'íóæíàÿ ìàññà îáðàòà
    LowFatMilkVolumeForMeduimFatProductsIteration = CurrentWeight - CreamWeightIteration
    
    'ìàññà îáðàòà îò ïðîäóêòîâ ñðåäíåé æèðíîñòè
    LowFatMilkVolumeForMeduimFatProducts = LowFatMilkVolumeForMeduimFatProducts + LowFatMilkVolumeForMeduimFatProductsIteration
    CreamWeightForMediumFatProducts = CreamWeightForMediumFatProducts + CreamWeightIteration 'ñëèâêè, íåîáõîäèìûå äëÿ ïðîèçâîäñòâà ïðîäóêòîâ ñðåäíåé æèäêîñòè
    RawMilkWeightForMediumFatProducts = RawMilkWeightForMediumFatProducts + RawMilkWeightIteration 'ñûðîå ìîëîêî äëÿ íîðìàëèçàöèè ñëèâîê
    
    If (CurrentProductType = 4) Then
   
     CottageCheeseWheyWeight = CottageCheeseWheyWeight + (CurrentWeight - CLng(AssortmentListControl.Column(1, Index)))
    
    End If
   
    If (CurrentProductType = 5) Then
   
     CheeseWheyWeight = CheeseWheyWeight + (CurrentWeight - CLng(AssortmentListControl.Column(1, Index)))
    
    End If
    
    'MsgBox "Äàííûå ïî ñûðüþ äëÿ òåêóùåãî ïðîäóêòà ñðåäíåé æèðíîñòè"
    'MsgBox "Òèï ïðîäóêòà " & CurrentProductType
    'MsgBox "Ìàññà ñûðîãî ìîëîêà " & RawMilkWeightIteration
    'MsgBox "Ìàññà ñëèâîê " & CreamWeightIteration
    'MsgBox "Ìàññà îáðàòà íà èçãîòîâëåíèå ïðîäóêòà " & LowFatMilkVolumeForMeduimFatProductsIteration
    
    'MsgBox "Îáùàÿ ìàññà íåîáõîäèìûõ ñëèâîê " & CreamWeightForMediumFatProducts
    'MsgBox "Îáùàÿ ìàññà ñûðîãî ìîëîêà " & RawMilkWeightForMediumFatProducts
    'MsgBox "Îáùàÿ ìàññà íåîáõîäèìîãî îáðàòà " & LowFatMilkVolumeForMeduimFatProducts
    
   Else 'Æèð ïðîäóêòà áîëüøå æèðà ñëèâîê
   
    CreamWeightIteration = CurrentWeight * (CurrentFat - ButterMilkFat.Value) / (CreamFat.Value - ButterMilkFat.Value)
    ButterMilkWeightIteration = CreamWeightIteration - CurrentWeight
    
    CreamWeightForHiFatProducts = CreamWeightForHiFatProducts + CreamWeightIteration 'ñëèâêè, íåîáõîäèìûå äëÿ ïðîèçâîäñòâà ïðîäóêòîâ âûñîêîé æèðíîñòè
    ButterMilkWeightFromHiFatProducts = ButterMilkWeightFromHiFatProducts + ButterMilkWeightIteration 'ïàõòà îò ïðîèçâîäñòâà ïðîäóêòîâ âûñîêîé æèäêîñòè
      
   End If
     
  End If
   
 Next Index
 
 RawMilkNeedWeight = 0
 RawMilkAdditionalSeparationNeedWeight = 0
 CreamOutputWeight = 0
 CreamCurrentWeight = 0
 CreamAdditionalSeparationWeight = 0
 FreeFatMilkAdditionalSeparationWeight = 0
 FreeFatMilkCurrentWeight = 0
 FreeFatMilkOutputWeight = 0
 
 'ìàëîæèðíàÿ ïðîäóêöèÿ
 RawMilkNeedWeight = RawMilkWeightForLowFatProducts + RawMilkWeightForMediumFatProducts 'îáùåå êîëè÷åñòâî ïîòðåáîâàâøåãîñÿ ñûðîãî ìîëîêà
 CreamOutputWeight = CreamWeightFromLowFatProducts 'òåêóùèé âûõîä ñëèâîê
 CreamCurrentWeight = CreamWeightFromLowFatProducts 'òåêóùåå êîëè÷åñòâî ñëèâîê
 
 'ïðîäóêöèÿ ñðåäíåé æèðíîñòè
 If (CreamCurrentWeight < CreamWeightForMediumFatProducts) Then 'íóæíî äåëàòü äîïîëíèòåëüíîå ñåïàðèðîâàíèå
 
  'MsgBox "Òðåáóåòñÿ äîïîëíèòåëüíîå ñåïàðèðîâàíèå äëÿ ïîëó÷åíèÿ ñëèâîê. Ïðîäóêòû ñðåäíåé æèðíîñòè"
  
  CreamAdditionalSeparationWeight = CreamWeightForMediumFatProducts - CreamCurrentWeight 'ñòîëüêî ñëèâîê íå õâàòàåò
  
  'MsgBox "Ìàññà ñëèâîê, êîòîðûå íóæíî ïîëó÷èòü ïðè äîïîëíèòåëüíîì ñåïàðèðîâàíèè: " & CreamAdditionalSeparationWeight
  'MsgBox "Æèð ñëèâîê: " & CreamFat.Value
  'MsgBox "Æèð îáðàòà: " & FreeFatMilkFat.Value
  'MsgBox "Æèð ñûðîãî ìîëîêà: " & RawMilkFat.Value
  
  'äîïîëíèòåëüíî ñåïàðèðóåì ñûðîå ìîëîêî äëÿ ïðèãîòîâëåíèÿ ïðîäóêöèè ñðåäíåé æèðíîñòè
  RawMilkAdditionalSeparationNeedWeight = CreamAdditionalSeparationWeight * (CreamFat.Value - FreeFatMilkFat.Value) / _
                                          (RawMilkFat.Value - FreeFatMilkFat.Value)
  
  'MsgBox "Ìàññà ñûðîãî ìîëîêà äëÿ äîïîëíèòåëüíîãî ñåïàðèðîâàíèÿ: " & RawMilkAdditionalSeparationNeedWeight
  
  RawMilkNeedWeight = RawMilkNeedWeight + RawMilkAdditionalSeparationNeedWeight
  CreamOutputWeight = CreamOutputWeight + CreamAdditionalSeparationWeight
  
  'MsgBox "Îáùàÿ ìàññà ñûðîãî ìîëîêà  äëÿ ïðîèçâîäñòâà: " & RawMilkNeedWeight
  'MsgBox "Îáùàÿ ìàññà ñëèâîê  äëÿ ïðîèçâîäñòâà: " & CreamOutputWeight
  
  CreamCurrentWeight = 0 'âñå ñëèâêè óøëè íà ñðåäíþþ æèðíîñòü
  
  FreeFatMilkAdditionalSeparationWeight = RawMilkAdditionalSeparationNeedWeight - CreamAdditionalSeparationWeight 'ïîÿâèëñÿ îáðàò
  FreeFatMilkCurrentWeight = FreeFatMilkAdditionalSeparationWeight
  FreeFatMilkOutputWeight = FreeFatMilkAdditionalSeparationWeight
  
  'MsgBox "Îáðàò ïîñëå äîïîëíèòåëüíîãî ñåïàðèðîâàíèÿ: " & FreeFatMilkAdditionalSeparationWeight
  'MsgBox "Òåêóùàÿ ìàññà îáðàòà äëÿ ðàñ÷åòîâ: " & FreeFatMilkCurrentWeight
  'MsgBox "Òåêóùàÿ ìàññà îáðàòà äëÿ âûâîäà: " & FreeFatMilkOutputWeight
  
 Else ' ñëèâîê äëÿ ïðîäóêöèè ñðåäíåé æèðíîñòè õâàòàåò
 
  CreamCurrentWeight = CreamCurrentWeight - CreamWeightForMediumFatProducts 'êîððåêòèðóåì êîëè÷åñòâî èìåþùèõñÿ ñëèâîê
  
 End If
 
 'ïðîäóêöèÿ áîëüøîé æèðíîñòè
 If (CreamCurrentWeight < CreamWeightForHiFatProducts) Then 'íóæíî äåëàòü äîïîëíèòåëüíîå ñåïàðèðîâàíèå
 
  'MsgBox "Òðåáóåòñÿ äîïîëíèòåëüíîå ñåïàðèðîâàíèå äëÿ ïîëó÷åíèÿ ñëèâîê. Æèðíûå ïðîäóêòû"
  
  CreamAdditionalSeparationWeight = CreamWeightForHiFatProducts - CreamCurrentWeight 'ñòîëüêî ñëèâîê íå õâàòàåò
 
  'MsgBox "Ìàññà ñëèâîê, êîòîðûå íóæíî ïîëó÷èòü ïðè äîïîëíèòåëüíîì ñåïàðèðîâàíèè: " & CreamAdditionalSeparationWeight
  'MsgBox "Æèð ñëèâîê: " & CreamFat.Value
  'MsgBox "Æèð îáðàòà: " & FreeFatMilkFat.Value
  'MsgBox "Æèð ñûðîãî ìîëîêà: " & RawMilkFat.Value
  
  'äîïîëíèòåëüíî ñåïàðèðóåì ñûðîå ìîëîêî äëÿ ïðèãîòîâëåíèÿ ïðîäóêöèè áîëüøîé æèðíîñòè
   RawMilkAdditionalSeparationNeedWeight = CreamAdditionalSeparationWeight * (CreamFat.Value - FreeFatMilkFat.Value) / _
                                           (RawMilkFat.Value - FreeFatMilkFat.Value)
                                           
  'MsgBox "Ìàññà ñûðîãî ìîëîêà äëÿ äîïîëíèòåëüíîãî ñåïàðèðîâàíèÿ: " & RawMilkAdditionalSeparationNeedWeight
  
  RawMilkNeedWeight = RawMilkNeedWeight + RawMilkAdditionalSeparationNeedWeight
  CreamOutputWeight = CreamOutputWeight + CreamAdditionalSeparationWeight
  
  'MsgBox "Îáùàÿ ìàññà ñûðîãî ìîëîêà  äëÿ ïðîèçâîäñòâà: " & RawMilkNeedWeight
  'MsgBox "Îáùàÿ ìàññà ñëèâîê  äëÿ ïðîèçâîäñòâà: " & CreamOutputWeight
  
  CreamCurrentWeight = 0 'âñå ñëèâêè óøëè íà ñðåäíþþ æèðíîñòü
  
  FreeFatMilkAdditionalSeparationWeight = RawMilkAdditionalSeparationNeedWeight - CreamAdditionalSeparationWeight 'ïîÿâèëñÿ îáðàò
  FreeFatMilkCurrentWeight = FreeFatMilkCurrentWeight + FreeFatMilkAdditionalSeparationWeight
  FreeFatMilkOutputWeight = FreeFatMilkOutputWeight + FreeFatMilkAdditionalSeparationWeight
  
  'MsgBox "Îáðàò ïîñëå äîïîëíèòåëüíîãî ñåïàðèðîâàíèÿ: " & FreeFatMilkAdditionalSeparationWeight
  'MsgBox "Òåêóùàÿ ìàññà îáðàòà äëÿ ðàñ÷åòîâ: " & FreeFatMilkCurrentWeight
  'MsgBox "Òåêóùàÿ ìàññà îáðàòà äëÿ âûâîäà: " & FreeFatMilkOutputWeight
  
 Else ' ñëèâîê äëÿ ïðîäóêöèè âûñîêîé æèðíîñòè õâàòàåò
 
  CreamCurrentWeight = CreamCurrentWeight - CreamWeightForHiFatProducts
  
 End If
 
 'íåæèðíûé òâîðîã (äåëàåì òîëüêî èç îáðàòà), åñëè îáðàòà íå õâàòàåò, òî äåëàåì äîïîëíèòåëüíîå ñåïàðèðîâàíèå
 '2022-09-13 - new
 'ïèòüåâûå ñëèâêè òàêæå äåëàåì ñ èñïîëüçîâàíèåì îáðàòà
 'ñêîëüêî íóæíî îáðàòà íà ïðîèçâîäñòâî
 LowFatMilkVolumeForAllProducts = LowFatMilkVolumeForMeduimFatProducts + LowFatMilkVolumeForLowFatProducts
 
 'MsgBox "Êîëè÷åñòâî îáðàòà, êîòîðîå íóæíî äëÿ ïðîèçâîäñòâà àññîðòèìåíòà" & LowFatMilkVolumeForAllProducts
  
 If (LowFatMilkVolumeForAllProducts >= FreeFatMilkCurrentWeight) Then
  
  'MsgBox "Òðåáóåòñÿ äîïîëíèòåëüíîå ñåïàðèðîâàíèå äëÿ ïîëó÷åíèÿ îáðàòà. Âñå ïðîäóêòû"
 
  'Êîëè÷åñòâà îáðàòà äëÿ ïðèãîòîâëåíèÿ òâîðîãà íå äîñòàòî÷íî, íóæíî äîïîëíèòåëüíîå ñåïàðèðîâàíèå
  DeltaLowFatMilkVolumeForLowFatProducts = LowFatMilkVolumeForAllProducts - FreeFatMilkCurrentWeight
   
  FreeFatMilkCurrentWeight = 0
  FreeFatMilkOutputWeight = FreeFatMilkOutputWeight + DeltaLowFatMilkVolumeForLowFatProducts
   
  DeltaRawMilkNeedForLowFatProducts = DeltaLowFatMilkVolumeForLowFatProducts * (CreamFat.Value - FreeFatMilkFat.Value) / _
                                                              (CreamFat.Value - RawMilkFat.Value)
                                                              
  DeltaCreamFromLowFatProducts = DeltaRawMilkNeedForLowFatProducts - DeltaLowFatMilkVolumeForLowFatProducts
   
  RawMilkNeedWeight = RawMilkNeedWeight + DeltaRawMilkNeedForLowFatProducts
  
  CreamOutputWeight = CreamOutputWeight + DeltaCreamFromLowFatProducts
  
  CreamCurrentWeight = CreamCurrentWeight + DeltaCreamFromLowFatProducts
   
 Else
  ' Îáðàòà äëÿ ïðèãîòîâëåíèÿ òâîðîãà è ïèòüåâûõ ñëèâîê äîñòàòî÷íî
  FreeFatMilkCurrentWeight = FreeFatMilkCurrentWeight - LowFatMilkVolumeForAllProducts

 End If
 
 RawMilkNeededQuantity.Value = CLng(RawMilkNeedWeight)
 RawMilkQuantityDifference.Value = CLng(RawMilkQuantity.Value - RawMilkNeededQuantity.Value)
   
 SeparatedCreamQuantity.Value = CLng(CreamOutputWeight)
 SeparatedCreamQuantityDifference.Value = CLng(CreamCurrentWeight)
   
 FreeFatMilkQuantity.Value = CLng(FreeFatMilkOutputWeight)
 FreeFatMilkQuantityDifference.Value = CLng(FreeFatMilkCurrentWeight)
 
 CheeseWhey.Value = CLng(CheeseWheyWeight)
 CottageCheeseWhey.Value = CLng(CottageCheeseWheyWeight)
 ButterMilk.Value = CLng(ButterMilkWeightFromHiFatProducts)
 
'Îêîí÷àíèå ðàñ÷åòa
 CalculationResultReady = True
 ReportButton.Enabled = True
 
ExitLabel:

PagesVisiability

End Sub


Private Sub CreamFatField_AfterUpdate()

 Dim CreamFat As TextBox
 
 Set CreamFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("CreamFatField")
 
 If (CreamFatAcc <> CreamFat.Value) Then
 
  CreamFatAcc = CreamFat.Value
  
  ResetCalculationResult
  
 End If
 
 CreamFat.SetFocus
 
 PagesVisiability
 ControlView_SourceBalancePage
 
End Sub

Private Sub Form_Open(Cancel As Integer)
 
 'SourceBalancePageControls
 Dim CalculationButton As CommandButton
 Dim ReportButton As CommandButton
 
 'Pages
 Dim SourceMixturePage As Page
 
 'DayBalance page controls
 Dim RawMilkQuantity As TextBox
 Dim RawMilkFat As TextBox
 Dim CreamFat As TextBox
 Dim FreeFatMilkFat As TextBox
 Dim ButterMilkFat As TextBox
 
 Dim RawMilkMorningQuantity As TextBox
 Dim RawMilkMorningFat As TextBox
 Dim RawMilkMorningQuantityStandard As TextBox
 
 Dim RawMilkDayQuantity As TextBox
 Dim RawMilkDayFat As TextBox
 Dim RawMilkDayQuantityStandard As TextBox
 
 Dim RawMilkEveningQuantity As TextBox
 Dim RawMilkEveningFat As TextBox
 Dim RawMilkEveningQuantityStandard As TextBox
 
 Dim SourceCalcDataType As ComboBox

 'SourceBalancePageControls
 Set CalculationButton = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("CalculateButton")
 Set ReportButton = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("ReportButton")
 
 Set RawMilkQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkQuantityField")
 Set RawMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkFatField")
 Set CreamFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("CreamFatField")
 Set FreeFatMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("FreeFatMilkFatField")
 Set ButterMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("ButterMilkFatField")
 
 'Pages
 Set SourceBalancePage = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab")
 Set SourceMixturePage = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceMixtureTab")
 
 'DayBalance page controls
 Set RawMilkMorningQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkMorningAmount")
 Set RawMilkMorningFat = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkMorningFat")
 Set RawMilkMorningQuantityStandard = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkMorningAmountStandard")
 
 Set RawMilkDayQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkDayAmount")
 Set RawMilkDayFat = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkDayFat")
 Set RawMilkDayQuantityStandard = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkDayAmountStandard")
 
 Set RawMilkEveningQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkEveningAmount")
 Set RawMilkEveningFat = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkEveningFat")
 Set RawMilkEveningQuantityStandard = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkEveningAmountStandard")
 
 Set SourceCalcDataType = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("SourceCalcDataTypeList")
 Set DataItemName = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("DataFiledNameLabel")
 Set CalcDataSource = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("CalculationDataSourceList")
   
 'SourceBalance page controls
 CalculationButton.Enabled = False
 ReportButton.Enabled = False
 
 SourceMixturePage.Visible = True
 SourceMixturePage.SetFocus
 
 RawMilkQuantity.Value = 0
 RawMilkFat.Value = 0
 CreamFat.Value = 0
 FreeFatMilkFat.Value = 0
 ButterMilkFat.Value = 0
  
 ResetCalculationResult
 
 RawMilkQuantityAcc = 0
 RawMilkFatAcc = 0
 CreamFatAcc = 0
 FreeFatMilkFatAcc = 0
 ButterMilkFatAcc = 0
 
 'DayBalance page controls
 DayBalanceCalculEnable = False
 
 RawMilkMorningQuantity.Value = 0
 RawMilkMorningFat.Value = 0
 RawMilkMorningQuantityStandard.Value = 0
 
 RawMilkDayQuantity.Value = 0
 RawMilkDayFat.Value = 0
 RawMilkDayQuantityStandard.Value = 0
 
 RawMilkEveningQuantity.Value = 0
 RawMilkEveningFat.Value = 0
 RawMilkEveningQuantityStandard.Value = 0
 
 SourceCalcDataType.Value = SourceCalcDataType.Column(0, 0)
 
 CalculationResultReady = False
 
 PagesVisiability
 ControlView_MixturesPage
 
End Sub

Private Sub FreeFatMilkFatField_AfterUpdate()
 
 Dim FreeFatMilkFat As TextBox

 Set FreeFatMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("FreeFatMilkFatField")
 
 If (FreeFatMilkFatAcc <> FreeFatMilkFat.Value) Then
 
  FreeFatMilkFatAcc = FreeFatMilkFat.Value
  
  ResetCalculationResult
  
 End If
 
 FreeFatMilkFat.SetFocus
 
 PagesVisiability
 ControlView_SourceBalancePage

End Sub

Private Sub RawMilkDayAmount_AfterUpdate()

 Dim RawMilkDayQuantity As TextBox
 Dim RawMilkDayFat As TextBox
 Dim RawMilkDayQuantityStandard As TextBox
 
 Set RawMilkDayQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkDayAmount")
 Set RawMilkDayFat = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkDayFat")
 Set RawMilkDayQuantityStandard = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkDayAmountStandard")
 
 RawMilkDayQuantityStandard.Value = CInt(RawMilkDayQuantity.Value * RawMilkDayFat.Value / 3.4)

End Sub

Private Sub RawMilkDayFat_AfterUpdate()

 Dim RawMilkDayQuantity As TextBox
 Dim RawMilkDayFat As TextBox
 Dim RawMilkDayQuantityStandard As TextBox
 
 Set RawMilkDayQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkDayAmount")
 Set RawMilkDayFat = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkDayFat")
 Set RawMilkDayQuantityStandard = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkDayAmountStandard")
 
 RawMilkDayQuantityStandard.Value = CInt(RawMilkDayQuantity.Value * RawMilkDayFat.Value / 3.4)

End Sub

Private Sub RawMilkFatField_AfterUpdate()
 
 Dim RawMilkFat As TextBox
   
 Set RawMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkFatField")
 
 If (RawMilkFatAcc <> RawMilkFat.Value) Then
 
  RawMilkFatAcc = RawMilkFat.Value
  
  ResetCalculationResult
  
 End If
 
 PagesVisiability
 ControlView_SourceBalancePage
   
End Sub

Private Sub RawMilkMorningAmount_AfterUpdate()

 Dim RawMilkMorningQuantity As TextBox
 Dim RawMilkMorningFat As TextBox
 Dim RawMilkMorningQuantityStandard As TextBox
 
 Set RawMilkMorningQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkMorningAmount")
 Set RawMilkMorningFat = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkMorningFat")
 Set RawMilkMorningQuantityStandard = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkMorningAmountStandard")
 
 RawMilkMorningQuantityStandard.Value = CInt(RawMilkMorningQuantity.Value * RawMilkMorningFat.Value / 3.4)

End Sub

Private Sub RawMilkMorningFat_AfterUpdate()

 Dim RawMilkMorningQuantity As TextBox
 Dim RawMilkMorningFat As TextBox
 Dim RawMilkMorningQuantityStandard As TextBox
 
 Set RawMilkMorningQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkMorningAmount")
 Set RawMilkMorningFat = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkMorningFat")
 Set RawMilkMorningQuantityStandard = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("RawMilkMorningAmountStandard")
 
 RawMilkMorningQuantityStandard.Value = CInt(RawMilkMorningQuantity.Value * RawMilkMorningFat.Value / 3.4)

End Sub

Private Sub RawMilkQuantityField_AfterUpdate()
  
 Dim RawMilkQuantity As TextBox
 
 Set RawMilkQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkQuantityField")
  
 If (RawMilkQuantityAcc <> RawMilkQuantity.Value) Then
 
  RawMilkQuantityAcc = RawMilkQuantity.Value
  
  ResetCalculationResult
 
 End If
 
 RawMilkQuantity.SetFocus
 
 PagesVisiability
 ControlView_SourceBalancePage
 
End Sub

Private Sub SystemMixtureListAddButton_Click()

 DoCmd.OpenForm "AddMixtureInSystemForm"
 
End Sub

Private Sub SystemMixtureListChangeButton_Click()

 DoCmd.OpenForm "ChangeMixtureInSystemForm"
 
End Sub

Private Sub ReportButton_Click()

'2016/07/29

Dim ExcelSheet As Object

Dim Row, TableRow, Column As Integer

Dim AssortmentListControl As ListBox

Dim ReportFileName As String
Dim CurMonth, CurDay, CurHour, CurMinute, CurSec As Integer
Dim CurMonthStr, CurDayStr, CurHourStr, CurMinuteStr, CurSecStr As String
Dim CurData As Date
Dim CurTime As Date

Dim RawMilkQuantity As TextBox
Dim RawMilkFat As TextBox
Dim CreamFat As TextBox
Dim FreeFatMilkFat As TextBox
Dim ButterMilkFat As TextBox

Dim RawMilkNeededQuantity As TextBox
Dim RawMilkQuantityDifference As TextBox
Dim SeparatedCreamQuantity As TextBox
Dim SeparatedCreamQuantityDifference As TextBox
Dim FreeFatMilkQuantity As TextBox
Dim FreeFatMilkQuantityDifference As TextBox
Dim CottageCheeseWhey As TextBox
Dim CheeseWhey As TextBox
Dim ButterMilk As TextBox

Dim ReportButton As CommandButton

Set AssortmentListControl = Forms("MainForm").Controls("TabControlMainForm").Pages("AssortmentTab").Controls("AssortmentproductList")

Set ExcelSheet = CreateObject("Excel.Sheet")

Set RawMilkQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkQuantityField")
Set RawMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkFatField")
Set CreamFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("CreamFatField")
Set FreeFatMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("FreeFatMilkFatField")
Set ButterMilkFat = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("ButterMilkFatField")

Set RawMilkNeededQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkNeededQuantityField")
Set RawMilkQuantityDifference = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("RawMilkQuantityDifferenceField")
Set SeparatedCreamQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("SeparatedCreamQuantityField")
Set SeparatedCreamQuantityDifference = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("SeparatedCreamQuantityDifferenceField")
Set FreeFatMilkQuantity = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("FreeFatMilkQuantityField")
Set FreeFatMilkQuantityDifference = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("FreeFatMilkQuantityDifferenceField")
Set CottageCheeseWhey = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("CottageCheeseWheyField")
Set CheeseWhey = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("CheeseWheyField")
Set ButterMilk = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("ButterMilkField")

Set ReportButton = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceBalanceTab").Controls("ReportButton")

'MsgBox "Test2"

'ReportButton.Enabled = False

'AssortmentListControl.Requery

'MsgBox AssortmentListControl.ListCount

If (AssortmentListControl.ListCount > 0) Then

 ExcelSheet.Application.Visible = False
 
 CurDate = Date
 CurTime = Time
 
 CurMonth = Month(CurDate)
 CurDay = Day(CurDate)
 CurHour = Hour(CurTime)
 CurMinute = Minute(CurTime)
 CurSec = Second(CurTime)
 
 CurMonthStr = CStr(CurMonth)
 CurDayStr = CStr(CurDay)
 CurHourStr = CStr(CurHour)
 CurMinuteStr = CStr(CurMinute)
 CurSecStr = CStr(CurSec)
 
 If (CurMonth < 10) Then
  CurMonthStr = "0" + CurMonthStr
 End If
 
 If (CurDay < 10) Then
  CurDayStr = "0" + CurDayStr
 End If
 
 If (CurHour < 10) Then
  CurHourStr = "0" + CurHourStr
 End If
 
 If (CurMinute < 10) Then
  CurMinuteStr = "0" + CurMinuteStr
 End If
 
 If (CurSec < 10) Then
  CurSecStr = "0" + CurSecStr
 End If
 
 'MsgBox "Test1"
 
 'MsgBox CurrentProject.Path
 
 ReportFileName = CurrentProject.Path + "\BalanceReport_" + CStr(Year(CurDate)) + "_" + CurMonthStr + "_" + CurDayStr + "_" _
                  + CurHourStr + CurMinuteStr + CurSecStr + ".xls"
 
 ' Make Excel visible through the Application object.
 ExcelSheet.Application.Visible = True
 ' Place some text in the first cell of the sheet.

 ExcelSheet.Application.Cells(1, 1).Value = "ÑÛÐÜÅÂÎÉ ÁÀËÀÍÑ"
 
 ExcelSheet.Application.Range("A1:B1").Select
        
 With ExcelSheet.Application.Selection
  .HorizontalAlignment = -4108
  .VerticalAlignment = -4107
  .WrapText = False
  .Orientation = 0
  .AddIndent = False
  .IndentLevel = 0
  .ShrinkToFit = False
  .ReadingOrder = -5002
  .MergeCells = False
 End With
    
 ExcelSheet.Application.Selection.Merge
 ExcelSheet.Application.Selection.Font.Bold = True
 
 
 ExcelSheet.Application.Cells(3, 1).Value = "I. Èñõîäíûå äàííûå"
 ExcelSheet.Application.Cells(3, 1).Font.Italic = True
 ExcelSheet.Application.Cells(3, 1).Font.Bold = True
  
 ExcelSheet.Application.Cells(4, 1).Value = "Ìîëîêî ñûðîå, ì.ä.æ., %:"
 ExcelSheet.Application.Cells(5, 1).Value = "Ìîëîêî ñûðîå, êîëè÷åñòâî, êã:"
 ExcelSheet.Application.Cells(6, 1).Value = "Ñëèâêè, ì.ä.æ., %:"
 ExcelSheet.Application.Cells(7, 1).Value = "Îáåçæèðåííîå ìîëîêî, ì.ä.æ., %:"
 ExcelSheet.Application.Cells(8, 1).Value = "Ïàõòà, ì.ä.æ., %:"

 ExcelSheet.Application.Cells(4, 2).Value = RawMilkFat.Value
 ExcelSheet.Application.Cells(5, 2).Value = RawMilkQuantity.Value
 ExcelSheet.Application.Cells(6, 2).Value = CreamFat.Value
 ExcelSheet.Application.Cells(7, 2).Value = FreeFatMilkFat.Value
 ExcelSheet.Application.Cells(8, 2).Value = ButterMilkFat.Value

 ExcelSheet.Application.Cells(10, 1).Value = "II. Àññîðòèìåíò"
 ExcelSheet.Application.Cells(10, 1).Font.Italic = True
 ExcelSheet.Application.Cells(10, 1).Font.Bold = True
 
 For TableRow = 0 To AssortmentListControl.ListCount - 1
  Row = 11 + TableRow
  ExcelSheet.Application.Cells(Row, 1).Value = CStr(AssortmentListControl.Column(0, TableRow)) + " , êã:"
  ExcelSheet.Application.Cells(Row, 2).Value = AssortmentListControl.Column(1, TableRow)
 Next TableRow

 Row = Row + 2
 ExcelSheet.Application.Cells(Row, 1).Value = "III. Áàëàíñ"
 ExcelSheet.Application.Cells(Row, 1).Font.Italic = True
 ExcelSheet.Application.Cells(Row, 1).Font.Bold = True
 
 Row = Row + 1
 ExcelSheet.Application.Cells(Row, 1).Value = "Ðàñõîä ñûðîãî ìîëîêà, êã:"
 ExcelSheet.Application.Cells(Row, 2).Value = RawMilkNeededQuantity.Value
 Row = Row + 1
 ExcelSheet.Application.Cells(Row, 1).Value = "Îñòàòîê ñûðîãî ìîëîêà, êã:"
 ExcelSheet.Application.Cells(Row, 2).Value = RawMilkQuantityDifference.Value
 Row = Row + 1
 ExcelSheet.Application.Cells(Row, 1).Value = "Âûõîä ñëèâîê, êã:"
 ExcelSheet.Application.Cells(Row, 2).Value = SeparatedCreamQuantity.Value
 Row = Row + 1
 ExcelSheet.Application.Cells(Row, 1).Value = "Îñòàòîê ñëèâîê, êã:"
 ExcelSheet.Application.Cells(Row, 2).Value = SeparatedCreamQuantityDifference.Value
 Row = Row + 1
 ExcelSheet.Application.Cells(Row, 1).Value = "Âûõîä îáåçæèðåííîãî ìîëîêà, êã:"
 ExcelSheet.Application.Cells(Row, 2).Value = FreeFatMilkQuantity.Value
 Row = Row + 1
 ExcelSheet.Application.Cells(Row, 1).Value = "Îñòàòîê îáåçæèðåííîãî ìîëîêà, êã:"
 ExcelSheet.Application.Cells(Row, 2).Value = FreeFatMilkQuantityDifference.Value
  
 Row = Row + 2
 ExcelSheet.Application.Cells(Row, 1).Value = "IV. Âòîðè÷íîå ñûðüå"
 ExcelSheet.Application.Cells(Row, 1).Font.Italic = True
 ExcelSheet.Application.Cells(Row, 1).Font.Bold = True
 
 Row = Row + 1
 ExcelSheet.Application.Cells(Row, 1).Value = "Ñûâîðîòêà òâîðîæíàÿ, êã:"
 ExcelSheet.Application.Cells(Row, 2).Value = CottageCheeseWhey.Value
 
 Row = Row + 1
 ExcelSheet.Application.Cells(Row, 1).Value = "Ñûâîðîòêà ïîäñûðíàÿ, êã:"
 ExcelSheet.Application.Cells(Row, 2).Value = CheeseWhey.Value
 
 Row = Row + 1
 ExcelSheet.Application.Cells(Row, 1).Value = "Ïàõòà, êã:"
 ExcelSheet.Application.Cells(Row, 2).Value = ButterMilk.Value
    
 ExcelSheet.Application.Columns(1).ColumnWidth = 31.14
 ExcelSheet.Application.Columns(2).ColumnWidth = 9.29
  
 'ExcelSheet.SaveAs FileName:="D:\Test.xlsx"
 ExcelSheet.SaveAs FileName:=ReportFileName, _
                   FileFormat:=18, _
                   Password:="", _
                   WriteResPassword:="", _
                   ReadOnlyRecommended:=False, _
                   CreateBackup:=False, _
                   AccessMode:=3, _
                   ConflictResolution:=2, _
                   AddToMru:=False, _
                   Local:=False
Else
 
 MsgBox "Àññîðòèìåíò ïóñò. Ñîçäàíèå îò÷åòà ïðåðâàíî."

End If

' Close Excel with the Quit method on the Application object.
ExcelSheet.Application.Quit

' Release the object variable.
Set ExcelSheet = Nothing

ReportButton.Enabled = True

End Sub

Private Sub SourceCalcDataTypeList_AfterUpdate()

  Dim SourceCalcDataType As ComboBox
  Dim DataItemName As Label
  
  Set SourceCalcDataType = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("SourceCalcDataTypeList")
 
  Set DataItemName = Forms("MainForm").Controls("TabControlMainForm").Pages("DayBalanceTab").Controls("DataFiledNameLabel")
  
  If (SourceCalcDataType.Value = 2) Then
  
   DataItemName.Caption = "Íàçâàíèå ïðîäóêòà"
  
  Else
  
   DataItemName.Caption = "Íàçâàíèå ñìåñè"
  
  End If

End Sub

Private Sub SystemMixtureList_Click()
 
 ControlView_MixturesPage

End Sub

Private Sub SystemMixureListAddButton_Click()

 DoCmd.OpenForm "AddMixtureToSystemForm"
 
End Sub

Private Sub SystemMixuresListChangeButton_Click()

 DoCmd.OpenForm "ChangeMixtureInSystemForm"

End Sub

Private Sub SystemMixuresListClearButton_Click()

 'MixturePageControls
 Dim SystemMixtureListControl As ListBox
 
   'ProductsPageControl
 Dim SystemProductListControl As ListBox
    
 Dim Response As Integer
  
 Dim dbsCurrent As Database
 Dim qdfQuery As QueryDef
 Dim QueryRecords As Recordset
 
 'MixturePageControls
 Set SystemMixtureListControl = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceMixtureTab").Controls("SystemMixtureList")
 
  'ProductsPageControls
 Set SystemProductListControl = Forms("MainForm").Controls("TabControlMainForm").Pages("ProductTab").Controls("SystemProductListForEdit")
   
 Set dbsCurrent = CurrentDb
 Set qdfQuery = dbsCurrent.CreateQueryDef("")
  
 Response = vbCancel
 
 SystemMixtureListControl.Requery
 SystemProductListControl.Requery
 
 If (SystemMixtureListControl.ListCount <> 0) Then
 
  If (SystemProductListControl.ListCount = 0) Then
  
   Response = MsgBox("Óäàëèòü âñå ñìåñè èç áàçû?", vbOKCancel)
   
   If (Response = vbOK) Then
  
    qdfQuery.SQL = "DELETE FROM SystemSourceMixtureTable;"
    
    qdfQuery.Execute
    
    PagesVisiability
    
    If (SystemMixtureListControl.ListCount <> 0) Then
     
     MsgBox "Âî âðåìÿ î÷èñòêè ñïèñêà ñìåñåé ïðîèçîøëà îøèáêà."
    
    End If
    
   Else
   
    MsgBox "Îïåðàöèÿ î÷èòñêè ñïèñêà ñìåñåé áûëà ïðåâàíà îïåðàòîðîì."
   
   End If
   
  Else
  
   MsgBox "Â ñèñòåìå åñòü ïðîäóêòû. Î÷èñòèòü ñïèñîê ñìåñåé íåâîçìîæíî. Îïåðàöèÿ ïðåðâàíà."
   
  End If
  
 Else
 
  MsgBox "Â ñèñòåìå íåò çàðåãèñòðèðîâàííûõ ñìåñåé. Îïåðàöèÿ ïðåðâàíà."
 
 End If
       
 qdfQuery.Close
 dbsCurrent.Close
 
 ControlView_MixturesPage

End Sub

Private Sub SystemMixuresListDeleteButton_Click()

    'MixturePageControls
 Dim SystemMixtureListControl As ListBox
    
 Dim TempInt As Integer
 Dim Response As Integer
  
 Dim dbsCurrent As Database
 Dim qdfQuery As QueryDef
 Dim QueryRecords As Recordset
 
  'MixturePageControls
 Set SystemMixtureListControl = Forms("MainForm").Controls("TabControlMainForm").Pages("SourceMixtureTab").Controls("SystemMixtureList")
   
 Set dbsCurrent = CurrentDb
 Set qdfQuery = dbsCurrent.CreateQueryDef("")
 
 Response = vbCancel
 
 If (SystemMixtureListControl.ListCount <> 0) Then
  
  If (SystemMixtureListControl.ListIndex > -1) Then
  
   Response = MsgBox("Óäàëèòü ñìåñü èç áàçû?", vbOKCancel)
   
   If (Response = vbOK) Then
   
    TempInt = SystemMixtureListControl.ListCount
    
    qdfQuery.SQL = "SELECT SystemProductID " & _
                   "FROM SystemProductListTable " & _
                   "WHERE (SystemMixtureID=" & _
                   SystemMixtureListControl.Column(0, SystemMixtureListControl.ListIndex) & ");"

    Set QueryRecords = qdfQuery.OpenRecordset
       
    If (QueryRecords.RecordCount = 0) Then 'ïîèñê â ñïèñêå ïðîäóêòîâ ïðîäóêòà, ó êîòîðîãî óêàçàíà âûäåëåííàÿ ñìåñü
       'ïðîäóêò íå íàéäåí, ñìåñü ìîæåò áûòü óäàëåíà
     qdfQuery.SQL = "DELETE FROM SystemSourceMixtureTable WHERE (IDSourceMixture = " & _
                    SystemMixtureListControl.Column(0, SystemMixtureListControl.ListIndex) & ");"
     
     qdfQuery.Execute
     
     PagesVisiability
     
     If ((TempInt - SystemMixtureListControl.ListCount) <> 1) Then
     
      MsgBox "Ïðè óäàëåíèè ñìåñè ïðîèçîøëà îøèáêà"
      
     End If
      
    Else
       'ïðîäóêò íàéäåí, ñìåñü íå ìîæåò áûòü óäàëåíà
     MsgBox "Îïåðàöèÿ óäàëåíèÿ ñìåñè áûëà ïðåâàíà. Â ñèñòåìå åñòü ïðîäóêò, ñâÿçàííûé ñ íåé"
      
    End If
       
    QueryRecords.Close

   Else
   
    MsgBox "Îïåðàöèÿ óäàëåíèÿ ñìåñè áûëà ïðåâàíà îïåðàòîðîì"
   
   End If
   
  Else
  
   MsgBox "Íå âûáðàíà íè îäíà ñìåñü äëÿ óäàëåíèÿ"
  
  End If
  
 Else
 
  MsgBox "Â ñèñòåìå íåò çàðåãèñòðèðîâàííûõ ñìåñåé"
 
 End If
   
 qdfQuery.Close
 dbsCurrent.Close
 
 ControlView_MixturesPage
 
End Sub

Private Sub SystemProductListAddButton_Click()

 DoCmd.OpenForm "AddProductToSystemForm"
 
End Sub


Private Sub SystemProductListChangeButton_Click()

 DoCmd.OpenForm "ChangeProductInSystemForm"
 
End Sub

Private Sub SystemProductListCleraAllButton_Click()
   
 Dim Response As Integer
 
 Dim dbsCurrent As Database
 Dim qdfQuery As QueryDef
 
 Dim AssortmentListControl As ListBox
 Dim SystemProductListControl As ListBox
 
 Set AssortmentListControl = Forms("MainForm").Controls("TabControlMainForm").Pages("AssortmentTab").Controls("AssortmentproductList")
 Set SystemProductListControl = Forms("MainForm").Controls("TabControlMainForm").Pages("ProductTab").Controls("SystemProductListForEdit")
 
 Set dbsCurrent = CurrentDb
 Set qdfQuery = dbsCurrent.CreateQueryDef("")
  
 Response = vbCancel
 
 AssortmentListControl.Requery
 SystemProductListControl.Requery
 
 If (SystemProductListControl.ListCount <> 0) Then
 
  If (AssortmentListControl.ListCount = 0) Then
  
   Response = MsgBox("Óäàëèòü âñå ïðîäóêòû èç áàçû?", vbOKCancel)
   
   If (Response = vbOK) Then
  
    qdfQuery.SQL = "DELETE FROM SystemProductListTable;"
    
    qdfQuery.Execute
    
    PagesVisiability
    
    If (SystemProductListControl.ListCount <> 0) Then
     
     MsgBox "Âî âðåìÿ î÷èñòêè ñïèñêà ïðîäóêòîâ ïðîèçîøëà îøèáêà."
    
    End If
    
   Else
   
    MsgBox "Îïåðàöèÿ î÷èñòêè ñïèñêà ïðîäóêòîâ áûëà ïðåâàíà îïåðàòîðîì."
   
   End If
  
  Else
  
   MsgBox "Ñïèñîê ïðîäóêòîâ â àññîðòèìåíòå äîëæåí áûòü ïóñò. Îïåðàöèÿ ïðåðâàíà."
  
  End If
  
 Else
 
  MsgBox "Â áàçå íåò ïðîäóêòîâ. Îïåðàöèÿ ïðåðâàíà."
 
 End If
 
 qdfQuery.Close
 dbsCurrent.Close
 
 ControlView_ProductPage
 
End Sub

Private Sub SystemProductListDeleteButton_Click()
 
 Dim Response As Integer
 Dim TempInt As Integer
 
 Dim dbsCurrent As Database
 Dim qdfQuery As QueryDef
 Dim QueryRecords As Recordset
  
 Dim SystemProductListControl As ListBox
 
 Set SystemProductListControl = Forms("MainForm").Controls("TabControlMainForm").Pages("ProductTab").Controls("SystemProductListForEdit")
 
 Set dbsCurrent = CurrentDb
 Set qdfQuery = dbsCurrent.CreateQueryDef("")
 
 Response = vbCancel
 
 If (SystemProductListControl.ListCount <> 0) Then 'Íà÷àëî îáðàáîòêè íàæàòèÿ êíîïêè; Ñïèñîê ñèñòåìíûõ ïðîäóêòîâ íå ïóñò
 
  If (SystemProductListControl.ListIndex > -1) Then
  
   Response = MsgBox("Óäàëèòü ïðîäóêò èç áàçû?", vbOKCancel)
   
   If (Response = vbOK) Then 'Ñïðîñèëè îá óäàëåíèè ïðîäóêòà èç ñïèñêà. Îòâåò îïåðàòîðà óòâåðäèòåëüíûé
   
    'Ãîòîâèì óäàëåíèå ïðîäóêòà èç ñïèñêà
    TempInt = SystemProductListControl.ListCount
    
    'Ïðîâåðÿåì, ÷òî ïðîäóêò îòñóòñòâóåò â òåêóùåì ñïèñêå àññîðòèìåíòà
    qdfQuery.SQL = "SELECT AssortmentProductID " & _
                   "FROM AssortmentProductTable " & _
                   "WHERE AssortmentProductTable.AssortmentProductID = " & SystemProductListControl.Column(0, SystemProductListControl.ListIndex) & ";"
    
    Set QueryRecords = qdfQuery.OpenRecordset
    
    If (QueryRecords.RecordCount = 0) Then
       'ïðîäóêò â àññîðòèìåíòå íå íàéäåí, ìîæíî óäàëÿòü
     
     qdfQuery.SQL = "DELETE FROM SystemProductListTable " & _
                    "WHERE SystemProductID = " & SystemProductListControl.Column(0, SystemProductListControl.ListIndex) & ";"
 
     qdfQuery.Execute
 
     PagesVisiability
     
     If ((TempInt - SystemProductListControl.ListCount) <> 1) Then
     
      MsgBox "Ïðè óäàëåíèè ïðîäóêòà ïðîèçîøëà îøèáêà."
      
     End If
   
    Else
      
     MsgBox "Îïåðàöèÿ óäàëåíèÿ ïðîäóêòà èç áàçû áûëà ïðåâàíà. Âûáðàííûé ïðîäóêò åñòü â àññîðòèìåíòå."
      
    End If
      
    QueryRecords.Close
   
   Else

    MsgBox "Îïåðàöèÿ óäàëåíèÿ ñìåñè áûëà ïðåâàíà îïåðàòîðîì."
    
   End If
  
  Else
  
   MsgBox "Íå âûáðàí íè îäèí ïðîäóêò äëÿ óäàëåíèÿ. Îïåðàöèÿ ïðåðâàíà."
  
  End If
  
 Else 'Íà÷àëî îáðàáîòêè íàæàòèÿ êíîïêè; Ñïèñîê ñèñòåìíûõ ïðîäóêòîâ ïóñò
 
  MsgBox "Â ñèñòåìå íåò çàðåãèñòðèðîâàííûõ ïðîäóêòîâ. Îïåðàöèÿ ïðåðâàíà."
  
 End If
 
 qdfQuery.Close
 dbsCurrent.Close
 
 ControlView_ProductPage
 
End Sub

Private Sub SystemProductListForEdit_Click()
 
 ControlView_ProductPage
  
End Sub

Private Sub TabControlMainForm_Change()
 
 Dim TabControl As TabControl
 
 Set TabControl = Forms("MainForm").Controls("TabControlMainForm")
 
 Select Case TabControl.Value
 
  Case 0 'Àññîðòèìåíò
  
   ControlView_AssortmentPage
  
  Case 1 'SourceBalancePage
   
   ControlView_SourceBalancePage
  
  Case 2 'DayBalancePage
  
   ControlView_DayBalancePage
  
  Case 3 'ProductPage
  
   ControlView_ProductPage

  Case 4 'Ñìåñè
  
   ControlView_MixturesPage
  
  Case 5 'Îáîðóäîâàíèå
  
  Case 6 'Ãðàôèê ðàáîòû
  
  Case 7 'Íàñòðîéêè
  
  Case Else
  
 End Select
 
End Sub
Private Sub Command205_Click()
On Error GoTo Err_Command205_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command205_Click:
    Exit Sub

Err_Command205_Click:
    MsgBox Err.Description
    Resume Exit_Command205_Click
    
End Sub

