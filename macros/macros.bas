Sub AddDropDownAuthorGender1()
  Dim oCC As ContentControl
  Set oCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList, Selection.Range)
  With oCC
    .DropdownListEntries.Add "vypracoval/a"
    .DropdownListEntries.Add "vypracoval"
    .DropdownListEntries.Add "vypracovala"
  End With
  Set oCC = Nothing
End Sub

Sub AddDropDownAuthorGender2()
  Dim oCC As ContentControl
  Set oCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList, Selection.Range)
  With oCC
    .DropdownListEntries.Add "rigorozant"
    .DropdownListEntries.Add "rigorozantka"
    .DropdownListEntries.Add "diplomant"
    .DropdownListEntries.Add "diplomantka"
    .DropdownListEntries.Add "disertant"
    .DropdownListEntries.Add "disertantka"
  End With
  Set oCC = Nothing
End Sub


Sub AddDropDownWorkType1()
  Dim oCC As ContentControl
  Set oCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList, Selection.Range)
  With oCC
    .DropdownListEntries.Add "Rigorózní"
    .DropdownListEntries.Add "Diplomová"
    .DropdownListEntries.Add "Diserta" + ChrW(&H10D) + "ní"
  End With
  Set oCC = Nothing
End Sub

Sub AddDropDownWorkType2()
  Dim oCC As ContentControl
  Set oCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList, Selection.Range)
  With oCC
    .DropdownListEntries.Add "rigorózní"
    .DropdownListEntries.Add "diplomovou"
    .DropdownListEntries.Add "diserta" + ChrW(&H10D) + "ní"
  End With
  Set oCC = Nothing
End Sub


Sub AddDropDownSupervisor()
  Dim oCC As ContentControl
  Set oCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList, Selection.Range)
  With oCC
    .DropdownListEntries.Add "Pov" + ChrW(&H11B) + ChrW(&H159) + "en" + ChrW(&HFD) + " akademick" + ChrW(&HFD) + " pracovník"
    .DropdownListEntries.Add "Vedoucí diplomové práce"
    .DropdownListEntries.Add ChrW(&H160) + "kolitel"
  End With
  Set oCC = Nothing
End Sub


Sub AddDropDownStudyField()
  Dim oCC As ContentControl
  Set oCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList, Selection.Range)
  With oCC
    .DropdownListEntries.Add "Tematick" + ChrW(&HFD) + " okruh"
    .DropdownListEntries.Add "Katedra"
    .DropdownListEntries.Add "Studijní program"
  End With
  Set oCC = Nothing
End Sub
