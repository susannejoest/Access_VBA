Public Function fct_Word_Check_Checkboxes(strTagName As String)
' Word document has checkboxes that are named e.g. "DS1"
        Set wdDoc = wdAppPUBLIC.Documents.Open(strPathFileNameOutput, , False)
            'Selection.ParentContentControl.Checked = False
            'wdDoc.FormFields("DS1").CheckBox.Value = True
            'wdDoc.ContentControls("DS1").Checked = True
            With wdDoc.SelectContentControlsByTitle(strTagName) ' e.g. "DS1"or SelectContentControlsByTitle
              For i = 1 To .Count
                With .Item(i)
                  If .Checked = False Then
                    .Checked = True
                    Exit For
                  End If
                End With
              Next
            End With
      
          Set rsFields = CurrentDb.OpenRecordset("qry_Fields") 'Table of tagged fields with attributes
              
              Do Until rs.EOF
                strFldCheckBoxName = rsFields!FldCheckBoxName
                strFldNameExcel = rsFields!FldNameExcel
          
                    For Each ctl In wdDoc.ContentControls
                        If ctl.Type = 8 Then ' Is a check box
                            If ctl.Title = strFldCheckBoxName Then
                                ctl.Checked = True
                            End If
                        End If
                    Next ctl
                    Set ctl = Nothing
          
                rsFields.MoveNext
                Loop
        Set ctl = Nothing
        
End Function
