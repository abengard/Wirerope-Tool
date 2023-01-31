VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WireRopeTool 
   Caption         =   "Wire Rope-Tool"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13250
   OleObjectBlob   =   "WireRopeTool.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "WireRopeTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButtonCancel_Click()

    Unload WireRopeTool

End Sub

Private Sub LabelHallo_Click()

    MsgBox "Hallo"


End Sub

'Bilder>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><


Private Sub OptionButtonCompression_Change()

    If Me.OptionButtonCompression = True Then
  
        Me.ImageCom.Visible = True
        Me.ImageShear.Visible = False
        Me.ImageRoll.Visible = False
        Me.ImageCompRoll.Visible = False
        Me.ImageTenRoll.Visible = False
    End If

End Sub

Private Sub OptionButtonShareRoll_Change()

    If Me.OptionButtonShareRoll = True Then
    
        Me.ImageCom.Visible = False
        Me.ImageShear.Visible = True
        Me.ImageRoll.Visible = True
        Me.ImageCompRoll.Visible = False
        Me.ImageTenRoll.Visible = False
    End If
    
End Sub

Private Sub OptionButtonCompRoll_Change()

    If Me.OptionButtonCompRoll = True Then
    
        Me.ImageCom.Visible = False
        Me.ImageShear.Visible = False
        Me.ImageRoll.Visible = False
        Me.ImageCompRoll.Visible = True
        Me.ImageTenRoll.Visible = False
    End If

End Sub

Private Sub OptionButtonTenRoll_Change()

    If Me.OptionButtonTenRoll = True Then
    
        Me.ImageCom.Visible = False
        Me.ImageShear.Visible = False
        Me.ImageRoll.Visible = False
        Me.ImageCompRoll.Visible = False
        Me.ImageTenRoll.Visible = True
        
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
End Sub

'Übersetzung Sprache>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub ComboBox2_Change()

    If ComboBox2.Value = "Deutsch" Then
    
        With WireRopeTool
            .FrameCase.Caption = "Fall"
            .FrameLD.Caption = "Belastungsrichtung"
    
            .LabelnWR.Caption = "Anzahl der Drahtseilfedern:"
            .LabelMain.Caption = "Fallhöhe [m]*:"
            .LabelVelo.Caption = "Geschwindigkeit [m/s]:"
            .LabelM.Caption = "Masse [kg]*:"
            .LabelField.Caption = "*Pflichtfeld"
            .LabelLanguage.Caption = "Sprache:"
    
            .OptionButtonCompression.Caption = "Druck"
            .OptionButtonfreeFall.Caption = "freier Fall"
            .OptionButtonShareRoll.Caption = "Abscherung/Rollen"
            .OptionButtonShock1.Caption = "Halbsinus Shock"
            .OptionButtonCompRoll.Caption = "45° Druck/Rollen"
            .OptionButtonTenRoll.Caption = "45° Zug/Rollen"
        
        
            If OptionButtonShock1 = True Then
    
                .LabelMain.Caption = "Impuls Zeit [s]*:"
                .LabelVelo.Caption = "Halb Sinus Beschleunigung [G]*:"
                
            ElseIf OptionButtonShock2 = True Then

                .LabelMain.Caption = "Impuls Zeit [s]*:"
                .LabelVelo.Caption = "Triangular-Beschleunigung [G]*:"
 
            End If
    
            .CommandButtonCancel.Caption = "Abbrechen"
            .CommandButtonRun.Caption = "Berechne"
        End With
    
        Me.ImageENG.Visible = False
        Me.ImageDE.Visible = True
    

    ElseIf ComboBox2.Value = "English" Then

        With WireRopeTool
            .FrameCase.Caption = "Case"
            .FrameLD.Caption = "Load Direction"
    
            .LabelnWR.Caption = "number of WRs:"
            .LabelMain.Caption = "Drop Height [m]*:"
            .LabelVelo.Caption = "Velocity [m/s]:"
            .LabelM.Caption = "Mass [kg]*:"
            .LabelField.Caption = "*Mandatory Field"
            .LabelLanguage.Caption = "Language:"
        
            If OptionButtonShock1 = True Then
    
                .LabelMain.Caption = "Impulse Time [s]*:"
                .LabelVelo.Caption = "Half Sine Acceleration [G]*:"

            ElseIf OptionButtonShock2 = True Then
        
                .LabelMain.Caption = "Impulse Time [s]*:"
                .LabelVelo.Caption = "Triangular Acceleration [G]*:"

            End If
    
            .OptionButtonCompression.Caption = "Compression"
            .OptionButtonfreeFall.Caption = "free Fall"
            .OptionButtonShareRoll.Caption = "Share/Roll"
            .OptionButtonShock1.Caption = "Half Sine Shock"
            .OptionButtonTenRoll.Caption = "45° Tension/Roll"

            .CommandButtonCancel.Caption = "Cancel"
            .CommandButtonRun.Caption = "Run"
        End With
    
        Me.ImageENG.Visible = True
        Me.ImageDE.Visible = False
    
    End If
    
End Sub

Private Sub OptionButtonShock1_Change()

    If ComboBox2 = "English" Then

        If OptionButtonShock1 = True Then
    
            With WireRopeTool
                .LabelMain.Caption = "Impulse Time [s]*:"
                .LabelVelo.Caption = "Half Sine Acceleration [G]*:"

            End With

        End If
    
    ElseIf ComboBox2 = "Deutsch" Then

        If OptionButtonShock1 = True Then
    
            With WireRopeTool
                .LabelMain.Caption = "Impuls Zeit [s]*:"
                .LabelVelo.Caption = "Halb Sinus Beschleunigung [G]*:"
            End With

        End If
    End If

End Sub

Private Sub OptionButtonShock2_Change()

    If ComboBox2 = "English" Then

        If OptionButtonShock2 = True Then
    
            With WireRopeTool
                .LabelMain.Caption = "Impulse Time [s]*:"
                .LabelVelo.Caption = "Triangular Acceleration [G]*:"
            End With

        End If

    ElseIf ComboBox2 = "Deutsch" Then

        If OptionButtonShock2 = True Then
    
            With WireRopeTool
                .LabelMain.Caption = "Impuls Zeit [s]*:"
                .LabelVelo.Caption = "Triangular-Beschleunigung [G]*:"
            End With

        End If

    End If

End Sub

Private Sub OptionbuttonfreeFall_Change()

    If ComboBox2 = "English" Then

        If OptionButtonfreeFall = True Then
    
            With WireRopeTool
                .LabelMain.Caption = "Drop Height [m]*:"
                .LabelVelo.Caption = "Velocity [m/s]:"
            End With
    
        End If

    ElseIf ComboBox2 = "Deutsch" Then
    
        If OptionButtonfreeFall = True Then
    
            With WireRopeTool
                .LabelMain.Caption = "Fallhöhe [m]*:"
                .LabelVelo.Caption = "Geschwindigkeit [m/s]:"
            End With
    
        End If
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
End Sub

'Textbox keine Buchstaben Zulassen >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub TxtBoxDropheight_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If InStr("1234567890.," & Chr$(8), Chr$(KeyAscii)) = 0 Then
  
        KeyAscii = 0
    End If

    If KeyAscii = Asc(".") Then
        KeyAscii = Asc(",")
    End If

    If InStr(1, TxtboxDropheight.Text, ",", 0) And KeyAscii = Asc(",") Then
  
        KeyAscii = 0
    End If
   
End Sub

Private Sub TextBoxVelocity_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If InStr("1234567890.," & Chr$(8), Chr$(KeyAscii)) = 0 Then
  
        KeyAscii = 0
    End If

    If KeyAscii = Asc(".") Then
        KeyAscii = Asc(",")
    End If

    If InStr(1, TextboxVelocity.Text, ",", 0) And KeyAscii = Asc(",") Then
  
        KeyAscii = 0
    End If
   
End Sub

Private Sub TxtBoxHSA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If InStr("1234567890.," & Chr$(8), Chr$(KeyAscii)) = 0 Then
  
        KeyAscii = 0
    End If

    If KeyAscii = Asc(".") Then
        KeyAscii = Asc(",")
    End If

    If InStr(1, TxtboxHSA.Text, ",", 0) And KeyAscii = Asc(",") Then
  
        KeyAscii = 0
    End If
   
End Sub

Private Sub TxtBoxMass_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If InStr("1234567890.," & Chr$(8), Chr$(KeyAscii)) = 0 Then
  
        KeyAscii = 0
    End If

    If KeyAscii = Asc(".") Then
        KeyAscii = Asc(",")
    End If

    If InStr(1, TxtboxMass.Text, ",", 0) And KeyAscii = Asc(",") Then
  
        KeyAscii = 0
    End If
   
End Sub

Private Sub Combobox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If InStr("1234567890" & Chr$(8), Chr$(KeyAscii)) = 0 Then
  
        KeyAscii = 0
    End If
   
End Sub

Private Sub Combobox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    If InStr("" & Chr$(8), Chr$(KeyAscii)) = 0 Then
  
        KeyAscii = 0
    End If
   
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub Userform_Initialize()

    Dim i As Integer

    With WireRopeTool.ComboBox1
        For i = 1 To 50
            .AddItem CStr(i)
        Next i
    
        .Value = 1


    End With

    With WireRopeTool.ComboBox2
        .AddItem ("English")
        .AddItem ("Deutsch")
    
        .Value = "English"
    
    End With

    Me.OptionButtonfreeFall.Value = 1
    Me.OptionButtonCompression.Value = 1

End Sub

Public Sub CommandButtonRun_Click()

    'Leeren der Charts>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    With ThisWorkbook.Worksheets("ChartCalculation").Activate
        Rows("2:65536").Select
        Selection.ClearContents
    End With

    With ThisWorkbook.Worksheets("ChartComparison").Activate
        Rows("2:65536").Select
        Selection.ClearContents
    End With

    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    'FALL FREIER FALL >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    If OptionButtonfreeFall = True Then


        'Wenn die Fallhöhe und die Masse nicht angegeben ist --> Warnung.Textbox andernfalls beginnt die Rechnung>>>>>>>>>>>>>>>>>>>>>>>
    
        If TxtboxDropheight.Value = "" Or TxtboxMass.Value = "" Then
    
            If ComboBox2 = "English" Then
    
                MsgBox "Please fill out the mandatory fields!", vbInformation + vbOKOnly, "Hint!"
      
            ElseIf ComboBox2 = "Deutsch" Then
    
                MsgBox "Bitte füllen Sie alle Pflichtfelder aus!", vbInformation + vbOKOnly, "Hint"
         
            End If
    
        Else
    

            h = TxtboxDropheight.Value
            m = TxtboxMass.Value
            n = ComboBox1.Value

            With WireRopeTool

                If TextboxVelocity.Value = "" Then

                    v = (2 * 9.81 * h) ^ 0.5
                    E = (0.5 * m * (v) ^ 2) / n
    
                Else
        
                    v = TextboxVelocity.Value
                    E = (0.5 * m * (v) ^ 2) / n
        
                End If
            End With

            E_ges = E
    
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            'Wenn Druck oder Share/Roll soll die gesamte Energeie mit dazugehörigen Chart verglichen die passenden in eine seperate Tabelle abgespeichert>>>>>>>>>>>>>>>>>>>>>>

            If OptionButtonCompression = True Then 'Druck>>>>>>>>>>>>>>>>>>
                                                   
                With Sheets("DatabaseCompression")
        
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("DatabaseCompression").Cells(i, j).Value
                            Next j
                    
                        Else
                
                            WRd = 0
                     
                        End If
                    Next i
                End With
            ElseIf OptionButtonShareRoll = True Then 'Abscherung/Rollen>>>>>>>>>>>>>>>>>>>>>
     
                With Sheets("DatabaseShareRoll")
           
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("DatabaseShareRoll").Cells(i, j).Value
                            Next j
                        Else
                            WRs = 0
            
                        End If
                    Next i
                End With
            ElseIf OptionButtonCompRoll = True Then 'Druck/Rollen>>>>>>>>>>>>>>>>>
     
                With Sheets("Database45°CompRoll")
           
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("Database45°CompRoll").Cells(i, j).Value
                            Next j
                        Else
                            WRs = 0
            
                        End If
                    Next i
                End With
            ElseIf OptionButtonTenRoll = True Then 'Zug/Rollen>>>>>>>>>>>>>
     
                With Sheets("Database45°TensionRoll")
           
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("Database45°TensionRoll").Cells(i, j).Value
                            Next j
                        Else
                            WRs = 0
            
                        End If
                    Next i
                End With
            Else

                MsgBox "Please choose a LOAD DIRECTION!", vbCritical + vbOKOnly, "Achtung!"
   
            End If
 
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            'Berechnung der restlichen Parameter und Übertragung der Daten auf weitere Tabelle>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

            Call Calculation

            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            'Userform 2 zeigt die Erebnis >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

            If OptionButtonCompression = True Or OptionButtonShareRoll = True Or OptionButtonCompRoll = True Or OptionButtonTenRoll = True Then

                With UserformFreeFall
                    .LabelMass.Caption = m
                    .LabelnumberofWRs.Caption = n
                    .LabelDropheight.Caption = h
                    .LabelVelocity.Caption = Round(v, (2))
                    .LabelE.Caption = Round(E, (1))
                
                    With UserformFreeFall.ListBoxpossibleWireRopes
                        Worksheets("ChartComparison").Activate
            
                        .MultiSelect = fmMultiSelectSingle
                        .ListStyle = fmListStyleOption
                
                    End With
                
                    If ComboBox2.Value = "English" Then
       
                        If OptionButtonCompression = True Then

                            .LabelLoadDirection = "Compression"
    
                        ElseIf OptionButtonShareRoll = True Then
    
                            .LabelLoadDirection.Caption = "Share/Roll"
                
                        ElseIf OptionButtonCompRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Compression/Roll"
                
                        ElseIf OptionButtonTenRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Tension/Roll"
        
                        End If
        
                
                    ElseIf ComboBox2.Value = "Deutsch" Then
                        If OptionButtonCompression = True Then

                            .LabelLoadDirection = "Druck"
    
                        ElseIf OptionButtonShareRoll = True Then
    
                            .LabelLoadDirection.Caption = "Abscherung/Rollen"
                                
                        ElseIf OptionButtonCompRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Druck/Rollen"
                
                        ElseIf OptionButtonTenRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Zug/Rollen"
        
                        End If
                
                    End If
            
                
                    For i = 2 To 100
                        If Sheets("ChartComparison").Cells(i, 1) <> "" Then
                        
                        
                            .ListBoxpossibleWireRopes.AddItem Sheets("ChartComparison").Cells(i, 1)
                        End If
                    
                    Next i
       
                    .Show
                End With
            End If
        End If
    
    
    
    
    
    
    
    

        'FALL HSA SHOCK >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><

    ElseIf OptionButtonShock1 = True Then

        If TxtboxDropheight.Value = "" Or TxtboxMass.Value = "" Or TextboxVelocity = "" Then

            If ComboBox2 = "English" Then
    
                MsgBox "Please fill out the mandatory fields!", vbInformation + vbOKOnly, "Hint!"
      
            ElseIf ComboBox2 = "Deutsch" Then
    
                MsgBox "Bitte füllen Sie alle Pflichtfelder aus!", vbInformation + vbOKOnly, "Hint"
         
            End If

        Else

            t = TxtboxDropheight.Value
            m = TxtboxMass.Value
            HSA = TextboxVelocity.Value
            n = ComboBox1.Value

            v = (2 * 9.81 / 3.14) * HSA * t
            E = (0.5 * m * (v) ^ 2) / n
    


            E_ges = E
    
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            'Wenn Druck oder Share/Roll soll die gesamte Energeie mit dazugehörigen Chart verglichen die passenden in eine seperate Tabelle abgespeichert>>>>>>>>>>>>>>>>>>>>>>

            If OptionButtonCompression = True Then
                                                   
                With Sheets("DatabaseCompression")
        
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("DatabaseCompression").Cells(i, j).Value
                            Next j
                    
                        Else
                
                            WRd = 0
                     
                        End If
                    Next i
                End With
            ElseIf OptionButtonShareRoll = True Then
     
                With Sheets("DatabaseShareRoll")
           
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("DatabaseShareRoll").Cells(i, j).Value
                            Next j
                        Else
                            WRs = 0
            
                        End If
                    Next i
                End With
            ElseIf OptionButtonCompRoll = True Then
     
                With Sheets("Database45°CompRoll")
           
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("Database45°CompRoll").Cells(i, j).Value
                            Next j
                        Else
                            WRs = 0
            
                        End If
                    Next i
                End With
            ElseIf OptionButtonTenRoll = True Then 'Zug/Rollen>>>>>>>>>>>>>
     
                With Sheets("Database45°TensionRoll")
           
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("Database45°TensionRoll").Cells(i, j).Value
                            Next j
                        Else
                            WRs = 0
            
                        End If
                    Next i
                End With
            End If
 
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            'Berechnung der restlichen Parameter und Übertragung der Daten auf weitere Tabelle>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

            Call Calculation

            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            'Userform 2 zeigt die Erebnis >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

            If OptionButtonCompression = True Or OptionButtonShareRoll = True Or OptionButtonCompRoll = True Or OptionButtonTenRoll = True Then

                With UserFormHSShock
                    .LabelMass.Caption = m
                    .LabelnumberofWRs.Caption = n
                    .LabelDropheight.Caption = t
                    .LabelVelocity.Caption = Round(v, (2))
                    .LabelE.Caption = Round(E, (1))
                    .LabelHSA.Caption = HSA
                
                    With UserFormHSShock.ListBoxpossibleWireRopes
                        Worksheets("ChartComparison").Activate
            
                        .MultiSelect = fmMultiSelectSingle
                        .ListStyle = fmListStyleOption
                    End With
        
                    If ComboBox2.Value = "English" Then
       
                        If OptionButtonCompression = True Then

                            .LabelLoadDirection = "Compression"
    
                        ElseIf OptionButtonShareRoll = True Then
    
                            .LabelLoadDirection.Caption = "Share/Roll"
                
                        ElseIf OptionButtonCompRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Compression/Roll"
                    
                        ElseIf OptionButtonTenRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Tension/Roll"
        
                        End If
        
                
                    ElseIf ComboBox2.Value = "Deutsch" Then
                        If OptionButtonCompression = True Then

                            .LabelLoadDirection = "Druck"
    
                        ElseIf OptionButtonShareRoll = True Then
    
                            .LabelLoadDirection.Caption = "Abscherung/Rollen"
                
                        ElseIf OptionButtonCompRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Druck/Rollen"
                    
                        ElseIf OptionButtonTenRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Zug/Rollen"
        
                        End If
                    End If
        
                    For i = 2 To 100
                        If Sheets("ChartComparison").Cells(i, 1) <> "" Then
                
                            .ListBoxpossibleWireRopes.AddItem Sheets("ChartComparison").Cells(i, 1)
                            .ListBoxpossibleWireRopes.List(.ListBoxpossibleWireRopes.ListCount - 1, 1) = Round(Cells(i, 5), 1)
            
                        End If
                    Next i
       
                    .Show
                End With
            End If
        End If
    
    
    
    
    
    
    
    
    
    
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        'FALL TSHOCK>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    ElseIf OptionButtonShock2 = True Then

        If TxtboxDropheight.Value = "" Or TxtboxMass.Value = "" Or TextboxVelocity = "" Then

            If ComboBox2 = "English" Then
    
                MsgBox "Please fill out the mandatory fields!", vbInformation + vbOKOnly, "Hint!"
      
            ElseIf ComboBox2 = "Deutsch" Then
    
                MsgBox "Bitte füllen Sie alle Pflichtfelder aus!", vbInformation + vbOKOnly, "Hint"
         
            End If

        Else

            t = TxtboxDropheight.Value
            m = TxtboxMass.Value
            TA = TextboxVelocity.Value
            n = ComboBox1.Value

            v = (9.81 / 2) * TA * t
            E = (0.5 * m * (v) ^ 2) / n


            E_ges = E
    
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            'Wenn Druck oder Share/Roll soll die gesamte Energeie mit dazugehörigen Chart verglichen die passenden in eine seperate Tabelle abgespeichert>>>>>>>>>>>>>>>>>>>>>>

            If OptionButtonCompression = True Then
                                                   
                With Sheets("DatabaseCompression")
        
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("DatabaseCompression").Cells(i, j).Value
                            Next j
                    
                        Else
                
                            WRd = 0
                     
                        End If
                    Next i
                End With
            ElseIf OptionButtonShareRoll = True Then
     
                With Sheets("DatabaseShareRoll")
           
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("DatabaseShareRoll").Cells(i, j).Value
                            Next j
                        Else
                            WRs = 0
            
                        End If
                    Next i
                End With
            ElseIf OptionButtonCompRoll = True Then
     
                With Sheets("Database45°CompRoll")
           
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("Database45°CompRoll").Cells(i, j).Value
                            Next j
                        Else
                            WRs = 0
            
                        End If
                    Next i
                End With
            ElseIf OptionButtonTenRoll = True Then 'Zug/Rollen>>>>>>>>>>>>>
     
                With Sheets("Database45°TensionRoll")
           
                    For i = 2 To 100
                        WRs = .Cells(i, 5).Value
                        lastcell = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
            
                        If WRs > E_ges Then
                            For j = 1 To 5
                                Sheets("ChartComparison").Cells(i, j).Value = Sheets("Database45°TensionRoll").Cells(i, j).Value
                            Next j
                        Else
                            WRs = 0
            
                        End If
                    Next i
                End With
            Else

                MsgBox "Please choose a LOAD DIRECTION!", vbCritical + vbOKOnly, "Achtung!"
   
            End If
 
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            'Berechnung der restlichen Parameter und Übertragung der Daten auf weitere Tabelle>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

            Call Calculation

            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            'Userform 2 zeigt die Erebnis >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

            If OptionButtonCompression = True Or OptionButtonShareRoll = True Or OptionButtonCompRoll = True Or OptionButtonTenRoll = True Then

                With UserFormTShock
                    .LabelMass.Caption = m
                    .LabelnumberofWRs.Caption = n
                    .LabelDropheight.Caption = t
                    .LabelVelocity.Caption = Round(v, (2))
                    .LabelE.Caption = Round(E, (1))
                    .LabelTA.Caption = TA
                
                    With UserFormTShock.ListBoxpossibleWireRopes
                        Worksheets("ChartComparison").Activate
            
                        .MultiSelect = fmMultiSelectSingle
                        .ListStyle = fmListStyleOption
                    End With
        
                    If ComboBox2.Value = "English" Then
       
                        If OptionButtonCompression = True Then

                            .LabelLoadDirection = "Compression"
    
                        ElseIf OptionButtonShareRoll = True Then
    
                            .LabelLoadDirection.Caption = "Share/Roll"
                
                        ElseIf OptionButtonCompRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Compression/Roll"
                    
                        ElseIf OptionButtonTenRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Tension/Roll"
        
                        End If
          
                    ElseIf ComboBox2.Value = "Deutsch" Then
                        If OptionButtonCompression = True Then

                            .LabelLoadDirection = "Druck"
    
                        ElseIf OptionButtonShareRoll = True Then
    
                            .LabelLoadDirection.Caption = "Abscherung/Rollen"
                
                        ElseIf OptionButtonCompRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Druck/Rollen"
                    
                        ElseIf OptionButtonCompRoll = True Then
    
                            .LabelLoadDirection.Caption = "45° Zug/Rollen"
        
                        End If
                    End If
            
                    For i = 2 To 100
                        If Sheets("ChartComparison").Cells(i, 1) <> "" Then
                
                            .ListBoxpossibleWireRopes.AddItem Sheets("ChartComparison").Cells(i, 1)
                            .ListBoxpossibleWireRopes.List(.ListBoxpossibleWireRopes.ListCount - 1, 1) = Round(Cells(i, 5), 1)
            
                        End If
                    Next i
       
                    .Show
                End With
            End If
        End If

    Else

        MsgBox "Please choose a CASE!", vbCritical + vbOKOnly, "Achtung!"

    End If



End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
