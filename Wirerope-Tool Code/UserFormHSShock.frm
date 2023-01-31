VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormHSShock 
   Caption         =   "Half sine Shock Results"
   ClientHeight    =   9590.001
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   11860
   OleObjectBlob   =   "UserFormHSShock.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormHSShock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Back_Click()

    Unload UserFormHSShock
    Load WireRopeTool

End Sub

Private Sub Cancel_Click()

    Unload UserFormHSShock
    Unload WireRopeTool

End Sub

Private Sub CommandButton1_Click()

    Const PDF_PATH = "https://www.ace-ace.de/media/acedownloads/ACE_WireRopes_Catalogue_EN_20200729.pdf"

    Select Case Left(Me.ListBoxpossibleWireRopes.Text, 4)

    Case "WR2-"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=4"
    
    Case "WR3-"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=8"

    Case "WR4-"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=12"
    
    Case "WR5-"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=16"
    
    Case "WR6-"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=24"
    
    Case "WR8-"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=28"
    
    Case "WR10"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=32"
    
    Case "WR12"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=40"
    
    Case "WR16"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=48"
    
    Case "WR20"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=52"
    
    Case "WR24"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=56"
    
    Case "WR28"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=60"
    
    Case "WR32"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=64"
    
    Case "WR36"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=68"
    
    Case "WR40"
        ThisWorkbook.FollowHyperlink PDF_PATH & "#page=72"
    
    End Select


End Sub

Private Sub ListBoxpossibleWireRopes_Change()


    With Tabelle3

        For i = 1 To 100

            If Me.ListBoxpossibleWireRopes.Text = .Cells(i, 1) Then
        
                Sheets("ChartComparison").Activate
    
                Me.LabelKS.Caption = Cells(i, 3).Value
                Me.LabelKV.Caption = Cells(i, 2).Value
                Me.LabelEne.Caption = Round(Cells(i, 5).Value, 1)
        
                Sheets("ChartCalculation").Activate
                Me.LabelShF.Caption = Round(Cells(i, 6).Value, 1)
                Me.LabelReSha.Caption = Round(Cells(i, 7).Value, 1)
                Me.LabelReShg.Caption = Round(Cells(i, 8).Value, 1)
                Me.LabelNatFr.Caption = Round(Cells(i, 9).Value, 1)
                Me.LabelShFr.Caption = Round(Cells(i, 10).Value, 1)
                Me.LabelstaD.Caption = Round(Cells(i, 3).Value, 1)
                Me.LabeldynaD.Caption = Round(Cells(i, 5).Value, 1)
        
                Sheets("DatabasePrice").Activate
                If n >= 1 And n <= 10 Then
        
                    Me.LabelTP€.Caption = Cells(i, 2).Value * n
                    Me.LabelPWR€.Caption = Cells(i, 2).Value
            
                ElseIf n >= 11 And n <= 20 Then
            
                    Me.LabelTP€.Caption = Cells(i, 3).Value * n
                    Me.LabelPWR€.Caption = Cells(i, 3).Value
            
                ElseIf n >= 21 And n <= 30 Then
        
                    Me.LabelTP€.Caption = Cells(i, 4).Value * n
                    Me.LabelPWR€.Caption = Cells(i, 4).Value
        
                ElseIf n >= 31 And n <= 40 Then
            
                    Me.LabelTP€.Caption = Cells(i, 5).Value * n
                    Me.LabelPWR€.Caption = Cells(i, 5).Value
        
                ElseIf n >= 41 And n <= 50 Then
            
                    Me.LabelTP€.Caption = Cells(i, 6).Value * n
                    Me.LabelPWR€.Caption = Cells(i, 6).Value
                End If

            End If
       
       
    
        Next i
    
    End With
End Sub

Private Sub Userform_Initialize()

    With WireRopeTool

        If .ComboBox2 = "Deutsch" Then


            With UserFormHSShock
                .Frame1.Caption = "Gesucht"
                .Frame4.Caption = "mögliche Drahtseilfedern"
        
                .LabelLD.Caption = "Belastungsrichtung:"
                .LabelnWR.Caption = "Anzahl der Federn:"
                .LabelIT.Caption = "Impuls Zeit [s]:"
                .LabelSV.Caption = "Shock Geschwindigkeit [m/s]:"
                .LabelM.Caption = "Masse [kg]:"
                .Labelchoose.Caption = "*wähle eine Drahtseilfeder aus um die Details zu sehen"
                .LabelcE.Caption = "berechnete Energie [Nm]:"
                .LabelEpWR.Caption = "Energie pro Feder [Nm]:"
                .LabelKSWR.Caption = "KS pro Feder [N/m]:"
                .LabelKVWR.Caption = "KV pro Feder [N/m]:"
                .LabelSF.Caption = "Shock Kraft [N]:"
                .LabelRS.Caption = "Rest Shock [m/s^2]:"
                .LabelRSg.Caption = "Rest Shock [g]:"
                .LabelNF.Caption = "Natürliche Frequenz [Hz]:"
                .LabelS.Caption = "Shock Frequenz [Hz]:"
                .LabelHA.Caption = "Halbsinusbeschleunigung [G]:"
                .LabelPWR.Caption = "Stückpreis [€]:"
                .LabelTP.Caption = "Gesamtpreis [€]:"
                .LabelstatD.Caption = "stat. Einfederung [mm]:"
                .LabeldynD.Caption = "dyn. Einfederung [mm]:"
        
                .Back.Caption = "<<Zurück"
                .Cancel.Caption = "Abbrechen"
        
            End With
    
        ElseIf .ComboBox2 = "English" Then

            With UserFormHSShock
                .Frame1.Caption = "Wanted"
                .Frame4.Caption = "possible Wire Ropes"
        
                .LabelLD.Caption = "Load Direction:"
                .LabelnWR.Caption = "number of WRs:"
                .LabelIT.Caption = "Impulse Time [s]:"
                .LabelSV.Caption = "Shock Velocity [m/s]:"
                .LabelM.Caption = "Mass [kg]:"
                .Labelchoose.Caption = "*choose a Wire Rope to see Details"
                .LabelcE.Caption = "calculated Energy [Nm]:"
                .LabelEpWR.Caption = "Energy per WR [Nm]:"
                .LabelKSWR.Caption = "KS per WR [N/m]:"
                .LabelKVWR.Caption = "KV per WR [N/m]:"
                .LabelSF.Caption = "Shock Force [N]:"
                .LabelRS.Caption = "Rest Shock [m/s^2]:"
                .LabelRSg.Caption = "Rest Shock [g]:"
                .LabelNF.Caption = "Natural Frequency[Hz]:"
                .LabelS.Caption = "Shock Frequency[Hz]:"
                .LabelPWR.Caption = "Price per WR [€]:"
                .LabelTP.Caption = "Total Price [€]:"
                .LabelstatD.Caption = "stat. Deflection [mm]:"
                .LabeldynD.Caption = "dyn. Deflection [mm]:"
        
            End With
    
        End If
    End With
End Sub

