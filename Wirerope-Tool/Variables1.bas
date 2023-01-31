Attribute VB_Name = "Variables1"
Option Explicit

Public h As Double                               'Höhe
Public m As Double                               'Masse
Public v As Double                               'Geschwindigkeit

Public HSA As Double                             'HalbSinusbeschleunigung
Public t As Double                               'Impulszeit

Public TA As Double

Public E As Double                               'Energie

Public n As Integer                              'Anzahl der WRs
Public E_ges As Double                           'Gesamte Energie der n-ten Anzahl WRs

Public i, j As Integer
Public WRd As Double                             'WR Energie aus Datenbank Druck
Public WRs As Double                             'WRs Energie Share/Roll
Public lastcell As Double



Public Kv As Double                              'KV-Wert
Public Kv_ges As Double
Public Ks As Double                              'KS-Wert
Public Ks_ges As Double
Public statD As Double                           'static deflection [mm]
Public dynD As Double                            'dyn deflection [mm]
Public precomF As Double                         'Statische Vorspannkraft [pre-compression Force]
Public ShF As Double                             'Shock Force [N]
Public ReSha As Double                           'Rest Shock [m/s²]
Public ReShg As Double                           'Rest Shock [g]
Public NatFr As Double                           'Natural Frequency [Hz]
Public ShFr As Double                            'Shock Frequency [Hz]

Public € As Double
Public TP As Double



