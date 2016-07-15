VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Regulate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class for AP2700 regulation

Private pmin As Double
Private pmax As Double
Private pinputAmount As Double
Private ptarget As Double
Private ptolerance As Double
Private pRegulationType As Integer

Public Property Let min(value As Double)
    pmin = value
End Property

Public Property Let max(value As Double)
    pmax = value
End Property

Public Property Let inputAmount(value As Double)
    pinputAmount = value
End Property

Public Property Let target(value As Double)
    ptarget = value
End Property

Public Property Let tolerance(value As Double)
    ptolerance = value
End Property

Public Property Let RegulationType(value As Integer)
    pRegulationType = value ' 0 is linear, 1 is +normal
End Property

'''''''''''''''''''

Public Sub runAP2700internalRegulation()
    AP.Anlr.FuncMode = 4 ' THDN mode
    AP.Reg.TargetID = 5910 ' Anlr THDN
    AP.Reg.SourceID = 5106 ' DGEN Ampl A
    AP.Application.ThrowErrors = True ' Pass errors from regulation routine to VBA
    On Error Resume Next ' Ignore error
    
    AP.Reg.SourceOperation = 0 ' Linear
    AP.Reg.SourceHigh("dBFS") = pmax
    AP.Reg.SourceLow("dBFS") = pmin
    AP.Reg.TargetValue("dB") = ptarget
    AP.Reg.TargetTolerance("%") = ptolerance
    AP.Reg.SourceOperation = pRegulationType
    
    AP.Reg.Start
    
End Sub
