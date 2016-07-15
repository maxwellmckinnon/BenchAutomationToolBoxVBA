Attribute VB_Name = "ShutdownCurrent_vs_SupplyVolt"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

'''
'Master function
'any updates or improvements should be made here
'''
Sub getShutdownCurrent_vs_SupplyVoltage()
    'Step through the min to max voltage of the supply while measuring the current and voltage
    'FLUKE 8845A for current
    'HP 34401A DMM for voltage
    'Agilent 3631A for PSU
    'GPIB connections needed for PSU, current DMM, voltage DMM
    
    Dim V_DMM_GPIB As String: V_DMM_GPIB = "GPIB::01"
    Dim I_DMM_GPIB As String: I_DMM_GPIB = "GPIB::02"
    Dim PSU_GPIB As String: PSU_GPIB = "GPIB::06"
    Dim PSU_OUTPUT_TERMINAL As String: PSU_OUTPUT_TERMINAL = "P25V" '"P6V", "P25V"
    
    Dim V_SUPPLY_MIN As Double: V_SUPPLY_MIN = 1.6
    Dim V_SUPPLY_MAX As Double: V_SUPPLY_MAX = 3.65
    
    Dim STEPSIZE_mV As Integer: STEPSIZE_mV = 50 ' mV
    
    Dim v As Double
    Dim measV As Double
    Dim i As Double
    Dim dataSize As Double: dataSize = (V_SUPPLY_MAX - V_SUPPLY_MIN) * 1000# / STEPSIZE_mV + 1
    ReDim data_iv(dataSize, 2) As Double
    
    Cells(2, 1).value = "Voltage"
    Cells(2, 2).value = "Current"
    Dim n As Integer: n = 0
    
    'Manually grab 1.2V value
'    Call Power_Supply_E3631A_.Supply_Set_Output(PSU_GPIB, "P25V", 1.2)
'        Sleep (200)
'        i = Fluke_Meter.ReadAve_Fluke(I_DMM_GPIB)
'        Call DMM_34401A_.DMM_Get_Reading(V_DMM_GPIB, measV)
'        data_iv(n, 1) = measV
'        data_iv(n, 2) = i
'        Cells(n + 2, 1).value = measV
'        Cells(n + 2, 2).value = i
    
    n = 1
    v = V_SUPPLY_MIN
    Do While v <= V_SUPPLY_MAX + 0.001
        DoEvents
        Call Power_Supply_E3631A_.Supply_Set_Output(PSU_GPIB, PSU_OUTPUT_TERMINAL, v)
        Sleep (200)
        i = Fluke_Meter.ReadAve_Fluke(I_DMM_GPIB)
        Call DMM_34401A_.DMM_Get_Reading(V_DMM_GPIB, measV)
        data_iv(n, 1) = measV
        data_iv(n, 2) = i
        Cells(n + 2, 1).value = measV
        Cells(n + 2, 2).value = i
        n = n + 1
        v = v + STEPSIZE_mV * 1# / 1000
    Loop
    
End Sub

Sub getShutdownCurrent_vs_SupplyVoltage_DVDDIO_ExternalMCLK()
    'Step through the min to max voltage of the supply while measuring the current and voltage
    'Track the MCLK voltage as well
    '
    'FLUKE 8845A for current
    'HP 34401A DMM for voltage
    'Agilent 3631A for PSU
    'GPIB connections needed for PSU, current DMM, voltage DMM
    
    Dim V_DMM_GPIB As String: V_DMM_GPIB = "GPIB::01"
    Dim I_DMM_GPIB As String: I_DMM_GPIB = "GPIB::02"
    Dim PSU_GPIB As String: PSU_GPIB = "GPIB::06"
    Dim CLK_GPIB As String: CLK_GPIB = "GPIB::07"
    
    Dim PSU_OUTPUT_TERMINAL As String: PSU_OUTPUT_TERMINAL = "P25V" '"P6V", "P25V"
    
    Dim V_SUPPLY_MIN As Double: V_SUPPLY_MIN = 1.6
    Dim V_SUPPLY_MAX As Double: V_SUPPLY_MAX = 3.65
    
    Dim STEPSIZE_mV As Integer: STEPSIZE_mV = 50 ' mV
    
    Dim v As Double
    Dim measV As Double
    Dim i As Double
    Dim dataSize As Double: dataSize = (V_SUPPLY_MAX - V_SUPPLY_MIN) * 1000# / STEPSIZE_mV + 1
    ReDim data_iv(dataSize, 2) As Double
    
    Cells(2, 1).value = "Voltage"
    Cells(2, 2).value = "Current"
    Dim n As Integer: n = 0
    
    'Manually grab 1.2V value
'    Call Power_Supply_E3631A_.Supply_Set_Output(PSU_GPIB, "P25V", 1.2)
'        Sleep (200)
'        i = Fluke_Meter.ReadAve_Fluke(I_DMM_GPIB)
'        Call DMM_34401A_.DMM_Get_Reading(V_DMM_GPIB, measV)
'        data_iv(n, 1) = measV
'        data_iv(n, 2) = i
'        Cells(n + 2, 1).value = measV
'        Cells(n + 2, 2).value = i
    
    n = 1
    v = V_SUPPLY_MIN
    Do While v <= V_SUPPLY_MAX + 0.001
        DoEvents
        Call Power_Supply_E3631A_.Supply_Set_Output(PSU_GPIB, PSU_OUTPUT_TERMINAL, v)
        Call FuncGen_33250.Func_Gen_Set_Output(CLK_GPIB, v)
        
        Sleep (200)
        i = Fluke_Meter.ReadAve_Fluke(I_DMM_GPIB)
        Call DMM_34401A_.DMM_Get_Reading(V_DMM_GPIB, measV)
        data_iv(n, 1) = measV
        data_iv(n, 2) = i
        Cells(n + 2, 1).value = measV
        Cells(n + 2, 2).value = i
        n = n + 1
        v = v + STEPSIZE_mV * 1# / 1000
    Loop
    
    Call Power_Supply_E3631A_.Supply_Set_Output(PSU_GPIB, PSU_OUTPUT_TERMINAL, V_SUPPLY_MIN)
    Call FuncGen_33250.Func_Gen_Set_Output(CLK_GPIB, V_SUPPLY_MIN)
End Sub

'Functions below are based off of master above

Sub getShutdownCurrent_vs_SupplyVoltageDVDDIO()
    'Step through the min to max voltage of the supply while measuring the current and voltage
    'FLUKE 8845A for current
    'HP34401A DMM for votlage
    'GPIB connections needed for PSU, current DMM, voltage DMM
    
    Dim V_DMM_GPIB As String: V_DMM_GPIB = "GPIB::01"
    Dim I_DMM_GPIB As String: I_DMM_GPIB = "GPIB::02"
    Dim PSU_GPIB As String: PSU_GPIB = "GPIB::03"
    
    Dim V_SUPPLY_MIN As Double: V_SUPPLY_MIN = 1.65
    Dim V_SUPPLY_MAX As Double: V_SUPPLY_MAX = 3.63
    
    Dim STEPSIZE_mV As Integer: STEPSIZE_mV = 50 ' 10mV
    
    Dim v As Double
    Dim measV As Double
    Dim i As Double
    Dim dataSize As Double: dataSize = (V_SUPPLY_MAX - V_SUPPLY_MIN) * 1000# / STEPSIZE_mV + 1
    ReDim data_iv(dataSize, 2) As Double
    
    Cells(1, 1).value = "Voltage"
    Cells(1, 2).value = "Current"
    Dim n As Integer: n = 0
    
    n = 1
    v = V_SUPPLY_MIN
    Do While v <= V_SUPPLY_MAX + 0.001
        Call Power_Supply_E3631A_.Supply_Set_Output(PSU_GPIB, "P25V", v)
        Sleep (200)
        i = Fluke_Meter.ReadAve_Fluke(I_DMM_GPIB)
        Call DMM_34401A_.DMM_Get_Reading(V_DMM_GPIB, measV)
        data_iv(n, 1) = measV
        data_iv(n, 2) = i
        Cells(n + 2, 1).value = measV
        Cells(n + 2, 2).value = i
        n = n + 1
        v = v + STEPSIZE_mV * 1# / 1000
    Loop
    
    
    
End Sub

Sub newTest()

    Cells(1, 2) = "test"
End Sub

