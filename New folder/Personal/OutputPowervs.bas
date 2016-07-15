Attribute VB_Name = "OutputPowervs"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub MeasureLoadValue()
    'Dump the measured loads to excel
    'Do this with the class D amplifier off and with all effective trace lengths in place
    Call LoadEVKITFile_I2CBridge_16bit("C:\Users\maxwell.mckinnon\Dropbox\Maxim\ICs and Data\AX80\Bench\Rev A1\2584 Output Power vs Load\Output Power vs Load Board 403A.98507t", &H74) ' reprogram part
    Call GlobalDisable(&H74)
    Dim i As Integer
    Dim j As Integer
    Dim loadBoard1Addr As Integer: loadBoard1Addr = &HB0
    Dim loadBoard2Addr As Integer: loadBoard2Addr = &H4E
    
    For j = 0 To 7 'cycle through second load board
        Call setupLoad(loadBoard2Addr, j)
        For i = 0 To 7 ' cycle through first load board
            DoEvents
            Call setupLoad(loadBoard1Addr, i)
            Sleep (500)
            LoadValue = MeasureLoad()
            
            Cells(3 + i + j * 8, 1) = i
            Cells(3 + i + j * 8, 2).value = LoadValue
        Next i
    Next j

End Sub

Sub setupLoad(addr As Integer, i As Integer)
    Call Load_Board.LoadboardHex(addr, i)
End Sub

Function MeasureLoad(Optional Load_GPIB As String = "GPIB::11") As Double
    Dim ohm As Double
    Call Equipment_GPIB.DMM_34401A_.DMM_Get_Reading(Load_GPIB, ohm)
    MeasureLoad = ohm
End Function

Sub OutputPowervsLoad()
    'Run at 1% THDN and 10% THDN, VBAT 2.5, 3.7, 4.3, sweep from 3ohm to 100ohm + 33uH
    
    epsilon = 2 ' Error if more than 2dB off in THDN
    MAX_RETRIES = 3
    Dim max As Double: max = 0
    Dim min As Double: min = -10 'Minimum for 1% too
    Dim THDNLEVELS(1) As Double: THDNLEVELS(0) = -40: THDNLEVELS(1) = -20 '1%, 10%
    Dim thdnlevel As Double
    Dim tolerance As Double: tolerance = 0.5 ' percent
    Dim VBATs(2) As Double: VBATs(0) = 2.5: VBATs(1) = 3.7: VBATs(2) = 4.3
    VBAT_GPIB = "GPIB::01"
    
    Dim i As Integer
    Dim j As Integer
    Dim loadBoard1Addr As Integer: loadBoard1Addr = &HB0
    Dim loadBoard2Addr As Integer: loadBoard2Addr = &H4E
    
    For Each VBAT In VBATs
        Call Equipment_GPIB.Power_Supply_E3631A_.Supply_Set_Output(VBAT_GPIB, "P6V", VBAT, 5)
        
        ActiveWorkbook.Save
        For Each thdnlev In THDNLEVELS
            Sheets.Add
            ActiveSheet.Name = ActiveSheet.Name & " " & Str(thdnlev) & " " & Str(VBAT)
            thdnlevel = thdnlev
            For j = 0 To 7
                
                Call setupLoad(loadBoard2Addr, j)
                For i = 0 To 7
                    DoEvents
                    If (j > 2 And (i = 3 Or i = 5 Or i = 6 Or i = 7)) Or (j > 3 And Not (i = 0 Or i = 1 Or i = 4)) Then ' eliminate repeat useless data BOARD SPECIFIC
                        GoTo ContinueLoop
                    End If
                    Call setupLoad(loadBoard1Addr, i)
                    Sleep (500)
                    retryCount = 0
                    
                    ''''BEGIN THDN REGULATION''''
                    Do
                        DoEvents
                        Call RegulateTHDN(min, max, -4, thdnlevel, tolerance)
                        'check THDN level
                        AP.Anlr.FuncMode = 4 ' Switch to THDN mode
                        Var = AP.Anlr.FuncRdg("dB")
                        retryCount = retryCount + 1
                        Call LoadEVKITFile_I2CBridge_16bit("C:\Users\maxwell.mckinnon\Dropbox\Maxim\ICs and Data\AX80\Bench\Rev A1\2584 Output Power vs Load\Output Power vs Load Board 403A_8V.98507t", &H74) ' reprogram part
                    Loop While (Abs(Var - thdnlevel) > epsilon) And retryCount < MAX_RETRIES
                    If retryCount >= MAX_RETRIES Then
                        Debug.Print ("Max Retry limit reached on THDN: " & Str(thdnlevel) & ", Iteration i: " & Str(i))
                        
                    End If
                    ''''END THDN REGULATION''''
                    THDN = Var
                    
                    AP.Anlr.FuncMode = 0
                    Sleep (100)
                    outputVoltage = AP.Anlr.FuncRdg("V")
                    
                    Cells(3 + i + j * 8, 3).value = outputVoltage
                    Cells(3 + i + j * 8, 5).value = THDN
                    Cells(3 + i + j * 8, 6).value = "x"
                    
ContinueLoop:
                Next i
            Next j
        Next
    Next
    
End Sub

Sub testthis()
    For j = 0 To &HC
        Debug.Print (j)
    Next j
End Sub

Sub OutputPowerVsBst()
    'Run at 1% THDN and 10% THDN, VBAT 3.7, 4.3, sweep from 6.5V boost to 10V boost, env tracker off
    
    Dim DUTFILE As String: DUTFILE = "C:\Users\maxwell.mckinnon\Dropbox\Maxim\ICs and Data\AX80\Bench\Rev A1\2583 Output Power vs Boost Voltage\403A_10V.98507t"
    Dim devaddr As Integer: devaddr = &H74
    epsilon = 2 ' Error if more than 2dB off in THDN
    MAX_RETRIES = 3
    Dim max As Double: max = 0
    Dim min As Double: min = -10 'Minimum for 1% too
    Dim THDNLEVELS(1) As Double: THDNLEVELS(0) = -40: THDNLEVELS(1) = -20 '1%, 10%
    Dim thdnlevel As Double
    Dim tolerance As Double: tolerance = 0.5 ' percent
    Dim VBATs(1) As Double: VBATs(0) = 3.7: VBATs(1) = 4.3
    VBAT_GPIB = "GPIB::01"
    
    Dim i As Integer
    Dim j As Integer
    Dim loadBoard1Addr As Integer: loadBoard1Addr = &HB0
    Dim loadBoard2Addr As Integer: loadBoard2Addr = &H4E
    
    

    
    For Each VBAT In VBATs
        Call Equipment_GPIB.Power_Supply_E3631A_.Supply_Set_Output(VBAT_GPIB, "P6V", VBAT, 5)
        
        ActiveWorkbook.Save
        For Each thdnlev In THDNLEVELS
            Sheets.Add
            ActiveSheet.Name = ActiveSheet.Name & " " & Str(thdnlev) & " " & Str(VBAT)
            thdnlevel = thdnlev
            Cells(2, 3).value = "outputVoltage"
            Cells(2, 4).value = "outputPower"
            Cells(2, 5).value = "THDN"
            Cells(2, 6).value = "x"
    
            For j = 0 To &H1C 'Boost loop 0x40 = from 00000 to 11100 '
                Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(devaddr, &H0, &H40, j)
                DoEvents
                
                Sleep (500)
                retryCount = 0
                
                ''''BEGIN THDN REGULATION''''
                Do
                    DoEvents
                    Call RegulateTHDN(min, max, -4, thdnlevel, tolerance)
                    'check THDN level
                    AP.Anlr.FuncMode = 4 ' Switch to THDN mode
                    Var = AP.Anlr.FuncRdg("dB")
                    retryCount = retryCount + 1
                    Call LoadEVKITFile_I2CBridge_16bit(DUTFILE, devaddr) ' reprogram part
                Loop While (Abs(Var - thdnlevel) > epsilon) And retryCount < MAX_RETRIES
                If retryCount >= MAX_RETRIES Then
                    Debug.Print ("Max Retry limit reached on THDN: " & Str(thdnlevel) & ", Iteration i: " & Str(i))
                    
                End If
                ''''END THDN REGULATION''''
                THDN = Var
                
                AP.Anlr.FuncMode = 0
                Sleep (100)
                outputVoltage = AP.Anlr.FuncRdg("V")
                
                Cells(3 + j, 3).value = outputVoltage
                Cells(3 + j, 5).value = THDN
                Cells(3 + j, 6).value = "x"
ContinueLoop:
            Next j
        Next
    Next
    
End Sub

Sub RegulateTHDN(min As Double, max As Double, inputP As Double, _
    target As Double, tolerance As Double)

    'Dim AP2700Reg As Regulate
    Set AP2700Reg = New Regulate
    
    AP2700Reg.min = (min) ' -15dBFS set to lower limit for regulation routine
    AP2700Reg.max = (max) ' 0dBFS set to upper limit for regulation routine
    AP2700Reg.target = (target) ' -40dBFS set to target
    AP2700Reg.tolerance = (tolerance) ' 1% set tolerance
    AP2700Reg.RegulationType = (1) '1 is +normal
    Call AP2700Reg.runAP2700internalRegulation
End Sub


