Attribute VB_Name = "OutputPowervsLoad_ManualLoad"
Sub OutputPowerVsLoad_ManualLoad()
    'Run at 1% THDN and 10% THDN, VBAT 3.7V PVDD 8V, VBAT 4.3V PVDD 10V, sweep load from 3ohm to 100ohm + 33uH
    'Program will prompt user to switch loads
    
    epsilon = 2 ' Error if more than 2dB off in THDN
    MAX_RETRIES = 3
    Dim max As Double: max = 0
    Dim min As Double: min = -20 'Minimum for 1% too
    Dim THDNLEVELS(1) As Double: THDNLEVELS(0) = -40: THDNLEVELS(1) = -20 '1%, 10%
    Dim thdnlevel As Double
    Dim tolerance As Double: tolerance = 0.5 ' percent
    Dim VBATs(1) As Double: VBATs(0) = 3.7: VBATs(1) = 4.3
    Dim PVDDs(1) As Double: PVDDs(0) = 8#: PVDDs(1) = 10#
    Dim PVDD As Double
    VBAT_GPIB = "GPIB::01"
    Dim devaddr As Integer: devaddr = &H74
    Dim Load_GPIB As String: Load_GPIB = "GPIB::11"
    
    Dim i As Integer
    Dim j As Integer
    Dim loadBoard1Addr As Integer: loadBoard1Addr = &HB0
    Dim loadBoard2Addr As Integer: loadBoard2Addr = &H4E
    
    For L = 0 To 12 'Regulate through 13 loads from 3 to 100ohms
        DoEvents
        GlobalDisable (devaddr)
        MsgBox ("Please connect Load #" & Str(L) & ", click OK when done")
        For lq = 0 To 100
            LoadValue = MeasureLoad(Load_GPIB)
            Res = MsgBox(Str(LoadValue) & " Ohms were measured - Click Yes if correct, No to Redo measurement", vbYesNo)
            If Res = vbYes Then Exit For
        Next lq
        Sleep (500)
        GlobalEnable (devaddr)
        
        For i = 0 To 1 ' The two test conditions, 3.7V/8V and 4.3V/10V
            DoEvents
            VBAT = VBATs(i)
            PVDD = PVDDs(i)
            Call Equipment_GPIB.Power_Supply_E3631A_.Supply_Set_Output(VBAT_GPIB, "P6V", VBAT, 5)
            Call LoadEVKITFile_I2CBridge_16bit("C:\Users\maxwell.mckinnon\Dropbox\Maxim\ICs and Data\AX80\Bench\Rev A1\2584 Output Power vs Load\Output Power vs Load Board 465A_8V.98507t", devaddr) ' reprogram part
            Call PVDD_setVoltage(PVDD)
            
            ActiveWorkbook.Save
            For t = 0 To 1 'Each thdnlev In THDNLEVELS
                DoEvents
                Sheets.Add
                thdnlevel = THDNLEVELS(t)
                ActiveSheet.Name = ActiveSheet.Name & " " & Str(thdnlevel) & " " & Str(VBAT) & " " & Str(PVDD)
                
                ''''BEGIN THDN REGULATION''''
                retryCount = 0
                Do
                    DoEvents
                    Call RegulateTHDN(min, max, -4, thdnlevel, tolerance)
                    'check THDN level
                    AP.Anlr.FuncMode = 4 ' Switch to THDN mode
                    Var = AP.Anlr.FuncRdg("dB")
                    retryCount = retryCount + 1
                    Call LoadEVKITFile_I2CBridge_16bit("C:\Users\maxwell.mckinnon\Dropbox\Maxim\ICs and Data\AX80\Bench\Rev A1\2584 Output Power vs Load\Output Power vs Load Board 465A_8V.98507t", devaddr) ' reprogram part
                    Call PVDD_setVoltage(PVDD)
                Loop While (Abs(Var - thdnlevel) > epsilon) And retryCount < MAX_RETRIES
                If retryCount >= MAX_RETRIES Then
                    Debug.Print ("Max Retry limit reached on THDN: " & Str(thdnlevel) & ", Iteration i: " & Str(i))
                    MsgBox ("THDN unable to resolve to desired level: " & Str(thdnlevel) & ". Please manually adjust then click OK.")
                End If
                ''''END THDN REGULATION''''
                
                Var = AP.Anlr.FuncRdg("dB")
                THDN = Var
                
                AP.Anlr.FuncMode = 0
                Sleep (100)
                outputVoltage = AP.Anlr.FuncRdg("V")
                
                colOff = 20
                'Header text
                Cells(36, colOff + 1 + 4 * t + 8 * i).value = "Output Voltage"
                Cells(36, colOff + 2 + 4 * t + 8 * i).value = "THDN"
                Cells(36, colOff + 3 + 4 * t + 8 * i).value = "x"
                Cells(36, colOff + 4 + 4 * t + 8 * i).value = "LoadValue"
                
                'Data
                Cells(37 + L, colOff + 1 + 4 * t + 8 * i).value = outputVoltage
                Cells(37 + L, colOff + 2 + 4 * t + 8 * i).value = THDN
                Cells(37 + L, colOff + 3 + 4 * t + 8 * i).value = "x"
                Cells(37 + L, colOff + 4 + 4 * t + 8 * i).value = LoadValue
                
            Next t
        Next i
    Next L
End Sub

Sub PVDD_setVoltage(v As Double, Optional devaddr As Integer = &H74)
    'PVDD voltages 8 and 10 only - crude function only for AX80
    vlt = CInt(v)
    If vlt = 8 Then
        Call I2C_bridge_16Bit_Write_Control(devaddr, &H0, &H40, &HC) '8V write
    End If
    If vlt = 10 Then
        Call I2C_bridge_16Bit_Write_Control(devaddr, &H0, &H40, &H1C) ' 10V write
    End If
End Sub
