Attribute VB_Name = "ToneGen"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub ToneGenFFTTest_AX80()
    ' Run through the modes of the tone generator bitfield. Capture FFT of each on ap2700
    ' Test 0000 through 1010
    DEV = &H74
    STARTCODE = &H0
    ENDCODE = &HA
    TONEADDRESS = &H38
    
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    I2CBridgeDemos.I2CBridgeConnected ' for debug
    
    c = 0
    For i = STARTCODE To ENDCODE
        DoEvents
        Call i2c.I2CWriteByte16bit(DEV, TONEADDRESS, i) ' Enable WDT: WDT_ENA = 1, 100ms
        AP.PSIA.Tx.FrameClk.Rate("Hz") = 48000# ' Weird bug where AP2700 would keep switching to 12kHz LRCLK - would affect tone gen frequencies - part is in slave mode - maybe buffer board related??
        Sleep (2000)
        If i = STARTCODE Then
            AP.Sweep.Append = False
        End If
        AP.Sweep.Start
        AP.Sweep.Append = True
        AP.Graph.Legend.comment(c + 1, 1) = "TONE_CONFIG = 0x" & Hex(i)
        c = c + 1
    Next i
End Sub

Sub ToneGenTHDNPerfTest_AX80()
    ' Run through the modes of the tone generator bitfield. Capture THDN of each on ap2700
    ' Test 0000 through 1010
    DEV = &H74
    STARTCODE = &H0
    ENDCODE = &HA
    TONEADDRESS = &H38
    
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    I2CBridgeDemos.I2CBridgeConnected ' for debug
    
    c = 0
    For i = STARTCODE To ENDCODE
        DoEvents
        Call i2c.I2CWriteByte16bit(DEV, TONEADDRESS, i) ' Enable WDT: WDT_ENA = 1, 100ms
        AP.PSIA.Tx.FrameClk.Rate("Hz") = 48000# ' Weird bug where AP2700 would keep switching to 12kHz LRCLK - would affect tone gen frequencies - part is in slave mode - maybe buffer board related??
        Sleep (4000)
        Cells(c + 1, 1).value = i
        Cells(c + 1, 2).value = AP.Anlr.ChAFreqRdg("Hz")
        Cells(c + 1, 3).value = AP.Anlr.FuncRdg("dB")
        
        c = c + 1
    Next i

End Sub
