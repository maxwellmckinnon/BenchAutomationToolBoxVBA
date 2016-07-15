Attribute VB_Name = "SquelchSweep"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub squelch()
    'Sweep AP2700 input while recording input and output
    'Repeat this sweep for all squelch settings
    'Squelch settings are at 0x50, 0x10 is enable, 0x0f are other bits
    'Load AP and DUT file before running this test
    
    Dim STARTROWdata As Integer: STARTROWdata = 3
    Dim STARTROWhexval As Integer: STARTROWhexval = 1
    Dim STARTCOLdata As Integer: STARTCOLdata = 2
    
    Dim devaddr As Integer: devaddr = &H74
    Dim SQUELCHADDR As Integer: SQUELCHADDR = &H50
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    I2CBridgeDemos.I2CBridgeConnected
    
    'Helpful AP functions
    'AP.DGen.ChAAmpl("dBFS") = -16.47826
    'Var = AP.Anlr.FuncRdg("dBV")
    'AP.Anlr.FuncFilter = 1 ' Aweighting
    'AP.Anlr.FuncFilter = 2 ' 20kHz Brick wall
    
    'Call I2C.I2CWriteByte16bit(DEVADDRI2C, RevIDReg, &H54)
    
    'run sweep once without squelch then for each setting of squelch on
    
    'UPSWEEP''''
    Call i2c.I2CWriteByte16bit(DEVADDRI2C, SQUELCHADDR, 0)
    Call OutputAPVoltageSweep(STARTROWdata, STARTCOLdata)
    
    STARTCOLdata = 4
    i = 15
    Cells(1, 4).value = &H10 + i ' Record Settings in the squelch register
    Call i2c.I2CWriteByte16bit(DEVADDRI2C, SQUELCHADDR, i + &H10)
    Call OutputAPVoltageSweep(STARTROWdata, STARTCOLdata)
    
    STARTCOLdata = 6
    For i = 0 To 14
        Cells(1, 6 + 2 * i).value = &H10 + i ' Record Settings in the squelch register
        Call i2c.I2CWriteByte16bit(DEVADDRI2C, SQUELCHADDR, i + &H10)
        Call OutputAPVoltageSweep(STARTROWdata, STARTCOLdata)
        
        STARTCOLdata = STARTCOLdata + 2
    Next i
    
    'Downsweep'''''
    Worksheets(ActiveSheet.index + 1).Select
    Call i2c.I2CWriteByte16bit(DEVADDRI2C, SQUELCHADDR, 0)
    Call OutputAPVoltageSweepDown(STARTROWdata, STARTCOLdata)
    
    STARTCOLdata = 4
    i = 15
    Cells(1, 4).value = &H10 + i ' Record Settings in the squelch register
    Call i2c.I2CWriteByte16bit(DEVADDRI2C, SQUELCHADDR, i + &H10)
    Call OutputAPVoltageSweepDown(STARTROWdata, STARTCOLdata)
    
    STARTCOLdata = 6
    For i = 0 To 14
        Cells(1, 6 + 2 * i).value = &H10 + i ' Record Settings in the squelch register
        Call i2c.I2CWriteByte16bit(DEVADDRI2C, SQUELCHADDR, i + &H10)
        Call OutputAPVoltageSweepDown(STARTROWdata, STARTCOLdata)
        
        STARTCOLdata = STARTCOLdata + 2
    Next i
      
End Sub

Sub OutputAPVoltageSweep(row As Integer, col As Integer)
    'Sweep from low to high voltagte, recording output voltage with both A weighted and non weighted filters
    'dump data into 2xsteps rowxcolumn where the second column is A weighted
    Dim VLOWSWEEP As Double: VLOWSWEEP = -100 ' dBFS
    Dim VHIGHSWEEP As Double: VHIGHSWEEP = -50 ' dBFS
    Dim inputV As Integer: inputV = VLOWSWEEP
    Dim VSTEP As Double: VSTEP = 1 ' dB step
    Dim MAXSTEPS As Integer: MAXSTEPS = (VHIGHSWEEP - VLOWSWEEP) / VSTEP + 1
    Dim count As Integer: count = 0
    
    AP.Anlr.FuncFilter = 2 ' 20kHz Brick wall
    Sleep (1000)
        
    Do While (count < MAXSTEPS)
        DoEvents
        
        AP.DGen.ChAAmpl("dBFS") = inputV
        Cells(row + count, 1).value = AP.DGen.ChAAmpl("dBFS")
        AP.Anlr.FuncFilter = 2 ' 20kHz Brick wall
        Sleep (1000)
        Cells(row + count, col).value = AP.Anlr.FuncRdg("dBV")
       
        AP.Anlr.FuncFilter = 1 ' A weighted
        Sleep (1000)
        Cells(row + count, col + 1).value = AP.Anlr.FuncRdg("dBV")
        
        count = count + 1
        inputV = inputV + VSTEP
    Loop
    
End Sub

Sub OutputAPVoltageSweepDown(row As Integer, col As Integer)
    'Sweep from low to high voltagte, recording output voltage with both A weighted and non weighted filters
    'dump data into 2xsteps rowxcolumn where the second column is A weighted
    Dim VLOWSWEEP As Double: VLOWSWEEP = -100 ' dBFS
    Dim VHIGHSWEEP As Double: VHIGHSWEEP = -50 ' dBFS
    Dim inputV As Integer: inputV = VHIGHSWEEP
    Dim VSTEP As Double: VSTEP = 1 ' dB step
    Dim MAXSTEPS As Integer: MAXSTEPS = (VHIGHSWEEP - VLOWSWEEP) / VSTEP + 1
    Dim count As Integer: count = 0
    
    AP.Anlr.FuncFilter = 2 ' 20kHz Brick wall
    Sleep (1000)
        
    Do While (count < MAXSTEPS)
        DoEvents
        
        AP.DGen.ChAAmpl("dBFS") = inputV
        Cells(row + count, 1).value = AP.DGen.ChAAmpl("dBFS")
        AP.Anlr.FuncFilter = 2 ' 20kHz Brick wall
        Sleep (1000)
        Cells(row + count, col).value = AP.Anlr.FuncRdg("dBV")
       
        AP.Anlr.FuncFilter = 1 ' A weighted
        Sleep (1000)
        Cells(row + count, col + 1).value = AP.Anlr.FuncRdg("dBV")
        
        count = count + 1
        inputV = inputV - VSTEP
    Loop
    
End Sub
    
