Attribute VB_Name = "WatchDogTester"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub WriteWatchDogTimer_Check_MODE()
    ' Repeatedly write 0xE9 to 0x13 to reset the SW watchdog timer
    ' The difference is that this time, the mode is not set to SW, set to HW
    ' Start with an already running device with WDT_ENA = 0 from the config file before running this test
    DEV = &H64
    MAX_ITERATIONS = 5#
    
    I2C_Controls_.I2C_Connect
    Sleep (1000)
    Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H12, &HF) ' Enable WDT: WDT_ENA = 1, HW 50ms
    


    'Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H12, &H1) ' Enable WDT: WDT_ENA = 1, 100ms
    'Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H12, &H5) ' Enable WDT: WDT_ENA = 1, 500ms
    'Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H12, &H9) ' Enable WDT: WDT_ENA = 1, 1000ms
    'Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H12, &HD) ' Enable WDT: WDT_ENA = 1, 2000ms
    
    n = 0
    Do While (n < MAX_ITERATIONS)
        Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H13, &HE9) ' reset SW WDT
        n = n + 1
    Loop
    
    I2C_Controls_.I2C_Disconnect
End Sub

Sub WriteWatchDogTimer_Check_MODE_AX80()
    ' Repeatedly write 0xE9 to 0x13 to reset the SW watchdog timer
    ' The difference is that this time, the mode is not set to SW, set to HW
    ' Start with an already running device with WDT_ENA = 0 from the config file before running this test
    DEV = &H74
    MAX_ITERATIONS = 5000#
    
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    I2CBridgeDemos.I2CBridgeConnected ' for debug
    
    Sleep (1000)
    'Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H12, &HF) ' Enable WDT: WDT_ENA = 1, HW 50ms
    Call i2c.I2CWriteByte16bit(DEV, &H0 * 256 + &H12, &H1) ' Enable WDT: WDT_ENA = 1, 100ms
    Call i2c.I2CWriteByte16bit(DEV, &H0 * 256 + &H12, &H13) ' Enable WDT: WDT_ENA = 1, WDT_MODE = 1 (HW)
    'Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H12, &H5) ' Enable WDT: WDT_ENA = 1, 500ms
    'Call I2C_Controls_.I2C_16Bit_Write_Control(DEV, &H0, &H12, &H9) ' Enable WDT: WDT_ENA = 1, 1000ms
    'Call I2C.I2CWriteByte16bit(DEV, &H0 * 256 + &H12, &HD) ' Enable WDT: WDT_ENA = 1, 2000ms
    
    Call i2c.I2CWriteByte16bit(DEV, &H0 * 256 + &HE, &HFF)  'reset IRQ
    Call i2c.I2CWriteByte16bit(DEV, &H0 * 256 + &HFF, 1) ' Enable the part
    n = 0
    Do While (n < MAX_ITERATIONS)
        DoEvents
        Call i2c.I2CWriteByte16bit(DEV, &H0 * 256 + &H13, &HE9) ' reset SW WDT
        n = n + 1
    Loop
    
End Sub

