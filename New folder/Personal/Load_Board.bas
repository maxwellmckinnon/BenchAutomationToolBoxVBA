Attribute VB_Name = "Load_Board"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Function Loadboard(address As Integer, load As Integer, Sense_On As Integer)
    
    'load: 0=all off, 1=Load 1 on only, 2=Load 2 on only, 3=load 3 on only
    ' Sense_On; 1=on 0=off
    
    Dim Set_Load As Byte: Set_Load = &H0
    
    Select Case load ' Picks value to drive load board
     Case 0
     Set_Load = &H0
     
     Case 1
     Set_Load = &H1
     
     Case 2
     Set_Load = &H2
     
     Case 3
     Set_Load = &H4
     
     Case 12
     Set_Load = &H3
     
     Case 13
     Set_Load = &H5
     
     Case 23
     Set_Load = &H6
     
     Case 123
     Set_Load = &H7
     
     End Select
     
    If Sense_On = 0 Then
        Set_Load = Set_Load + &HF0
    End If
    
    Call I2C_bridge_8Bit_Write_Control(address, &H6, &H0)  ' Set Port 1 as output
    Call I2C_bridge_8Bit_Write_Control(address, &H7, &H0)  ' Set Port 2 as output
    Call I2C_bridge_8Bit_Write_Control(address, &H2, &HF0)   ' Turn off all loads first
    Call I2C_bridge_8Bit_Write_Control(address, &H2, Set_Load) 'Sets the programmed load number

End Function


Function LoadboardHex(address As Integer, load As Integer)

    'load: 0=all off, 1=Load 1 on only, 2=Load 2 on only, 3=load 3 on only
    ' Sense_On; 1=on 0=off
    
    Dim Set_Load As Byte: Set_Load = load
    
    Call I2C_bridge_8Bit_Write_Control(address, &H6, &H0)  ' Set Port 1 as output
    Call I2C_bridge_8Bit_Write_Control(address, &H7, &H0)  ' Set Port 2 as output
    Call I2C_bridge_8Bit_Write_Control(address, &H2, &HF0)   ' Turn off all loads first
    Sleep (200)
    Call I2C_bridge_8Bit_Write_Control(address, &H2, Set_Load) 'Sets the programmed load number

End Function

Private Sub load_test()

    'Call LoadboardHex(&HB0, 2)
    'Call LoadboardHex(&H4E, 0)
    address = &H4E
    'address = &HB0
    Call I2C_bridge_8Bit_Write_Control(address, &H6, &H0)  ' Set Port 1 as output
    Call I2C_bridge_8Bit_Write_Control(address, &H7, &H0)  ' Set Port 2 as output
    Call I2C_bridge_8Bit_Write_Control(address, &H2, &HF0)   ' Turn off all loads first
    Sleep (200)
    Call I2C_bridge_8Bit_Write_Control(address, &H2, &H47)   'Sets the programmed load number
    
    '
    
    'Dim i As Integer: i = 0
    '
    '
    '
    'While i < 32
    'Call I2C_bridge_8Bit_Write_Control(&HB0, &H2, i)
    'MsgBox ("Check leds")
    'i = i + 1
    'Wend

End Sub
