Attribute VB_Name = "I2C_Controls_"
'Attribute VB_Name = "I2C_Controls_"
Option Explicit
Option Base 0
Option Compare Text
' This module is used to interface between main programs and the cmodcomm mudule for I2C controls.  Each function will pass more useable
' arguments to the cmodcomm module.  This module requires to work with the cmodcomm mudule.
'
'       Below is a list of Functions that have been created for this module.  You can use this list and the search function to jump to a given
'       function.  Simply highlight or have your curser within the name of the function and hold [Ctrl] and [f] keys to bring up the search function.
'       you should see the name of the function you want to jump to is already in the search window.  Now just hit [Enter], and poof.  you are there.
'           List of Functions:
'                       I2C_Connect
'                       I2C_Disconnect
'                       I2C_8Bit_Write_Control
'                       I2C_8Bit_Read_Control
'                       I2C_16Bit_Write_Control
'                       I2C_16Bit_Read_Control
'    ****************** New Cammands for Bills I2C bridge ***********************************
'                       I2C_bridge_Connect
'                       I2C_device_Connect
'                       I2C_search_for_devices
'                       I2C_bridge_8Bit_Write_Control
'                       I2C_bridge_8Bit_Read_Control
'                       I2C_bridge_16Bit_Write_Control
'                       I2C_bridge_16Bit_Read_Control
'
'       Modification Log: (Date, By, Modification)
'                           04-26-2013, Chris Sibley,   Original Version
'
'
'
'

Private Function I2C_Connect() As Boolean
' Written By: Chris Sibley
' Last modifide: 9/10/2012 - Created
'
' Purpose: This block connects to the command module on the evaluation kit
'
'   Variable Delcaration
    Dim Result As Boolean                   ' Result will check if the program connects to the CMOD
    
    Result = True                           ' set the result to true
    
    If Not CmodBoardConnected() Then
        Call cmodcomm.CmodBoardConnect
        Call cmodcomm.CmodSupportsNativeSMBusCommands
    End If
    
    Result = CmodBoardConnected()           ' Checking if the board connects
        
    If Result = False Then                  ' if the programm doesn't connect it will trigger an error message
        MsgBox "ERROR:  Could not connect to the command module.", vbOKOnly, "Error 100.1"       '
    End If
    
    I2C_Connect = Result
End Function

Private Sub I2C_Disconnect()
' Written By: Chris Sibley
' Last modifide: 9/10/2012 - Created
'
' Purpose: This block disconnects from the command module on the evaluation kit
'
    CmodBoardDisconnect

    'MsgBox "The Command Module is now disconnected.", vbOKOnly, "CMOD OUT! Pease..."
End Sub

Private Function I2C_8Bit_Write_Control(ByVal Device_Address As Byte, ByVal Addr_Byte As Byte, ByVal Data_byte As Byte) As Boolean
' Written By: Chris Sibley
' Last modifide: 9/10/2012 - Created
'
' Purpose: This block connects to the command module on the evaluation kit
'
'   Variable Delcaration
'   Dim Device_Address As Byte      ' Device address for the given device
    Dim Results As Boolean
'   Device_Address = &H62           ' Set the device address for MAX98711 in hex (&H##)
    
    Results = CmodSMBusWriteByte(Device_Address, Addr_Byte, Data_byte)      ' Write to device and returns true if byte sent correctly
    
    I2C_8Bit_Write_Control = Results
    
'    If Results = False Then         ' If data didn't write correctly then send an error
'        MsgBox "Error:  Data Byte " & CStr(Data_Byte) & " to register " & CStr(Addr_Byte) & _
'                        " was not written to the device", vbOKOnly, "Error 100.2"
'    End If

End Function

Private Function I2C_8Bit_Read_Control(ByVal Device_Address As Byte, ByVal Addr_Byte As Byte, ByRef Data_byte As Byte) As Boolean
' Written By: Chris Sibley
' Last modifide: 9/10/2012 - Created
'
' Purpose: This block connects to the command module on the evaluation kit
'
'   Variable Delcaration
'    Dim Device_Address As Byte      ' Device address for the given device
    Dim Results As Boolean
          ' Device address for the given device
'
'    Device_Address = &H62           ' Set the device address for MAX98711 in hex (&H##)
    
    Results = CmodSMBusReadByte(Device_Address, Addr_Byte, Data_byte)      ' Write to device and returns true if byte sent correctly
    
    I2C_8Bit_Read_Control = Results
    
'    If Results = False Then         ' If data didn't write correctly then send an error
'        MsgBox "Error:  Was not able to read from register " & CStr(Addr_Byte) & _
'                        " .", vbOKOnly, "Error 100.3"
'    End If

    
End Function


Private Function I2C_16Bit_Write_Control(ByVal Device_Address As Byte, ByVal HAddr_Byte As Byte, ByVal LAddr_Byte As Byte, ByVal Data_byte As Byte) As Boolean
' Written By: Chris Sibley
' Last modifide: 9/10/2012 - Created
'
' Purpose: This block connects to the command module on the evaluation kit
'
'   Variable Delcaration
    Dim Data_byt(2) As Byte             ' array that contains the register address and data byte
    Dim Receive_Data_byt As Byte        ' Dummy variable
    Dim Results As Boolean
    
    Data_byt(0) = HAddr_Byte
    Data_byt(1) = LAddr_Byte
    Data_byt(2) = Data_byte
        
    Results = CmodI2CWriteAndReadBytes_native(Device_Address, 3, 0, Data_byt(0), Receive_Data_byt, 0)      ' Write to device and returns true if byte sent correctly
'    Results = CmodI2CWriteAndReadBytes(Device_Address, 3, 0, Data_byt(0), Receive_Data_byt, 0)      ' Write to device and returns true if byte sent correctly
    
    I2C_16Bit_Write_Control = Results
    
'    If Results = False Then         ' If data didn't write correctly then send an error
'        MsgBox "Error:  Data Byte " & CStr(Data_Byte) & " to register " & CStr(HAddr_Byte) & " " & _
'                        CStr(LAddr_Byte) & " was not written to the device", vbOKOnly, "Error 100.2"
'    End If

End Function

Private Function I2C_16Bit_Read_Control(ByVal Device_Address As Byte, ByVal HAddr_Byte As Byte, ByVal LAddr_Byte As Byte, ByRef Data_byte As Byte) As Boolean
' Written By: Chris Sibley
' Last modifide: 9/10/2012 - Created
'
' Purpose: This block connects to the command module on the evaluation kit
'
'   Variable Delcaration
    Dim Data_byt(2) As Byte      ' Device address for the given device
    Dim Results As Boolean
    Dim i As Integer                    ' Dummy variable
    
    Data_byt(0) = HAddr_Byte
    Data_byt(1) = LAddr_Byte
        
    Results = CmodI2CWriteAndReadBytes_native(Device_Address, 2, 1, Data_byt(0), Data_byte, 0)      ' Write to device and returns true if byte sent correctly
    
    I2C_16Bit_Read_Control = Results
    
'    If Results = False Then         ' If data didn't write correctly then send an error
'        MsgBox "Error:  Data Byte " & CStr(Data_Byte) & " to register " & CStr(HAddr_Byte) & " " & _
'                        CStr(LAddr_Byte) & " was not written to the device", vbOKOnly, "Error 100.2"
'    End If
'
End Function



Function I2C_bridge_Connect() As Boolean
' Written By: Chris Sibley
' Last modifide: 12/01/2015 - Created
'
' Purpose: This block checks to see if the I2C bridge is connected to the PC.
'
'   Variable Delcaration
    Dim Result As Boolean                   ' Result will check if the program connects to the CMOD
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Result = i2c.Connect()                           ' set the result to true
    
        
    If Result = False Then                  ' if the programm doesn't connect it will trigger an error message
        MsgBox "ERROR:  Could not connect to the I2C Bridge.", vbOKOnly, "Error 101.1"       '
    End If
    
    I2C_bridge_Connect = Result
End Function

Function I2C_device_Connect(ByVal Dev_ADDR_byte As Byte) As Boolean
' Written By: Chris Sibley
' Last modifide: 12/01/2015 - Created
'
' Purpose: This block checks to see if the I2C bridge is connected the device with the specified slave address.
'
'   Variable Delcaration
    Dim Result As Boolean                   ' Result will check if the program connects to the CMOD
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Result = i2c.I2CDevicePresent(Dev_ADDR_byte)                           ' set the result to true
    
    If Result = False Then                  ' if the programm doesn't connect it will trigger an error message
        MsgBox "ERROR:  Could not connect to the device with slave address: 0x" & Hex(Dev_ADDR_byte) & ".", vbOKOnly, "Error 102.1"       '
    End If
    
    I2C_device_Connect = Result
End Function

Function I2C_search_for_devices(ByRef Device_Count_int As Integer, ByRef Dev_ADDR_byte() As Byte) As Boolean
' Written By: Chris Sibley
' Last modifide: 12/01/2015 - Created
'
' Purpose: This block checks to see if the I2C bridge is connected to the PC.
'
'   Variable Delcaration
    Dim Result As Boolean                   ' Result will check if the program connects to the CMOD
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    
    Result = i2c.I2CDeviceSearch(Device_Count_int, Dev_ADDR_byte)                           ' set the result to true
    
    If Result = False Then                  ' if the programm doesn't connect it will trigger an error message
        MsgBox "ERROR:  Could not connect to the I2C Bridge.", vbOKOnly, "Error 101.1"       '
    ElseIf Device_Count_int = 0 Then
        MsgBox "ERROR: No devices found on the I2C bus.  Please make sure the device is powered up properly and the I2C bus is connected properly.", vbOKOnly, "Error 103.1"       '
        Result = False
    End If
    
    I2C_search_for_devices = Result
End Function


Function I2C_bridge_8Bit_Write_Control(ByVal Device_Address As Byte, ByVal Addr_Byte As Byte, ByVal Data_byte As Byte) As Boolean
' Written By: Chris Sibley
' Last modifide: 12/01/2015 - Created
'
' Purpose: This block writes a byte to a devcie with an 8-bit register address.
'
'   Variable Delcaration
    Dim Results As Boolean
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    Dim Write_Data_byte(0 To 1) As Byte
    Dim Read_Data_byte(0 To 1) As Byte
    
    Write_Data_byte(0) = Addr_Byte
    Write_Data_byte(1) = Data_byte
    
    
    Results = i2c.I2CWriteAndRead(Device_Address, Write_Data_byte, Read_Data_byte, 2, 0)      ' Write to device and returns true if byte sent correctly
    
    I2C_bridge_8Bit_Write_Control = Results
    
'    If Results = False Then         ' If data didn't write correctly then send an error
'        MsgBox "Error:  Data Byte " & CStr(Data_Byte) & " to register " & CStr(Addr_Byte) & _
'                        " was not written to the device", vbOKOnly, "Error 100.2"
'    End If

End Function

Function I2C_bridge_8Bit_Read_Control(ByVal Device_Address As Byte, ByVal Addr_Byte As Byte, ByRef Data_byte As Byte) As Boolean
' Written By: Chris Sibley
' Last modifide: 12/01/2015 - Created
'
' Purpose: This block read a byte to a devcie with an 8-bit register address.
'
'   Variable Delcaration
'    Dim Device_Address As Byte      ' Device address for the given device
    Dim Results As Boolean
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    Dim Write_Data_byte(0 To 1) As Byte
    Dim Read_Data_byte(0 To 1) As Byte
    
    Write_Data_byte(0) = Addr_Byte
    
    
    Results = i2c.I2CWriteAndRead(Device_Address, Write_Data_byte, Read_Data_byte, 1, 1)      ' Write to device and returns true if byte sent correctly
    Data_byte = Read_Data_byte(0)
    
    I2C_bridge_8Bit_Read_Control = Results
    
'    If Results = False Then         ' If data didn't write correctly then send an error
'        MsgBox "Error:  Was not able to read from register " & CStr(Addr_Byte) & _
'                        " .", vbOKOnly, "Error 100.3"
'    End If

    
End Function


Function I2C_bridge_16Bit_Write_Control(ByVal Device_Address As Byte, ByVal HAddr_Byte As Byte, ByVal LAddr_Byte As Byte, ByVal Data_byte As Byte) As Boolean
' Written By: Chris Sibley
' Last modifide: 12/01/2015 - Created
'
' Purpose: This block connects to the command module on the evaluation kit
'
'   Variable Delcaration
    Dim Results As Boolean
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    Dim Write_Data_byte(0 To 2) As Byte
    Dim Read_Data_byte(0 To 1) As Byte
    
    Write_Data_byte(0) = HAddr_Byte
    Write_Data_byte(1) = LAddr_Byte
    Write_Data_byte(2) = Data_byte
    
    
    Results = i2c.I2CWriteAndRead(Device_Address, Write_Data_byte, Read_Data_byte, 3, 0)      ' Write to device and returns true if byte sent correctly
    
    I2C_bridge_16Bit_Write_Control = Results
    
'    If Results = False Then         ' If data didn't write correctly then send an error
'        MsgBox "Error:  Data Byte " & CStr(Data_Byte) & " to register " & CStr(HAddr_Byte) & " " & _
'                        CStr(LAddr_Byte) & " was not written to the device", vbOKOnly, "Error 100.2"
'    End If

End Function

Function I2C_bridge_16Bit_Read_Control(ByVal Device_Address As Byte, ByVal HAddr_Byte As Byte, ByVal LAddr_Byte As Byte, ByRef Data_byte As Byte) As Boolean
' Written By: Chris Sibley
' Last modifide: 12/01/2015 - Created
'
' Purpose: This block connects to the command module on the evaluation kit
'
'   Variable Delcaration
    Dim Results As Boolean
    Dim i2c As I2CBridge.I2Ccontrol
    Set i2c = New I2CBridge.I2Ccontrol
    Dim Write_Data_byte(0 To 1) As Byte
    Dim Read_Data_byte(0 To 1) As Byte
    
    Write_Data_byte(0) = HAddr_Byte
    Write_Data_byte(1) = LAddr_Byte
    
    
    Results = i2c.I2CWriteAndRead(Device_Address, Write_Data_byte, Read_Data_byte, 2, 1)      ' Write to device and returns true if byte sent correctly
    Data_byte = Read_Data_byte(0)
    
    I2C_bridge_16Bit_Read_Control = Results
    
'    If Results = False Then         ' If data didn't write correctly then send an error
'        MsgBox "Error:  Data Byte " & CStr(Data_Byte) & " to register " & CStr(HAddr_Byte) & " " & _
'                        CStr(LAddr_Byte) & " was not written to the device", vbOKOnly, "Error 100.2"
'    End If
'
End Function
Private Sub Test_Case()

    Dim Device_Address As Byte
    Dim HReg_Address As Byte
    Dim LReg_Address As Byte
    Dim WData_ As Byte
    Dim RData_ As Byte
    Dim Check As Boolean
    
    Device_Address = &H64
    HReg_Address = &H0
    LReg_Address = &HFF
    WData_ = &H1
    
    
'    Check = I2C_Connect
    Check = I2C_16Bit_Write_Control(Device_Address, HReg_Address, LReg_Address, WData_)
'    Check = I2C_16Bit_Read_Control(Device_Address, HReg_Address, LReg_Address, RData_)
'    Call I2C_Disconnect

End Sub

Private Sub Test_I2C_bridge_Connect()

    Dim results_bool As Boolean
    
    results_bool = I2C_bridge_Connect()


End Sub
Private Sub Test_I2C_device_Connect()

    Dim results_bool As Boolean
    
    results_bool = I2C_device_Connect(&H62)


End Sub

Private Sub Test_I2C_search_for_devices()

    Dim results_bool As Boolean
    Dim Count_int As Integer
    Dim Address_List_byte(0 To 60) As Byte
    
    results_bool = I2C_search_for_devices(Count_int, Address_List_byte)


End Sub

Private Sub Test_I2C_bridge_16Bit_Read_Control()

    Dim results_bool As Boolean
    Dim Device_Address_byte As Byte
    Dim Hi_Address_byte As Byte
    Dim Lo_Address_byte As Byte
    Dim Data_byte As Byte
    
    Device_Address_byte = &H62
    Hi_Address_byte = &H7F
    Lo_Address_byte = &HFF
    
    results_bool = I2C_bridge_16Bit_Read_Control(Device_Address_byte, Hi_Address_byte, Lo_Address_byte, Data_byte)
    
End Sub

Private Sub Test_I2C_bridge_16Bit_Write_Control()

    Dim results_bool As Boolean
    Dim Device_Address_byte As Byte
    Dim Hi_Address_byte As Byte
    Dim Lo_Address_byte As Byte
    Dim Data_byte As Byte
    
    Device_Address_byte = &H62
    Hi_Address_byte = &H0
    Lo_Address_byte = &H80
    Data_byte = &H1
    
    results_bool = I2C_bridge_16Bit_Write_Control(Device_Address_byte, Hi_Address_byte, Lo_Address_byte, Data_byte)
    
End Sub



