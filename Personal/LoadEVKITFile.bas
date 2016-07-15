Attribute VB_Name = "LoadEVKITFile"
Sub LoadEVKITFile_I2CBridge_16bit(InputFile As String, devaddr As Integer)
    'Load an AX80 style evkit save file and write all the registers on the I2CBridge board
    
    
    Dim hf As Integer: hf = FreeFile
    Dim lines() As String, i As Long
    Dim leng As Integer
    
    Open InputFile For Input As #hf
        lines = Split(Input$(LOF(hf), #hf), vbNewLine)
    Close #hf
    
    ctrl_reg_found = False
    test_reg_found = False
    
    Call GlobalDisable(devaddr)
    
    'First check if there are test registers to write
    For i = 0 To UBound(lines)
        'Debug.Print "Line"; i; "="; lines(i)
        If lines(i) = "[Test Registers]" Then
            Debug.Print ("TEST REGISTERS FOUND")
            Debug.Print ("Chris Sucks!")
            test_reg_found = True
            i = i + 1
        End If
        If test_reg_found Then
            If Len(lines(i)) = 11 Then  ' check if it's a valid line (contains 11 entries is the test)
                haddr = Right(Left(lines(i), 4), 2)
                laddr = Right(Left(lines(i), 6), 2)
                bitVal = Right(lines(i), 2)
                hAddr_hex = CDec("&H" & haddr)
                lAddr_hex = CDec("&H" & laddr)
                bitVal_hex = CDec("&H" & bitVal)
                Call I2C_bridge_16Bit_Write_Control(devaddr, hAddr_hex, lAddr_hex, bitVal_hex)
                Debug.Print ("0x" & haddr & laddr & "  0x" & bitVal)
            Else
                Exit For
            End If
        End If
    Next
    
    'Now write user space registers, assume global enable is at end of list
    For i = 0 To UBound(lines)
        'Debug.Print "Line"; i; "="; lines(i)
        If lines(i) = "[Control Registers]" Then
            Debug.Print ("CONTROL REGISTERS FOUND")
            Debug.Print ("Just kidding he's all right")
            ctrl_reg_found = True
            i = i + 1
        End If
        If ctrl_reg_found Then
            If Len(lines(i)) = 11 Then  ' check if it's a valid line (contains 11 entries is the test)
                haddr = Right(Left(lines(i), 4), 2)
                laddr = Right(Left(lines(i), 6), 2)
                bitVal = Right(lines(i), 2)
                hAddr_hex = CDec("&H" & haddr)
                lAddr_hex = CDec("&H" & laddr)
                bitVal_hex = CDec("&H" & bitVal)
                Call I2C_bridge_16Bit_Write_Control(devaddr, hAddr_hex, lAddr_hex, bitVal_hex)
                Debug.Print ("0x" & haddr & laddr & "  0x" & bitVal)
            Else
                Exit For
            End If
        End If
    Next
    If test_reg_found Then Call EnterTestMode(devaddr) ' re-enter test mode if it's a test mode file
End Sub

Sub EnterTestMode(devaddr As Integer)
    Call I2C_bridge_16Bit_Write_Control(devaddr, &H1, &HFF, &H54)
    Call I2C_bridge_16Bit_Write_Control(devaddr, &H1, &HFF, &H4D)
End Sub

Sub GlobalDisable(devaddr As Integer)
    Call I2C_bridge_16Bit_Write_Control(devaddr, &H0, &HFF, &H0) ' Assumes 0x00FF = 0 is global disable
End Sub

Sub GlobalEnable(devaddr As Integer)
    Call I2C_bridge_16Bit_Write_Control(devaddr, &H0, &HFF, &H1) ' Assumes 0x00FF = 0 is global disable
End Sub

Sub test_example()
    Dim DUTFILE As String: DUTFILE = "C:\Users\maxwell.mckinnon\Dropbox\Maxim\ICs and Data\AX80\Bench\Rev A1\2584 Output Power vs Load\Output Power vs Load Board 403A.98507t"
    Call LoadEVKITFile_I2CBridge_16bit(DUTFILE, &H74)
End Sub
