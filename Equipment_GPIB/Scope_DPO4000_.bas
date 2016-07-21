Attribute VB_Name = "Scope_DPO4000_"
Option Explicit
Option Base 0
Option Compare Text
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function URLDownloadToFileA Lib "urlmon" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As LongPtr, ByVal lpfnCB As LongPtr) As LongPtr

'This Module will Control the Agilent E3631A only.  This module requires that the Agilent IO Libraries Suite 16.3 be installed.
'   You can download and install the VISA from following web site:
'       http://www.home.agilent.com/en/pd-1985909-pn-E2094/io-libraries-suite-162?nid=-33330.977662.00&cc=US&lc=eng&cmpid=zzfindiosuite
'   After the Agilent IO Libraries Suite 16.3 is installed, it will install VISA that allows this module to function.  Go the "Tools" toolbar
'   and click on "References...".  Here you want to make sure "VISA COM 3.0 Type Library" is selected.
'
'       Below is a list of Functions that have been created for this module.  You can use this list and the search function to jump to a given
'       function.  Simply highlight or have your curser within the name of the function and hold [Ctrl] and [f] keys to bring up the search function.
'       you should see the name of the function you want to jump to is already in the search window.  Now just hit [Enter], and poof.  you are there.
'           List of Functions:
'                       Scope_OnOffSingle
'                       Scope_Set_Channel
'                       Scope_Get_Channel
'                       Scope_Set_Measurement
'                       Scope_Get_Measurement
'                       Scope_Set_Horizontal
'                       Scope_Get_Horizontal
'                       Scope_Set_Delay
'                       Scope_Get_Delay
'                       Scope_Get_State
'                       Scope_Set_Trigger
'                       Scope_Get_Trigger
'                       Scope_Set_Acquire
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'                           12-5-2013, Chris Sibley,    Added the Supply_Get_Output_Enable function.
'
'


'********************************************************************************************************************************************************
' Sub Routine Scope_OnOffSingle
'********************************************************************************************************************************************************
'   This function will control state of the Scope run, stop, or single sequence.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, String, Output_Name = This sets the range of the measurnemt.  The valid srtrings can be:
'                                                   "On" = Run
'                                                   "Off"  = Stop
'                                                   "Single" = Single
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_OnOffSingle(ByVal IP_Address As String, ByVal State As String, Optional ByVal Auto As Boolean = True)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    If State = "On" Then
    
        instrument.WriteString "ACQuire:STOPAfter RUNSTop"
        instrument.WriteString "ACQuire:STATE 1"
        
    ElseIf State = "Single" Then
    
        instrument.WriteString "ACQuire:STOPAfter SEQuence"
        instrument.WriteString "ACQuire:STATE 1"

    ElseIf State = "Off" Then
    
        instrument.WriteString "ACQuire:STATE 0"
    
    End If

End Sub

'********************************************************************************************************************************************************
' Sub Routine Scope_Set_Channel
'********************************************************************************************************************************************************
'   This sub routine Configure a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, Intiger, Channel = This calls out the channel you wish to control.  Channels range from 1 to 4.
'                   Optional, String, State =  This controls if the channel is on or off.
'                                                   "Off"
'                                                   "On"
'                   Optional, Double, Scale_ =  This controls Scale for the given channel. If greater than 50 nothing will changed. Typical Scale settings are:
'                                                   0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1, 2, 5, 10, 20, 50
'                   Optional, Double, Position =  This controls Position for the given channel. If greater than 9 nothing will changed. Typical Scale settings are:
'                                                   -5 to 5
'
'
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_Set_Channel(ByVal IP_Address As String, _
                        ByVal Channel As Integer, _
                        Optional ByVal State As String = "Off", _
                        Optional ByVal Scale_ As Double = 51, _
                        Optional ByVal Position As Double = 10, _
                        Optional ByVal Couple As String = "DC")
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    If Channel > 0 And Channel < 5 Then
        If State = "On" Then
        
            instrument.WriteString "SELect:CH" & CStr(Channel) & " ON"
            
        ElseIf State = "off" Then
        
            instrument.WriteString "SELect:CH" & CStr(Channel) & " OFF"
        End If
        
        If Scale_ > 0 And Scale_ < 50 Then
        
            instrument.WriteString "CH" & CStr(Channel) & ":SCAle " & CStr(Scale_)
            
        End If
        
        If Position > -5.1 And Position < 5.1 Then
        
            instrument.WriteString "CH" & CStr(Channel) & ":POSition " & CStr(Position)
            
        End If
        
        If Couple = "DC" Or Couple = "DC" Then
            
            instrument.WriteString "CH" & CStr(Channel) & ":COUPling " & Couple
            
        End If
        
    End If

End Sub

'********************************************************************************************************************************************************
' Sub Routine Scope_Get_Channel
'********************************************************************************************************************************************************
'   This sub routine Retrieves the setting of a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, Intiger, Channel = This calls out the channel you wish to control.  Channels range from 1 to 4.
'                   Required, String, State =  Returns the following.
'                                                   "Off"
'                                                   "On"
'                   Required, Double, Scale_ =  This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'                   Required, Double, Position =  This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_Get_Channel(ByVal IP_Address As String, _
                        ByVal Channel As Integer, _
                        ByRef State As String, _
                        ByRef Scale_ As Double, _
                        ByRef Position As Double)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Check As Integer
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    If Channel > 0 And Channel < 5 Then
    
        instrument.WriteString "SELect:CH" & CStr(Channel) & "?"
        Check = CInt(instrument.ReadString())
         
        If Check = 1 Then
            State = "On"
        Else
            State = "Off"
        End If
        
        instrument.WriteString "CH" & CStr(Channel) & ":SCAle?"
        Scale_ = CDbl(instrument.ReadString())
         
        instrument.WriteString "CH" & CStr(Channel) & ":POSition?"
        Position = CDbl(instrument.ReadString())
    
    End If

End Sub

'********************************************************************************************************************************************************
' Sub Routine Scope_Set_Horizontal
'********************************************************************************************************************************************************
'   This sub routine Retrieves the setting of a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, Double, Scale_ =  This controls Scale for the horizontal
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_Set_Horizontal(ByVal IP_Address As String, ByVal Scale_ As Double)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Check As Integer
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    If Scale_ > 0 And Scale_ < 10 Then
    
        instrument.WriteString "HORIZONTAL:MAIN:SCALE " & CStr(Scale_)
        
    End If

End Sub

'********************************************************************************************************************************************************
' Sub Routine Scope_Get_Horizontal
'********************************************************************************************************************************************************
'   This sub routine Retrieves the setting of a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, Double, Scale_ =  This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_Get_Horizontal(ByVal IP_Address As String, ByRef Scale_ As Double)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Check As Integer
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    instrument.WriteString "HORIZONTAL:MAIN:SCALE?"
    Scale_ = CDbl(instrument.ReadString())

End Sub

'********************************************************************************************************************************************************
' Sub Routine Scope_Set_Delay
'********************************************************************************************************************************************************
'   This sub routine Retrieves the setting of a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, String, Delay_on_Off =  This controls to enable Delay on the horizontal trigger.  This only except the following"
'                                           "On"
'                                           "Off"
'                   Optional, Double, Delay_Sec =  This controls the delay on the horizontal trigger in seconds when the delay is enable.
'                                           Default is set to 0.0sec
'                   Optional, Double, Position_Percentage =  This controls the trigger position when the delay is disabled.  The position is percenage.
'                                           The range is between 0% (far left) to 100% (far Right)
'                                           Default is set to 50% (Middle)
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-27-2014, Chris Sibley,   Original Version
'
Sub Scope_Set_Delay(ByVal IP_Address As String, ByVal Delay_on_Off As String, Optional ByVal Delay_Sec As Double = 0, Optional ByVal Position_Percentage As Double = 50)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Check As Integer
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    instrument.WriteString "HORIZONTAL:DELay:MODe?"
    Check = CInt(instrument.ReadString())
    
    If Delay_on_Off = "On" Then
        instrument.WriteString "HORIZONTAL:DELay:MODe On"
        instrument.WriteString "HORIZONTAL:DELay:TIMe " & CStr(Delay_Sec)
    ElseIf Delay_on_Off = "Off" Then
        instrument.WriteString "HORIZONTAL:DELay:MODe Off"
        If Position_Percentage >= 0 Or Position_Percentage <= 100 Then
            instrument.WriteString "HORIZONTAL:POSition " & CStr(Position_Percentage)
        End If
    End If
    
End Sub


'********************************************************************************************************************************************************
' Sub Routine Scope_Get_Delay
'********************************************************************************************************************************************************
'   This sub routine Retrieves the setting of a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, String, Delay_on_Off =  This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'                   Required, Double, Delay_Sec =  This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'                   Required, Double, Position_Percentage =  This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-27-2014, Chris Sibley,   Original Version
'
Sub Scope_Get_Delay(ByVal IP_Address As String, ByRef Delay_on_Off As String, ByRef Delay_Sec As Double, ByRef Position_Percentage As Double)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Check As Integer
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    instrument.WriteString "HORIZONTAL:DELay:MODe?"
    Check = CInt(instrument.ReadString())
    
    If Check = 1 Then
        Delay_on_Off = "On"
    ElseIf Check = 0 Then
        Delay_on_Off = "Off"
    End If
    
    instrument.WriteString "HORIZONTAL:DELay:TIMe?"
    Delay_Sec = CDbl(instrument.ReadString())
    
    instrument.WriteString "HORIZONTAL:POSition?"
    Position_Percentage = CDbl(instrument.ReadString())
    
End Sub

'********************************************************************************************************************************************************
' Sub Routine Scope_Set_Measurement
'********************************************************************************************************************************************************
'   This sub routine Configure a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, Intiger, Meas_Number = This calls out the measuemnt number in which you want to configure.
'                                                   There are only 8 slots availbale.
'                   Required, Intiger, Channel = This calls out the channel you wish to control.  Channels range from 1 to 4.
'                   Optional, String, Type_ =  This will select the type of measurement. the following types are availbale through
'                                              this sub routine:
'                                                   "FREQuency"
'                                                   "PERIod"
'                                                   "PK2Pk"
'                                                   "AMPlitude"
'                                                   "MAXimum"
'                                                   "MINImum"
'                   Optional, String, State =  This controls if the measurement is on or off.
'                                                   "Off"
'                                                   "On"
'
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_Set_Measurement(ByVal IP_Address As String, _
                        ByVal Meas_Number, _
                        ByVal Channel As Integer, _
                        Optional ByVal Type_ As String = "NC", _
                        Optional ByVal State As String = "Off")
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    If Meas_Number > 0 And Meas_Number < 9 Then
        
        If Type_ = "FREQuency" Or _
           Type_ = "PERIod" Or _
           Type_ = "PK2Pk" Or _
           Type_ = "AMPlitude" Or _
           Type_ = "MAXimum" Or _
           Type_ = "MINImum" Then
        
            instrument.WriteString "MEASUrement:MEAS" & CStr(Meas_Number) & ":TYPE " & Type_
            
        End If
        
        If Channel > 0 And Channel < 5 Then
            
            instrument.WriteString "MEASUrement:MEAS" & CStr(Meas_Number) & ":SOUrce1 CH" & CStr(Channel)
            
        End If
        
        If State = "On" Then
        
            instrument.WriteString "MEASUrement:MEAS" & CStr(Meas_Number) & ":STATE ON"
            
        ElseIf State = "Off" Then
        
            instrument.WriteString "MEASUrement:MEAS" & CStr(Meas_Number) & ":STATE OFF"
            
        End If
        
    End If

End Sub

'********************************************************************************************************************************************************
' Sub Routine Scope_Get_Measurement
'********************************************************************************************************************************************************
'   This sub routine Configure a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, Intiger, Meas_Number = This calls out the measurement number in which you want to configure.
'                                                   There are only 8 slots availbale.
'                   Required, Double, Measure = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'                   Optional, String, Type =  This controls what measurement is returned.
'                                                   "VALue" (Default)
'                                                   "MINImum"
'                                                   "MAXimum"
'                                                   "MEAN"
'                                                   "STDdev"
'                                                   "Average"
'                   Optional, Integer, Num_of_Average =  This controls the number of averages.  the default is 10
'
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_Get_Measurement(ByVal IP_Address As String, ByVal Meas_Number, _
                            ByRef Measure As Double, Optional ByVal Type_ As String = "Value", Optional ByVal Num_of_Average As Integer = 10)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Dim RunningTotal As Double
    Dim i As Double
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")
    
    If Meas_Number > 0 And Meas_Number < 9 Then
    
        
        Select Case Type_
            Case "MIMImum"
                instrument.WriteString "MEASUrement:MEAS" & CStr(Meas_Number) & ":MINImum?"
            Case "MAXimum"
                instrument.WriteString "MEASUrement:MEAS" & CStr(Meas_Number) & ":MAXimum?"
            Case "MEAN"
                instrument.WriteString "MEASUrement:MEAS" & CStr(Meas_Number) & ":MEAN?"
            Case "STDdev"
                instrument.WriteString "MEASUrement:MEAS" & CStr(Meas_Number) & ":STDdev?"
            Case "Average"
                RunningTotal = 0
                For i = 1 To Num_of_Average
                    instrument.WriteString "MEASUrement:MEAS" & CStr(Meas_Number) & ":VALue?"
                    Call Sleep(200)
                    RunningTotal = RunningTotal + CDbl(instrument.ReadString())
                Next i
            Case Else
                instrument.WriteString "MEASUrement:MEAS" & CStr(Meas_Number) & ":VALue?"
        End Select
         
        If Type_ = "Average" Then
            Measure = RunningTotal / Num_of_Average
        Else
            Measure = CDbl(instrument.ReadString())
        End If

    End If

End Sub




'********************************************************************************************************************************************************
' Sub Routine Scope_Get_State
'********************************************************************************************************************************************************
'   This sub routine Configure a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, Double, State = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_Get_State(ByVal IP_Address As String, ByRef State As String)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Check As Integer
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    instrument.WriteString "ACQuire:STATE?"
    Check = CInt(instrument.ReadString())

    If Check = 1 Then
        State = "on"
    Else
        State = "off"
    End If

End Sub


'********************************************************************************************************************************************************
' Sub Routine Scope_Set_Trigger
'********************************************************************************************************************************************************
'   This sub routine Configure a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, Intiger, Channel = This calls out the channel you wish to control.  Channels range from 1 to 4.
'                   Optional, String, Slope =  This will select which slope the trigger will capture on:
'                                                   "RISe"
'                                                   "FALL"
'                   Optional, Double, Level_ =  This controls the level of the trigger
'
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_Set_Trigger(ByVal IP_Address As String, _
                        ByVal Channel As Integer, _
                        Optional ByVal Slope As String = "RISe", _
                        Optional ByVal Level_ As Double = 0)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    If Channel > 0 And Channel < 5 Then
    
        instrument.WriteString "TRIGger:A:EDGE:SOUrce CH" & CStr(Channel)
        
        If Slope = "RISe" Or Slope = "FALL" Then
            instrument.WriteString "TRIGger:A:EDGE:SLOPe " & Slope
        End If
        
        instrument.WriteString "TRIGger:A:LEVEl " & CStr(Level_)
    
    End If
    
End Sub

'********************************************************************************************************************************************************
' Sub Routine Scope_Get_Trigger
'********************************************************************************************************************************************************
'   This sub routine Configure a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, Intiger, Meas_Number = This calls out the measuemnt number in which you want to configure.
'                                                   There are only 8 slots availbale.
'                   Required, Double, Measure = This is sent by reference to return the string that is read from the equipment.
'                                                   This argument has to be a variable.
'                   Optional, String, Type =  This controls what measurement is returned.
'                                                   "VALue" (Default)
'                                                   "MINImum"
'                                                   "MAXimum"
'                                                   "MEAN"
'                                                   "STDdev"
'
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_Get_Trigger(ByVal IP_Address As String, _
                        ByRef Channel As Integer, _
                        ByRef Slope As String, _
                        ByRef Level_ As Double)
   
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Dim Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    instrument.WriteString "TRIGger:A:EDGE:SOUrce?"
    Check = instrument.ReadString()
    Channel = CInt(Right(Left(Check, 3), 1))
    
    instrument.WriteString "TRIGger:A:EDGE:SLOPe?"
    Slope = instrument.ReadString()
    
    instrument.WriteString "TRIGger:A:LEVEl?"
    Level_ = CDbl(instrument.ReadString())
    
End Sub

'********************************************************************************************************************************************************
' Sub Routine Scope_Set_Acquire
'********************************************************************************************************************************************************
'   This sub routine Configure a given channel.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, String, Type_ =  This controls what measurement is returned.
'                                                   "SAMple"
'                                                   "PEAKdetect"
'                                                   "HIRes"
'                                                   "AVErage"
'                                                   "ENVelope"
'                   Optional, Integer, Average_Number = this sets the number of averages
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-22-2014, Chris Sibley,   Original Version
'
Sub Scope_Set_Acquire(ByVal IP_Address As String, ByRef Type_ As String, Optional ByVal Average_Number As Integer = 128)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Dim Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    If Type_ = "SAMple" Or _
        Type_ = "PEAKdetect" Or _
        Type_ = "HIRes" Or _
        Type_ = "AVErage" Or _
        Type_ = "ENVelope" Then
    
        instrument.WriteString "ACQuire:MODe " & Type_
        
        If Type_ = "AVErage" Then
            instrument.WriteString "ACQuire:NUMAVg " & Average_Number
        End If
        
    End If
    
End Sub


'********************************************************************************************************************************************************
' Sub Routine Scope_Save_Image
'********************************************************************************************************************************************************
'   This function will save an image of the oscilloscope screen (capture a scopeshot) to a local USB drive.  The user speficies the folder name for the file path.
'   Image is saved as a .PNG file
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                                   The address is from the scope it should be something like "10.33.89.208"
'                   Required, String, Folder = This sets the folder name in the file path for the image saved on the local USB drive.
'                                                For example:  E:\Folder\Imagename
'                   Required, String, Imagename = This sets the name of the image file to be saved on the local USB drive.
'                                                For example:  E:\Folder\Imagename
'
'
'
'       Modification Log: (Date, By, Modification)
'                           06-19-2014, Evan Ragsdale,   Original Version
'
Sub Scope_Save_Image(ByVal IP_Address As String, ByVal Folder As String, ByVal Imagename As String)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    instrument.WriteString "SAVe:IMAGe " & Chr(34) & Folder & "\" & Imagename & ".png" & Chr(34)

End Sub



'********************************************************************************************************************************************************
' Sub Routine Scope_Save_Image_to_File
'********************************************************************************************************************************************************
'   This function will save an image of the oscilloscope screen (capture a scopeshot) to a local USB drive.  The user speficies the folder name for the file path.
'   Image is saved as a .PNG file
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                                   The address is from the scope it should be something like "10.33.89.208"
'                   Required, String, Folder = This sets the folder name in the file path for the image saved on the local USB drive.
'                                                For example:  E:\Folder\Imagename
'                   Required, String, Imagename = This sets the name of the image file to be saved on the local USB drive.
'                                                For example:  E:\Folder\Imagename
'
'
'
'       Modification Log: (Date, By, Modification)
'                           06-19-2014, Evan Ragsdale,   Original Version
'
Sub Scope_Save_Image_to_File(ByVal IP_Address As String, ByVal Folder As String, ByVal Imagename As String, Optional ByVal Scope_Port_int As Integer = 81)
    
    Dim lngRetVal As Long
    Dim url As String
    Dim LocalFilename As String
    
    url = "http://" & IP_Address & ":" & Scope_Port_int & "/image.png" ' Buil URL string
    LocalFilename = Folder & "\" & Imagename & ".png"
    
    
    lngRetVal = URLDownloadToFileA(0, url, LocalFilename, 0, 0)
End Sub


'********************************************************************************************************************************************************
' Sub Routine Scope_Make_Directory
'********************************************************************************************************************************************************
'   This function will make a directory on the USB drive inserted in the port of the oscilloscope
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                                   The address is from the scope it should be something like "10.33.89.208"
'                   Required, String, Directory = This sets the directory to create
'                                                For example:  "E:\LoadTransient"
'
'       Modification Log: (Date, By, Modification)
'                           06-19-2014, Evan Ragsdale,   Original Version
'
Sub Scope_Make_Directory(ByVal IP_Address As String, ByVal Directory As String)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    instrument.WriteString "FILESystem:MKDir " & Chr(34) & "E:\" & Directory & Chr(34)

End Sub

'********************************************************************************************************************************************************
' Sub Routine Scope_Set_MessageBox
'********************************************************************************************************************************************************
'   This sub routine loads the message box that can be displayed on the oscilloscope.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Required, String, Quote_ =  This will be the string that gets displayed on the screen of the oscilloscope.
'
'                   Optional, Boolean, On_Off = this will display the message box on the oscilloscope when true, else not (default)
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-04-2016, Chris Sibley,   Original Version
'
Sub Scope_Set_MessageBox(ByVal IP_Address As String, ByRef Quote_ As String, Optional ByVal On_Off As Boolean = False)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Dim Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    
    instrument.WriteString "MESSage:SHOW """ & Quote_ & """"
       
    If On_Off = True Then
        instrument.WriteString "MESSage:STATE ON"
    Else
        instrument.WriteString "MESSage:STATE OFF"
    End If
    
End Sub


'********************************************************************************************************************************************************
' Sub Routine Scope_Clear_MessageBox
'********************************************************************************************************************************************************
'   This sub routine Configure clears the Message box that can be displayed on the screen.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, IP_Address = the Ip address for the given instrument that you want to control
'                                           The address is from the scope it should be something like "10.33.89.208"
'                   Optional, Boolean, On_Off = this will display the message box on the oscilloscope when true, else not (default)
'
'
'
'       Modification Log: (Date, By, Modification)
'                           05-04-2016, Chris Sibley,   Original Version
'
Sub Scope_Clear_MessageBox(ByVal IP_Address As String, Optional ByVal On_Off As Boolean = False)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Dim Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    
    instrument.WriteString "MESSage:CLEAR"
       
    If On_Off = True Then
        instrument.WriteString "MESSage:STATE ON"
    Else
        instrument.WriteString "MESSage:STATE OFF"
    End If
    
End Sub


'********************************************************************************************
'Test Program
Private Sub Tetest()
'    Dim ioMgr As VisaComLib.ResourceManager
'    Dim instrument As VisaComLib.FormattedIO488
'    Dim Error_Check As String
    Dim Str_State As String
    Dim Delay_ As Double
    Dim State As String
    Dim Position_ As Double
    Dim Str_Message As String
    Dim CH_ As Integer
    Dim Slope As String
    Dim Level_ As Double
    
    'Call Scope_Save_Image("10.33.89.115", "AAA", "Scopeshot")
    Call Scope_Make_Directory("10.33.89.115", "AAA")

End Sub




'********************************************************************************************************************************************************
' Sub Routine Scope_No_Return_Template
'********************************************************************************************************************************************************
'   This function will set the voltage to the given output.
'
'       Arguments:  (Required/Optional, Data Type, Name = description)
'                   Required, String, GPIB_Address = the GPIB address for the given instrument that you want to control
'                                           The address is ranged between "GPIB::00" to "GPIB::31"
'                   Required, String, Output_Name = This sets the range of the measurnemt.  The valid srtrings can be:
'                                                   "On" = positive 6V output
'                                                   "Off"  = positive 25V output
'                                                   "Single" = negative 25V output
'
'
'
'       Modification Log: (Date, By, Modification)
'                           04-25-2013, Chris Sibley,   Original Version
'
Sub Scope_No_Return_Template(ByVal IP_Address As String, ByVal State As String, Optional ByVal Auto As Boolean = True)
    
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    Dim Error_Check As String
    Set ioMgr = New VisaComLib.ResourceManager

    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("TCPIP0::" & IP_Address & "::inst0::INSTR")

    If State = "On" Then
    
        instrument.WriteString "ACQuire:STOPAfter RUNSTop"
        instrument.WriteString "ACQuire:STATE 1"
        
    ElseIf State = "Single" Then
    
        instrument.WriteString "ACQuire:STOPAfter SEQuence"
        instrument.WriteString "ACQuire:STATE 1"

    ElseIf State = "Off" Then
    
        instrument.WriteString "ACQuire:STATE 0"
    
    End If

End Sub

Private Sub test_string()
        Dim a As String
        Dim Quote_ As String
        Quote_ = "PVDD = 10V"
        a = Chr(34) & "10.33.89.115"
        
        MsgBox a
        a = "MESSage:SHOW """ & Quote_ & """"
        
        MsgBox a
End Sub
Private Sub test_Scope_Set_MessageBox()


    Call Scope_Set_MessageBox("10.33.89.193", "PVDD = 10V", True)
    
End Sub


Private Sub test_Scope_Clear_MessageBox()


    Call Scope_Clear_MessageBox("10.33.89.193", True)
    
End Sub


Private Sub Test_Scope_Save_Image_to_File()
    Dim scope_add As String
    Dim Pix_filename As String
    Dim pix_path As String
    
    scope_add = "10.33.89.193"
    Pix_filename = "pix2"
    pix_path = ActiveWorkbook.Path
    
    Call Scope_Save_Image_to_File(scope_add, pix_path, Pix_filename)
End Sub





'Private Sub test()
'Dim scope_add As String: scope_add = "10.33.89.89"
'Dim scope_port As Integer: scope_port = 81
'Dim Pix_filename As String: Pix_filename = "c:\temp\pix1.png"
'
'Call DownloadFile(scope_add, scope_port, Pix_filename)
'
'End Sub
'
'Public Function DownloadFile(scope_add As String, scope_port As Integer, Pix_filename As String) As Boolean
'Dim lngRetVal As Long
'Dim url As String
'Dim LocalFilename As String
'
'url = "http://" & scope_add & ":" & scope_port & "/image.png" ' Buil URL string
'LocalFilename = Pix_filename
'
'
'lngRetVal = URLDownloadToFileA(0, url, LocalFilename, 0, 0)
''If lngRetVal = 0 Then
''    If Dir(LocalFilename) <> vbNullString Then
''        DownloadFile = True
''    End If
''End If
'End Function


