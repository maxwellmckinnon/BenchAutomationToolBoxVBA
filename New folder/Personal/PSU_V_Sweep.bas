Attribute VB_Name = "PSU_V_Sweep"
Sub ModifyVBAT_BDELevel()
    'Modify VBAT to change the BDE level, run THD stereo sweep on Class D outputs
    'AX84 file : "Clipper Setup.98726"
    
    Dim V_LEVELS As Variant: V_LEVELS = Array(4.1, 3.7, 3.3, 2.9, 2.7, 2.5)
    Dim PSU_GPIB As String: PSU_GPIB = "GPIB::03"
    
    Dim v As Variant
    AP.Sweep.Append = False
    For Each v In V_LEVELS
        Call Power_Supply_E3631A_.Supply_Set_Output(PSU_GPIB, "P6V", v)
        AP.Sweep.Start
        AP.Sweep.Append = True
    Next v
End Sub
