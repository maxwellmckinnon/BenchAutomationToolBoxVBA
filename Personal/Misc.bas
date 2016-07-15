Attribute VB_Name = "Misc"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Sub WriteSWEN()
    
    'Call AX80BestBoost.bestBoostWrite
    DEV = &H74
    
    Do While (1)
        DoEvents
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HFF, 0)
        Sleep (1000)
        AP.BarGraph.Reset 1
        Sleep (1000)
        Call I2C_Controls_.I2C_bridge_16Bit_Write_Control(DEV, 0, &HFF, 1)
        Sleep (1000)
        AP.BarGraph.Reset 1
        Sleep (1000)
    Loop
    
End Sub

Function SheetExists(SheetName As String, Optional wb As Excel.Workbook)
   Dim s As Excel.Worksheet
   If wb Is Nothing Then Set wb = ActiveWorkbook
   On Error Resume Next
   Set s = wb.Sheets(SheetName)
   On Error GoTo 0
   SheetExists = Not s Is Nothing
End Function


