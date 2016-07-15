Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets("403A Efficiency vs frequen 1k").Select
    Sheets("403A Efficiency vs frequen 1k").Copy Before:=Sheets(1)
    Sheets("403A Efficiency vs frequen  (2").Select
    Sheets("403A Efficiency vs frequen  (2").Name = "403A Efficiency vs fq 2"
    Range("K31").Select
    Sheets("403A Efficiency vs frequen 1k").Select
    Sheets("403A Efficiency vs frequen 1k").Name = "403A Efficiency vs fq 1k"
    Sheets("403A Efficiency vs fq 2").Select
End Sub
