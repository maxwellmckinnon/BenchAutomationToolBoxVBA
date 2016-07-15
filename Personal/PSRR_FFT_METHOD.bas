Attribute VB_Name = "PSRR_FFT_METHOD"
Sub PSRR_by_FFTs()
    'The AP's crosstalk has a wideband filter. Use very high resolution FFT to get better results.
    'put frequencies to sweep starting at A2
    
    Dim Frequencies(1000) As Double
    Dim F_array_length As Integer
    
    'Dim FREQUENCIES_ARRAY As Variant: FREQUENCIES_ARRAY = Array(20, 30, 50, 80, 100, 217, 300, 500, 800, 1000, 2000, 3000, 5000, 8000, 10000, 15000, 20000)
    
    Call getFrequencies(Frequencies, F_array_length)
    Call Run_FFTs(Frequencies, F_array_length)
    
End Sub

Sub getFrequencies(ByRef frequency_array() As Double, ByRef F_array_length As Integer)
    ' Read frequencies from Excel starting at A2
    Dim n As Integer: n = 0
    For r = 2 To 1000
        If IsEmpty(Cells(r, 1)) Then
            Exit For ' quit loop
        End If
        n = n + 1
        frequency_array(r - 2) = Cells(r, 1)
        Cells(r, 2).value = 5
    Next r
    
    F_array_length = n
    
End Sub

Sub Run_FFTs(frequency_array() As Double, F_array_length As Integer)
    'Pass the frequencies in Hz to sweep
    
    Dim f As Variant
    AP.Sweep.Append = False
    For n = 0 To F_array_length - 1
        DoEvents
        AP.Gen.Freq("Hz") = frequency_array(n)
        AP.Sweep.Start
        AP.Sweep.Append = True
        
        'Dim sweptFreq_array As Variant: sweptFreq_array = AP.Data.XferToArray(n, 0, "Hz")
        'Dim Source_array As Variant: Source_array = AP.Data.XferToArray(n, 1, "dBV")
        'Dim Target_array As Variant: Target_array = AP.Data.XferToArray(n, 2, "dBV")
        'Dim xTalkdB As Double: xTalkdB = CalcCrossTalk(Source_array, Target_array, sweptFreq_array, frequency_array(n))
        
    Next n
    
End Sub

Function CalcCrossTalk(SourceTalk As Variant, TargetTalk As Variant, Frequencies As Variant, frequency As Double) As Double
    'Return the crosstalk value in dB - scane through sourcetalk and target talk arrays for the frequency in question. Check a few bins around frequency and grab maximum value
    'Not finished
    Dim match As Boolean: match = False
    Dim oldDelta As Double: oldDelta = 30000
    Dim newDelta As Double: newDelta = 30000
    Dim index As Integer: index = 0
    Dim n As Integer: n = 0
    
    'Find the index of the correct frequency
    For c = 0 To Frequencies.Length
        newDelta = Math.Abs(Frequencies(n) - frequency)
        If newDelta < oldDelta Then
            oldDelta = newDelta
            index = n
            
        End If
        
        n = n + 1
    Next c
    
    'Find the max value in the SourceTalk near the frequency bin - search a few nearby bins
    If index = 0 Then
    
    End If
    
    
    
    
End Function


