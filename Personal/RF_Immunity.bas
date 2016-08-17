Attribute VB_Name = "RF_Immunity"
Sub LabView_APDump_DataCleanUp()
    'This processes the data obtained from Hani's Labview RFI code and readies it to be plotted
    
    'INPUT
        'Assumes the FFT data is listed in columns A (HZ), B (dB ch 1), C (dB ch 2)
    
    'OUTPUT
        'puts data to plot in columns
        '   G (Carrier Frequency),
        '   H (Modulated frequency - near 1kHz by default)
        '   I (Ch1 output)
        '   J (Ch2 output)
        '   K (Modulated frequency found for left output - just curious if there are any differences to the right output, could be debug information)
    
    'Assumes the swept frequencies were from 100MHz to 1100MHz in steps of 50MHz, then from 1100MHz to 3000MHz in steps of 100MHz.
    
    'Step through column A for the row that contains the closest number to the modulated frequency (1kHz default)
    'Take the adjacent +/- x (4 by default) rows and grab the largest datapoint from them
    'Put datapoints in column H, I, and J
    
    STARTROW_INPUT = 5
    DATAPOINTS = 40 + 1 ' (1100 - 100) / 50 + 1 + (3000 - 1100) / 100 + 1 = 40 + 1     ||| The plus one is because the automation also includes no modulation base point
    STARTROW_OUTPUT = 2
    FREQ_COL = 1
    ch1_COL = 2
    ch2_COL = 3
    FFTLength = 511 ' number of rows per FFT data
    ADJACENT_SEARCH_WIDTH = 2 ' +/- search width around modulated freq bin
    
    MODULATED_HZ = 1000
    MOD_FREQ_COL = 8
    Ch1_OUT_COL = 9
    Ch2_OUT_COL = 10
    MOD_FREQ_COL_Ch2 = 11
    
    n = STARTROW_INPUT
    
    For datapoint = 1 To DATAPOINTS
        newdiff = 99999
        bestdiff = 999999
        Do While (newdiff < bestdiff) ' Find closest to modulated frequency value and row where it lives - assumes values are ordered, hill climber optimizer
            n = n + 1
            bestdiff = newdiff
            newdiff = Abs(Cells(n + 1, FREQ_COL).value - MODULATED_HZ)
        Loop
        
        'Find highest peak in FFT around adjacent datapoints to modulation frequency
        bestCh1 = -999
        bestCh1_row = 1
        bestCh2 = -999
        bestCh2_row = 1
        For m = (n - ADJACENT_SEARCH_WIDTH) To (n + ADJACENT_SEARCH_WIDTH)
            If (Cells(m, ch1_COL).value) > bestCh1 Then
                bestCh1 = Cells(m, ch1_COL).value
                bestCh1_row = m
            End If
            If (Cells(m, ch2_COL).value) > bestCh2 Then
                bestCh2 = Cells(m, ch2_COL).value
                bestCh2_row = m
            End If
        Next m
        
        'Report frequency and amplitude in output
        Cells(STARTROW_OUTPUT + datapoint - 1, MOD_FREQ_COL).value = Cells(bestCh1_row, FREQ_COL).value
        Cells(STARTROW_OUTPUT + datapoint - 1, Ch1_OUT_COL).value = Cells(bestCh1_row, ch1_COL).value
        Cells(STARTROW_OUTPUT + datapoint - 1, Ch2_OUT_COL).value = Cells(bestCh2_row, ch2_COL).value
        Cells(STARTROW_OUTPUT + datapoint - 1, MOD_FREQ_COL_Ch2).value = Cells(bestCh2_row, FREQ_COL).value
        
        n = n + FFTLength
    Next datapoint
    
End Sub
