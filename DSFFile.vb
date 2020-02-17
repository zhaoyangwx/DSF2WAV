Public Class DSFFile
    Public Property FileSize As UInt32
    Public Property ID3Tag As UInt32
    Public Property ChunkSize As UInt32
    Public Property ChanelCount As UInt32
    Public Property Sampling As UInt32
    Public Property BitDepth As UInt32
    Public Property SampleSize As UInt32 'Bits per channel
    Public ReadOnly Property SampleBytes As UInt32
        Get
            Return SampleSize / 8
        End Get
    End Property
    Public Property BlockSize As UInt32
    Public Property DataBytes As UInt32 '>= SampleBytes*ChannelCount
    Public Property Data As Byte(,)
    Public Property DataHead As Integer


    Private FreqResR() As Double
    Private FIRCoefficient() As Double
    Private _Taps As Integer = 1024
    Private _PCMSampleFrequency As Long = 44100
    Private _OverSampling As Integer = 64
    Private _fFreqLow As Long = 0
    Private _fFreqHigh As Long = 22050
    Private _K_Alpha As Double = 8
    Private _StopdB As Double = -200
    Public _FIRCoeffUpdateSuspended As Boolean = False
    Public Property FIRCoeffUpdateSuspended As Boolean
        Set(value As Boolean)
            _FIRCoeffUpdateSuspended = value
            If _FIRCoeffUpdateSuspended = False Then
                UpdateFIRCoeff()
            End If
        End Set
        Get
            Return _FIRCoeffUpdateSuspended
        End Get
    End Property
    Public Sub SuspendFIRCoeffUpdate()
        FIRCoeffUpdateSuspended = True
    End Sub
    Public Sub ResumeFIRCoeffUpdate()
        FIRCoeffUpdateSuspended = False
    End Sub
    Public Property Taps As Integer 'For FIR Proc
        Set(value As Integer)
            _Taps = value
            If FIRCoeffUpdateSuspended Then Exit Property
            UpdateFIRCoeff()
        End Set
        Get
            Return _Taps
        End Get
    End Property
    Public Property PCMSampleFrequency As Long
        Set(value As Long)
            _PCMSampleFrequency = value
            If FIRCoeffUpdateSuspended Then Exit Property
            UpdateFIRCoeff()
        End Set
        Get
            Return _PCMSampleFrequency
        End Get
    End Property
    Public Property OverSampling As Integer
        Set(value As Integer)
            _OverSampling = value
            If FIRCoeffUpdateSuspended Then Exit Property
            UpdateFIRCoeff()
        End Set
        Get
            Return _OverSampling
        End Get
    End Property
    Public Property fFreqLow As Long
        Set(value As Long)
            _fFreqLow = value
            If FIRCoeffUpdateSuspended Then Exit Property
            UpdateFIRCoeff()
        End Set
        Get
            Return _fFreqLow
        End Get
    End Property
    Public Property fFreqHigh As Long
        Set(value As Long)
            _fFreqHigh = value
            If FIRCoeffUpdateSuspended Then Exit Property
            UpdateFIRCoeff()
        End Set
        Get
            Return _fFreqHigh
        End Get
    End Property
    Public Property K_Alpha As Double
        Set(value As Double)
            _K_Alpha = value
            UpdateFIRCoeff()
        End Set
        Get
            Return _K_Alpha
        End Get
    End Property
    Public Property StopdB As Double
        Set(value As Double)
            _StopdB = value
            UpdateFIRCoeff()
        End Set
        Get
            Return _StopdB
        End Get
    End Property
    Public Sub UpdateFIRCoeff()
        CreateFres(PCMSampleFrequency * OverSampling, fFreqLow, fFreqHigh, FreqResR)
        CalcCoeffs(FreqResR, FIRCoefficient)
        DoKaiser(FIRCoefficient, K_Alpha)
    End Sub

    'Generate Frequency Responce. 
    Private Sub CreateFres(ByVal SampleFrequency As Long, ByVal fFreqLow As Long, ByVal fFreqHigh As Long, ByRef FreqResR() As Double)
        'gain = pow(10.0, desired_dB / 20.0);
        'There are folding at Fs/2. _____L~~~~~H_____|_____H'~~~L'____. |=Fs/2


        ReDim FreqResR(Taps - 1)
        Dim i As Integer
        Dim taps_half As Integer = Taps / 2
        Dim stopgain As Double = 10 ^ (StopdB / 20)
        For i = 0 To taps_half - 1
            Dim CurrentFreq As Double = SampleFrequency * (i + 1) / Taps
            If CurrentFreq < fFreqLow Then
                FreqResR(i) = stopgain
            ElseIf CurrentFreq < fFreqHigh Then
                FreqResR(i) = 1
            Else 'If CurrentFreq < SampleFrequency / 2 Then
                FreqResR(i) = stopgain
            End If
        Next
        For i = taps_half To Taps - 1
            FreqResR(i) = FreqResR(Taps - i - 1)
        Next
    End Sub

    'make FIR coeff, from Frequency Response.
    Private Sub CalcCoeffs(ByRef FreqResR() As Double, ByRef FIRCoefficient() As Double)
        ReDim FIRCoefficient(Taps - 1)
        Dim costable(Taps - 1) As Double
        Dim Tempcoeff_r(Taps - 1) As Double
        For i As Integer = 0 To Taps - 1
            costable(i) = Math.Cos(2 * Math.PI * i / Taps)
        Next
        For i As Integer = 0 To Taps - 1
            Tempcoeff_r(i) = 0
            Dim Coeff As Double = 0
            For j As Integer = 0 To Taps - 1
                Coeff += FreqResR(j) * costable((i * j) Mod Taps)
            Next
            Tempcoeff_r(i) = Coeff / Taps
        Next
        For i As Integer = 0 To Taps / 2 - 1
            FIRCoefficient(i) = Tempcoeff_r(Taps / 2 - i)
        Next
        For i As Integer = Taps / 2 To Taps - 1
            FIRCoefficient(i) = Tempcoeff_r(i - (Taps / 2))
        Next
    End Sub
    Private Sub DoKaiser(ByRef FIRCoefficient() As Double, ByVal K_Alpha As Double)
        Dim i As Integer
        Dim Numer As Double, Denom As Double
        Dim center As Double, Kg As Double, Kd As Double
        Dim Kwindow(Taps) As Double

        Denom = Bessel(K_Alpha)
        center = (Taps - 1) / 2
        For i = 0 To Taps - 1
            Kg = (i * 1.0 - center) / center
            Kd = (K_Alpha * Math.Sqrt(1.0 - Kg * Kg))
            Numer = Bessel(Kd)
            Kwindow(i) = Numer / Denom
        Next
        For i = 0 To Taps - 1
            FIRCoefficient(i) *= Kwindow(i)
        Next
    End Sub
    Private Function Bessel(ByVal alpha As Double) As Double
        Dim delta As Double = 0.000000000001
        Dim BesselValue As Double
        Dim Term, k, F As Double
        k = 0
        BesselValue = 0
        F = 1
        Term = 0.5
        While (Term < -delta Or delta < Term)
            k += 1
            F *= (alpha / 2) / k
            Term = F * F
            BesselValue += Term
        End While
        Return BesselValue
    End Function


    Public Sub New()
        UpdateFIRCoeff()
    End Sub
    Public Shared Function FromFile(ByVal FileName As String) As DSFFile
        If Not My.Computer.FileSystem.FileExists(FileName) Then Return Nothing
        Dim fb() As Byte = My.Computer.FileSystem.ReadAllBytes(FileName)
        If BitConverter.ToString(fb, 0, 4) <> "44-53-44-20" Then
            Return Nothing
        End If
        Dim DSFFile1 As New DSFFile
        Try
            With DSFFile1
                .FileSize = BitConverter.ToUInt32(fb, &HC)
                .ID3Tag = BitConverter.ToUInt32(fb, &H14)
                .ChunkSize = BitConverter.ToUInt32(fb, &H20)
                .ChanelCount = BitConverter.ToUInt32(fb, &H34)
                .Sampling = BitConverter.ToUInt32(fb, &H38)
                .BitDepth = BitConverter.ToUInt32(fb, &H3C)
                .SampleSize = BitConverter.ToUInt32(fb, &H40)
                .BlockSize = BitConverter.ToUInt32(fb, &H48)
                .DataBytes = BitConverter.ToUInt32(fb, &H54)
                Dim b(.DataBytes + .BlockSize * .ChanelCount - 1) As Byte
                .DataHead = &H5C
                .Taps = 1024
                .UpdateFIRCoeff()
                Array.Copy(fb, .DataHead, b, 0, .DataBytes)
                ReDim .Data(.ChanelCount - 1, .SampleBytes - 1)
                Parallel.For(0, .SampleBytes * .ChanelCount,
                        Sub(i As Integer)
                            Dim DataOfs As Integer = (i \ (.BlockSize * .ChanelCount)) * .BlockSize + (i Mod .BlockSize)
                            If DataOfs >= .SampleBytes Then Exit Sub
                            .Data((i \ .BlockSize) Mod .ChanelCount, DataOfs) = b(i)
                        End Sub)
            End With
        Catch ex As Exception
            Return Nothing
        End Try
        Return DSFFile1
    End Function
    Public Sub ToTextFile(ByVal FileName As String, Optional ByVal TargetSampling As UInt32 = 44100)
        If Locked Then Exit Sub
        Locked = True
        Dim th As New Threading.Thread(
            Sub()

                Dim DwnSampleRatio As Integer = Sampling / TargetSampling / 8
                'fFreqHigh = TargetSampling / 2
                Dim OutputData(,) As Double = ToFloatData(TargetSampling)

                Dim s As String = ""
                For i As Integer = 0 To SampleBytes / DwnSampleRatio - 1
                    For chanel As Integer = 0 To ChanelCount - 1
                        s &= OutputData(chanel, i) & vbCrLf
                        If s.Length > 10000 Then
                            My.Computer.FileSystem.WriteAllText(FileName, s, True)
                            s = ""
                        End If
                    Next
                Next
                My.Computer.FileSystem.WriteAllText(FileName, s, True)
                s = ""
                Locked = False
            End Sub)
        th.Start()
    End Sub
    Public Sub ToWaveFile(ByVal FileName As String, Optional ByVal TargetSampling As UInt32 = 44100, Optional ByVal OutputBitdepth As Integer = 24)
        If Locked Then Exit Sub
        Locked = True
        If OutputBitdepth <> 16 And OutputBitdepth <> 24 Then OutputBitdepth = 24
        Dim th As New Threading.Thread(
            Sub()
                PCMSampleFrequency = TargetSampling
                'fFreqHigh = TargetSampling / 2
                Dim OutputData(,) As Double = ToFloatData(TargetSampling)
                Dim fs As IO.FileStream = IO.File.Open(FileName, IO.FileMode.Create)
                fs.Write(System.Text.Encoding.ASCII.GetBytes("RIFF"), 0, 4)
                Dim DataSize As UInteger = OutputData.Length * (OutputBitdepth / 8)
                Dim FileSize As UInteger = DataSize + 8 + 28
                fs.Write(BitConverter.GetBytes(FileSize), 0, 4)
                fs.Write(System.Text.Encoding.ASCII.GetBytes("WAVEfmt "), 0, 8)
                fs.Write(BitConverter.GetBytes(CType(16, UInteger)), 0, 4)
                fs.Write(BitConverter.GetBytes(CType(1, Short)), 0, 2)
                fs.Write(BitConverter.GetBytes(CType(ChanelCount, Short)), 0, 2)
                fs.Write(BitConverter.GetBytes(CType(TargetSampling, UInteger)), 0, 4)
                fs.Write(BitConverter.GetBytes(CType(TargetSampling * ChanelCount * OutputBitdepth / 8, UInteger)), 0, 4)
                fs.Write(BitConverter.GetBytes(CType(ChanelCount * OutputBitdepth / 8, Short)), 0, 2)
                fs.Write(BitConverter.GetBytes(CType(OutputBitdepth, Short)), 0, 2)
                fs.Write(System.Text.Encoding.ASCII.GetBytes("data"), 0, 4)
                fs.Write(BitConverter.GetBytes(DataSize), 0, 4)
                Dim ProgMax As Integer = OutputData.GetLength(1)
                Dim Prog As Integer = 0
                For j As Integer = 0 To OutputData.GetLength(1) - 1
                    For iChannel As Integer = 0 To OutputData.GetLength(0) - 1
                        Dim PCMValue As UInteger = (OutputData(iChannel, j) + 1) / 2 * 256 * 256 * 256
                        PCMValue = (PCMValue And &HF0FF) Or ((((PCMValue >> 16) And &HFF) Xor &H80) << 16)
                        fs.Write(BitConverter.GetBytes(PCMValue), 3 - OutputBitdepth / 8, OutputBitdepth / 8)
                    Next
                    Prog += 1
                    ProgVal = Prog / ProgMax * 10000
                Next
                fs.Close()
                Locked = False
            End Sub)
        th.Start()
    End Sub
    Public Function ToFloatData(Optional ByVal TargetSampling As UInt32 = 44100) As Double(,)
        Dim DwnSampleRatio As Integer = Sampling / TargetSampling / 8
        Dim OutputData(ChanelCount - 1, SampleBytes / DwnSampleRatio - 1) As Double
        Dim Prog As Integer = 0
        Dim TotalProg As Integer = SampleBytes / DwnSampleRatio
        Parallel.For(Taps / 8 - 1, SampleBytes / DwnSampleRatio,
                Sub(ti As Integer)
                    Dim i As Integer = ti * DwnSampleRatio
                    For Chanel As Integer = 0 To ChanelCount - 1
                        Dim PCMTmp As Double = 0
                        For ofs As Integer = 0 To Taps / 8 - 1
                            Dim bt As Byte = 0
                            If i - ofs >= 0 Then bt = Data(Chanel, i - ofs)
                            For j As Integer = 0 To 7
                                PCMTmp += ((bt Mod 2) - 0.5) * 2 * FIRCoefficient((Taps / 8 - 1 - ofs) * 8 + j)
                                bt >>= 1
                            Next
                        Next
                        OutputData(Chanel, i / DwnSampleRatio) = Math.Max(-1, Math.Min(PCMTmp, 1))
                    Next
                    Threading.Interlocked.Add(Prog, 1)
                    ProgVal = Prog / TotalProg * 10000
                End Sub)
        Return OutputData
    End Function

    Private ProgVal As Integer = 0
    Private _Locked As Boolean = False
    Private Property Locked As Boolean
        Set(value As Boolean)
            If _Locked = False And value = True Then
                _Locked = value
                ProgMonitor = New Threading.Thread(
                    Sub()
                        While Locked
                            RaiseEvent ProgressReport(ProgVal)
                            Threading.Thread.Sleep(ProgReportIntv)
                        End While
                    End Sub)
                ProgMonitor.Start()
                RaiseEvent ProgressStarted()
            ElseIf _Locked = True And value = False Then
                RaiseEvent ProgressFinished()
            End If
            _Locked = value
        End Set
        Get
            Return _Locked
        End Get
    End Property
    Public Property ProgReportIntv As Integer = 100
    Public Event ProgressReport(ByVal ProgVal As Integer)
    Public Event ProgressStarted()
    Public Event ProgressFinished()
    Private ProgMonitor As Threading.Thread

End Class
