Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim DSFFile1 As DSFFile = DSFFile.FromFile(TextBox1.Text)
        If DSFFile1 IsNot Nothing Then
            DSFFile1.Taps = NumericUpDown2.Value
            DSFFile1.PCMSampleFrequency = NumericUpDown1.Value
            DSFFile1.fFreqHigh = DSFFile1.PCMSampleFrequency / 2
            AddHandler DSFFile1.ProgressReport,
                Sub(ProgVal As Integer)
                    Invoke(Sub() Text = ProgVal / 100 & "%")
                End Sub
            AddHandler DSFFile1.ProgressStarted,
                Sub()
                    Invoke(Sub()
                               Button1.Enabled = False
                               Button2.Enabled = False
                           End Sub)
                End Sub
            AddHandler DSFFile1.ProgressFinished,
                Sub()
                    Invoke(Sub()
                               Button1.Enabled = True
                               Button2.Enabled = True
                           End Sub)
                End Sub
            DSFFile1.ToWaveFile(TextBox2.Text, NumericUpDown1.Value)
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim DSFFile1 As DSFFile = DSFFile.FromFile(TextBox1.Text)
        If DSFFile1 IsNot Nothing Then
            DSFFile1.Taps = NumericUpDown2.Value
            AddHandler DSFFile1.ProgressReport,
                Sub(ProgVal As Integer)
                    Invoke(Sub() Text = ProgVal / 100 & "%")
                End Sub
            AddHandler DSFFile1.ProgressStarted,
                Sub()
                    Invoke(Sub()
                               Button1.Enabled = False
                               Button2.Enabled = False
                           End Sub)
                End Sub
            AddHandler DSFFile1.ProgressFinished,
                Sub()
                    Invoke(Sub()
                               Button1.Enabled = True
                               Button2.Enabled = True
                           End Sub)
                End Sub
            DSFFile1.ToTextFile(TextBox2.Text, NumericUpDown1.Value)
        End If
    End Sub
End Class
