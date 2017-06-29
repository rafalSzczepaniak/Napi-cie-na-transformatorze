Public Class Form1
    Dim isGood As Boolean = True
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        isGood = True
        Dim Rk, xk, rfe, xm
        funkcjaTransformatora(getSn, getU1phn, getIon, getPon, getPkn, getUk, getf, Rk, xk, rfe, xm)
        If (Not isGood) Then
            Return
        End If
        napisy(Rk, xk, rfe, xm)
        przecietnyWykres(Zmiana(-90, 90, 0.1))
    End Sub
    Private Sub funkcjaTransformatora(sn As Double, U1phn As Double, Ion As Double, Pon As Double, Pkn As Double, Uk As Double, f As Double, ByRef Rk As Double, ByRef xk As Double, ByRef Rfe As Double, ByRef xm As Double)
        Dim IoacphN, I1phn, IaphN, ImpnH, UkphN, UKRphN, UkXphN As Double


        IoacphN = Pon / (3 * U1phn)
        I1phn = sn / (3 * U1phn)
        IaphN = I1phn * (Ion / 100)
        ImpnH = Math.Sqrt(Math.Pow(IaphN, 2) - Math.Pow(IoacphN, 2))
        Rfe = U1phn / IoacphN
        xm = U1phn / ImpnH
        UkphN = U1phn * (Uk / 100)
        UKRphN = Pkn / (3 * I1phn)
        UkXphN = Math.Sqrt(Math.Pow(UkphN, 2) - Math.Pow(UKRphN, 2))
        Rk = UKRphN / I1phn
        xk = UkXphN / I1phn

    End Sub
    Private Sub funkcjaTransformatora(sn As Double, U1phn As Double, Ion As Double, Pon As Double, Pkn As Double, Uk As Double, f As Double, ByRef Rk As Double, ByRef xk As Double, ByRef Rfe As Double, ByRef xm As Double, ByRef ukphn1 As Double)
        Dim IoacphN, I1phn, IaphN, ImpnH, UkphN, UKRphN, UkXphN As Double


        IoacphN = Pon / (3 * U1phn)
        I1phn = sn / (3 * U1phn)
        IaphN = I1phn * (Ion / 100)
        ImpnH = Math.Sqrt(Math.Pow(IaphN, 2) - Math.Pow(IoacphN, 2))
        Rfe = U1phn / IoacphN
        xm = U1phn / ImpnH
        UkphN = U1phn * (Uk / 100)
        UKRphN = Pkn / (3 * I1phn)
        UkXphN = Math.Sqrt(Math.Pow(UkphN, 2) - Math.Pow(UKRphN, 2))
        Rk = UKRphN / I1phn
        xk = UkXphN / I1phn
        ukphn1 = I1phn
    End Sub
    Private Function Zmiana(ByVal poczatekZakresuZmiennosci As Double, ByVal koniecZakresuZmiennosi As Double, ByVal Krok As Double) As List(Of Double)
        Dim iloscKrokow As Integer
        iloscKrokow = Math.Floor((koniecZakresuZmiennosi - poczatekZakresuZmiennosci) / Krok)
        Dim x As Double
        Dim wynik As New List(Of Double)
        For i As Integer = 0 To iloscKrokow
            x = poczatekZakresuZmiennosci + i * Krok


            Dim Z, Ip, Rw, Xw, rfe, xm, kat, Rz, Xz, U As Double
            Z = 30
            kat = x * Math.PI / 180
            funkcjaTransformatora(getSn, getU1phn, getIon, getPon, getPkn, getUk, getf, Rw, Xw, rfe, xm, Ip)
            Rz = Z * Math.Cos(kat)
            Xz = Z * Math.Sin(kat)
            U = Rz * Ip + Xz * Ip
            wynik.Add(U)

        Next
        Return wynik
    End Function
    Private Sub przecietnyWykres(ByVal wartosci As List(Of Double))
        Dim Top, Left, Right, Bottom, x1, y1, x2, y2, ox, rozmiarXStrzalki, rozmiarYStrzalki, oxr As Integer
        Top = 50
        Bottom = 500
        Left = 250
        Right = 1000
        rozmiarXStrzalki = (Right - Left) / 25
        rozmiarYStrzalki = (Bottom - Top) / 25
        Dim skala, krokow, krok As Double
        skala = (Bottom - Top) / (wartosci.Max - wartosci.Min)
        krokow = Math.Ceiling((Right - Left) / wartosci.Count)
        krok = (Right - Left) / wartosci.Count
        ox = wartosci.Min * skala
        If (ox > 0) Then
            oxr = 0
        Else
            oxr = ox
        End If

        Dim graph As Graphics
        graph = Me.CreateGraphics
        graph.Clear(BackColor)
        Dim p As New Pen(Color.Black, 2)
        graph.DrawLine(p, Left, Top, Left, Bottom)
        graph.DrawLine(p, Left, Top, Right, Top)
        graph.DrawLine(p, Left, Bottom, Right, Bottom)
        graph.DrawLine(p, Right, Top, Right, Bottom)

        graph.DrawLine(p, Left, Bottom + oxr, Right, Bottom + oxr)
        graph.DrawLine(p, Right, Bottom + oxr, Right - rozmiarXStrzalki, Bottom + oxr - rozmiarYStrzalki)
        graph.DrawLine(p, Right, Bottom + oxr, Right - rozmiarXStrzalki, Bottom + oxr + rozmiarYStrzalki)

        p.Color = Color.Red
        For x As Integer = 0 To wartosci.Count - 2
            x1 = Left + x * krok
            y1 = Math.Round(Bottom + ox - wartosci(x) * skala)
            x2 = Left + (x + 1) * krok
            y2 = Math.Round(Bottom + ox - wartosci(x + 1) * skala)
            graph.DrawLine(p, x1, y1, x2, y2)
        Next

    End Sub
    Private Sub napisy(Rk As Double, Xk As Double, rfe As Double, xm As Double)
        Label8.Text = "Rk " & Math.Round(Rk, 2) & " Ohm"
        Label9.Text = "Xk " & Math.Round(Xk, 2) & " Ohm"
        Label10.Text = "Rfe " & Math.Round(rfe, 2) & " Ohm"
        Label11.Text = "Xm " & Math.Round(xm, 2) & " Ohm"
    End Sub
    Private Function getSn() As Double
        Dim x As Double
        Try
            x = Double.Parse(TextBox1.Text)
        Catch ex As Exception
            MessageBox.Show("Wystąpił błąd z pobraniem wartości Sn. Sprawdź to!")
            isGood = False
        End Try
        Return x
    End Function
    Private Function getU1phn() As Double
        Dim x As Double
        Try
            x = Double.Parse(TextBox2.Text)
        Catch ex As Exception
            MessageBox.Show("Wystąpił błąd z pobraniem wartości U1phn. Sprawdź to!")
            isGood = False
        End Try
        Return x
    End Function
    Private Function getIon() As Double
        Dim x As Double
        Try
            x = Double.Parse(TextBox5.Text)
        Catch ex As Exception
            MessageBox.Show("Wystąpił błąd z pobraniem wartości Ion. Sprawdź to!")
            isGood = False
        End Try
        Return x
    End Function
    Private Function getPon() As Double
        Dim x As Double
        Try
            x = Double.Parse(TextBox4.Text)
        Catch ex As Exception
            MessageBox.Show("Wystąpił błąd z pobraniem wartości Pon. Sprawdź to!")
            isGood = False
        End Try
        Return x
    End Function
    Private Function getPkn() As Double
        Dim x As Double
        Try
            x = Double.Parse(TextBox7.Text)
        Catch ex As Exception
            MessageBox.Show("Wystąpił błąd z pobraniem wartości Pkn. Sprawdź to!")
            isGood = False
        End Try
        Return x
    End Function
    Private Function getUk() As Double
        Dim x As Double
        Try
            x = Double.Parse(TextBox6.Text)
        Catch ex As Exception
            MessageBox.Show("Wystąpił błąd z pobraniem wartości Uk. Sprawdź to!")
            isGood = False
        End Try
        Return x
    End Function
    Private Function getf() As Double
        Dim x As Double
        Try
            x = Double.Parse(TextBox2.Text)
        Catch ex As Exception
            MessageBox.Show("Wystąpił błąd z pobraniem wartości f. Sprawdź to!")
            isGood = False

        End Try
        Return x
    End Function
End Class
