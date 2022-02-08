Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms

Namespace WordCloudGenerator
    Public Partial Class Form1
        Inherits Form

        Private keyWords As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)()
        Private blackList As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)()
        Private totalWords As Integer = 0

        Public Sub New()
            Me.InitializeComponent()
            LoadBlacklist()
        End Sub

        ' Eventi
        Private Sub btnBrowse_Click(ByVal sender As Object, ByVal e As EventArgs)
            BrowseAndLoad()
        End Sub

        Private Sub btnGenerate_Click(ByVal sender As Object, ByVal e As EventArgs)
            GenerateWordCloud2()
        End Sub

        ' Metodi
        Private Sub LoadBlacklist()
            ' Carica l'elenco delle parole da non conteggiare
            Dim lines = File.ReadAllLines("Dati\Black List.txt")

            For Each line In lines
                If Not blackList.ContainsKey(line) Then blackList.Add(line, 0)
            Next
        End Sub

        Private Sub BrowseAndLoad()

            ' Se la parola priva di spazi 1) non compare nella BlackList, 2) ha una lunghezza >=2 e 3) non rappresenta un numero
            ' essa iene aggiunta al dizionario e il relativo contatore viene incrementato di 1.
            Dim n As Decimal = Nothing

            If Me.openFileDialog1.ShowDialog() = DialogResult.OK Then
                Try
                    Me.txtFileName.Text = Me.openFileDialog1.FileName
                    Dim _keyWords = New Dictionary(Of String, Integer)()

                    ' Carica le linee di testo del file di testo selezionato
                    Dim lines = File.ReadLines(Me.openFileDialog1.FileName)
                    ' Cicla su ciascuna linea
                    For Each line In lines
                        ' rimpiazza alcuni caratteri speciali con uno spazio
                        Dim cleanLine = line.Replace(":"c, " "c)
                        cleanLine = cleanLine.Replace(";"c, " "c)
                        cleanLine = cleanLine.Replace(","c, " "c)
                        cleanLine = cleanLine.Replace("."c, " "c)
                        cleanLine = cleanLine.Replace("_"c, " "c)
                        cleanLine = cleanLine.Replace("="c, " "c)
                        cleanLine = cleanLine.Replace("*"c, " "c)
                        cleanLine = cleanLine.Replace("'"c, " "c)



                        ' Separa le parole della linea depurata dai caratteri speciali
                        Dim words = cleanLine.Split(" "c)

                        ' scorre l'elenco delle parole trovate 
                        For Each word In words
                            ' Elimina gli spazi e forza il minuscolo
                            Dim cleanWord = word.ToLower().Trim()

                            If Not blackList.ContainsKey(cleanWord) AndAlso word.Length >= 2 AndAlso Not Decimal.TryParse(word, n) Then
                                If Not _keyWords.ContainsKey(cleanWord) Then _keyWords.Add(cleanWord, 0)
                                _keyWords(cleanWord) += 1
                            End If
                        Next
                    Next

                    ' Filtra solo le parole comparse più di una volta per creare il dizionario finale
                    keyWords = (From item In _keyWords Where item.Value > 1 Select item).ToDictionary(Function(x) x.Key, Function(x) x.Value)

                    ' Imposta il numero totale di parole 
                    totalWords = keyWords.Count()
                Catch ex As Exception
                    MessageBox.Show($"Errore.

Error message: {ex.Message}

" & $"Details:

{ex.StackTrace}")
                End Try
            End If
        End Sub

        Private Sub GenerateWordCloud()
            Me.richTextBox1.Text = String.Empty

            ' Calcol gli elementi con il minimo e il massimo delle occorrene e divide per 5 per creare 5 classi di appartenenza
            Dim minOccur = (From item In keyWords Select item).Min(Function(x) x.Value)
            Dim maxOccur = (From item In keyWords Select item).Max(Function(x) x.Value)
            Dim splitter = (maxOccur - minOccur) / 5

            ' Cicla su tutte le coppie (parola, contatore) del dizionario
            For Each word In keyWords
                Dim fontSize = 8
                Dim fontStyle = Drawing.FontStyle.Regular
                Dim color = Drawing.Color.Black

                ' Controlla il numero di volte che ogni parola compare nel testo e in funzione di esso assegna una dimensione e un colore 
                If word.Value >= splitter * 3 Then
                    fontSize = 48
                    fontStyle = FontStyle.Bold
                    color = Color.Orange
                ElseIf word.Value >= splitter * 2 Then
                    fontSize = 24
                    fontStyle = FontStyle.Italic
                    color = Color.Red
                ElseIf word.Value >= splitter Then
                    fontSize = 20
                    fontStyle = FontStyle.Underline
                    color = Color.Blue
                ElseIf word.Value >= Math.Round(splitter / 2, 0) Then
                    fontSize = 16
                    fontStyle = FontStyle.Regular
                    color = Color.LightGreen
                ElseIf word.Value >= Math.Round(splitter / 3, 0) Then
                    fontSize = 12
                    fontStyle = FontStyle.Regular
                    color = Color.Cyan
                End If

                ' Inserisce la parola nel RichTextBox
                Me.richTextBox1.AppendText(word.Key)
                Me.richTextBox1.AppendText(" ")

                ' Seleziona la parola e cambia dimensione e colore
                Me.richTextBox1.Select(Me.richTextBox1.Text.Length - word.Key.Trim().Length - 1, word.Key.Trim().Length)
                Me.richTextBox1.SelectionFont = New Font("Tahoma", fontSize, fontStyle)
                Me.richTextBox1.SelectionColor = color
            Next
        End Sub

        Private Sub GenerateWordCloud2()
            Dim fontSizes = New Integer() {8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42, 44, 46, 48}
            Dim fontColors = New Color() {Color.Black, Color.PaleVioletRed, Color.CadetBlue, Color.Coral, Color.DarkCyan, Color.DarkKhaki, Color.DarkOrange, Color.DarkSlateGray, Color.DarkOliveGreen, Color.IndianRed, Color.OrangeRed, Color.LightPink, Color.LightSalmon, Color.LightGreen, Color.LightBlue, Color.Tomato, Color.Turquoise, Color.Violet, Color.Plum, Color.YellowGreen, Color.Sienna, Color.Silver}
            Me.richTextBox1.Text = String.Empty

            ' Calcola gli elementi con il minimo e il massimo delle occorrene e divide per il numero di dimensioni possibili per creare altrettante classi di appartenenza
            Dim minOccur = (From item In keyWords Select item).Min(Function(x) x.Value)
            Dim maxOccur = (From item In keyWords Select item).Max(Function(x) x.Value)
            Dim splitter = (maxOccur - minOccur) / fontSizes.Count()
            Dim styles As List(Of Tuple(Of Integer, Integer, Integer)) = New List(Of Tuple(Of Integer, Integer, Integer))()
            styles.Add(New Tuple(Of Integer, Integer, Integer)(0, 0, splitter))
            Dim i = 1

            While i < fontSizes.Count() - 1
                styles.Add(New Tuple(Of Integer, Integer, Integer)(i, splitter * i, splitter * (i + 1)))
                i += 1
            End While

            styles.Add(New Tuple(Of Integer, Integer, Integer)(i, splitter * i, maxOccur))

            ' Cicla su tutte le coppie (parola, contatore) del dizionario
            For Each word In keyWords
                ' Controlla il numero di volte che ogni parola compare nel testo e in funzione di esso assegna una dimensione e un colore 
                Dim style = styles.Where(Function(x) x.Item2 < word.Value AndAlso x.Item3 >= word.Value).FirstOrDefault()
                Dim fontSize = fontSizes(style.Item1)
                Dim color = fontColors(style.Item1)
                Dim fontStyle = Drawing.FontStyle.Regular

                ' Inserisce la parola nel RichTextBox
                Me.richTextBox1.AppendText(word.Key.Trim())
                Me.richTextBox1.AppendText(" ")

                ' Seleziona la parola e cambia dimensione e colore
                Me.richTextBox1.Select(Me.richTextBox1.Text.Length - word.Key.Trim().Length - 1, word.Key.Trim().Length)
                Me.richTextBox1.SelectionFont = New Font("Tahoma", fontSize, fontStyle)
                Me.richTextBox1.SelectionColor = color
            Next
        End Sub
    End Class
End Namespace
