Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Windows
Imports DevExpress.XtraRichEdit
Imports System.Globalization
Imports DevExpress.XtraRichEdit.Utils.NumberConverters
Imports DevExpress.XtraRichEdit.Utils

Namespace Expander

    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Public Partial Class MainWindow
        Inherits Window

        Public Sub New()
            Me.InitializeComponent()
        End Sub

        Private Sub richEditControl1_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.richEditControl1.ApplyTemplate()
            Me.richEditControl1.CreateNewDocument()
            AddHandler Me.richEditControl1.AutoCorrect, New AutoCorrectEventHandler(AddressOf richEditControl1_AutoCorrect)
        End Sub

        Private Function CreateImageFromResx(ByVal name As String) As Image
            Dim im As Image
            Dim assembly As Assembly = Assembly.GetExecutingAssembly()
            Using stream As Stream = assembly.GetManifestResourceStream("AutoCorrectEvent.Images." & name)
                im = Image.FromStream(stream)
            End Using

            Return im
        End Function

'#Region "#autocorrect"
        Private Sub richEditControl1_AutoCorrect(ByVal sender As Object, ByVal e As AutoCorrectEventArgs)
            Dim info As AutoCorrectInfo = e.AutoCorrectInfo
            e.AutoCorrectInfo = Nothing
            If info.Text.Length <= 0 Then Return
            While True
                If Not info.DecrementStartPosition() Then Return
                If IsSeparator(info.Text(0)) Then Return
                If info.Text(0) = "$"c Then
                    info.ReplaceWith = CreateImageFromResx("dollar_pic.png")
                    e.AutoCorrectInfo = info
                    Return
                End If

                If info.Text(0) = "%"c Then
                    Dim replaceString As String = CalculateFunction(info.Text)
                    If Not String.IsNullOrEmpty(replaceString) Then
                        info.ReplaceWith = replaceString
                        e.AutoCorrectInfo = info
                    End If

                    Return
                End If
            End While
        End Sub

'#End Region  ' #autocorrect
        Private Function CalculateFunction(ByVal name As String) As String
            name = name.ToLower()
            If name.Length > 2 AndAlso name(0) = "%"c AndAlso name.EndsWith("%") Then
                Dim value As Integer
                If Integer.TryParse(name.Substring(1, name.Length - 2), value) Then
                    Dim converter As OrdinalBasedNumberConverter = OrdinalBasedNumberConverter.CreateConverter(Model.NumberingFormat.CardinalText, LanguageId.English)
                    Return converter.ConvertNumber(value)
                End If
            End If

            Select Case name
                Case "%date%"
                    Return Date.Now.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern)
                Case "%time%"
                    Return Date.Now.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern)
                Case Else
                    Return String.Empty
            End Select
        End Function

        Private Function IsSeparator(ByVal ch As Char) As Boolean
            Return ch <> "%"c AndAlso (ch = Microsoft.VisualBasic.Strings.ChrW(13) OrElse ch = Microsoft.VisualBasic.Strings.ChrW(10) OrElse Char.IsPunctuation(ch) OrElse Char.IsSeparator(ch) OrElse Char.IsWhiteSpace(ch))
        End Function
    End Class
End Namespace
