Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Windows

Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Services
Imports System.Globalization
Imports DevExpress.XtraRichEdit.Utils.NumberConverters
Imports DevExpress.XtraRichEdit.Utils

Namespace Expander
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub richEditControl1_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			richEditControl1.ApplyTemplate()
			richEditControl1.CreateNewDocument()
			AddHandler richEditControl1.AutoCorrect, AddressOf richEditControl1_AutoCorrect
		End Sub

		Private Function CreateImageFromResx(ByVal name As String) As Image
			Dim im As Image
			Dim [assembly] As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
			Using stream As Stream = [assembly].GetManifestResourceStream("AutoCorrectEvent.Images." & name)
				im = Image.FromStream(stream)
			End Using
			Return im
		End Function
		#Region "#autocorrect"
		Private Sub richEditControl1_AutoCorrect(ByVal sender As Object, ByVal e As DevExpress.XtraRichEdit.AutoCorrectEventArgs)
			Dim info As AutoCorrectInfo = e.AutoCorrectInfo
			e.AutoCorrectInfo = Nothing

			If info.Text.Length <= 0 Then
				Return
			End If
			Do
				If (Not info.DecrementStartPosition()) Then
					Return
				End If

				If IsSeparator(info.Text(0)) Then
					Return
				End If

				If info.Text(0) = "$"c Then
					info.ReplaceWith = CreateImageFromResx("dollar_pic.png")
					e.AutoCorrectInfo = info
					Return
				End If

				If info.Text(0) = "%"c Then
					Dim replaceString As String = CalculateFunction(info.Text)
					If (Not String.IsNullOrEmpty(replaceString)) Then
						info.ReplaceWith = replaceString
						e.AutoCorrectInfo = info
					End If
					Return
				End If
			Loop
		End Sub
		#End Region ' #autocorrect
		Private Function CalculateFunction(ByVal name As String) As String
			name = name.ToLower()

			If name.Length > 2 AndAlso name.Chars(0) = "%"c AndAlso name.EndsWith("%") Then
				Dim value As Integer
				If Int32.TryParse(name.Substring(1, name.Length - 2), value) Then
					Dim converter As OrdinalBasedNumberConverter = OrdinalBasedNumberConverter.CreateConverter(DevExpress.XtraRichEdit.Model.NumberingFormat.CardinalText, LanguageId.English)
					Return converter.ConvertNumber(value)
				End If
			End If

			Select Case name
				Case "%date%"
					Return DateTime.Now.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern)
				Case "%time%"
					Return DateTime.Now.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern)
				Case Else
					Return String.Empty
			End Select
		End Function
		Private Function IsSeparator(ByVal ch As Char) As Boolean
			Return ch <> "%"c AndAlso (ch = ControlChars.Cr OrElse ch = ControlChars.Lf OrElse Char.IsPunctuation(ch) OrElse Char.IsSeparator(ch) OrElse Char.IsWhiteSpace(ch))
		End Function

	End Class
End Namespace
