using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace LightFeather
{
	public class RhythmChecker
	{
		public static List<ChangedSentence> ChangedSentences = new List<ChangedSentence>();

		public static string CurrentParagraphText;

		public static void CheckRhythm()
		{
			Debug.WriteLine("[Rhythm] Check rhythm started.");
			var timer = new Timer
			{
				Interval = 1000
			};
			timer.Tick += (o, e) => CheckRhythmInternal();
			timer.Start();
			Debug.WriteLine("[Rhythm] Timer started.");
		}

		private static void CheckRhythmInternal()
		{

			//CheckRhythmForWholeDocument();
			var selection = Globals.ThisAddIn.Application.Selection;
			if (selection == null) return;
			var paragprah = selection.Paragraphs[1];
			if (CurrentParagraphText == null) {
				Debug.WriteLine("[Rhythm] Check rhythm internal tick. First check.");
				CheckRhythmForParagraph(paragprah);
				CurrentParagraphText = paragprah.Range.Text;
			}
			else if (CurrentParagraphText != paragprah.Range.Text)
			{
				Debug.WriteLine("[Rhythm] Paragraph changed");
				CleanupChangedSentences();
				CheckRhythmForParagraph(paragprah);
				CurrentParagraphText = paragprah.Range.Text;
			} else
			{
				Debug.WriteLine("[Rhythm] Check rhythm internal tick. No changes.");
			}
		}

		private static void CheckRhythmForWholeDocument()
		{
			var document = Globals.ThisAddIn.Application.ActiveDocument;

			foreach (Paragraph paragraph in document.Paragraphs)
			{
				CheckRhythmForParagraph(paragraph);
			}
		}

		private static void CheckRhythmForParagraph(Paragraph paragraph)
		{
			Debug.WriteLine("[Rhythm] Internal check. Sentence changed: " + ChangedSentences.Count);
			Debug.WriteLine("Paragpraph. Sentences: " + paragraph.Range.Sentences.Count);
			int previousSenteceWordCount = 0;
			foreach (Range sentence in paragraph.Range.Sentences)
			{
				if (sentence.Text.Trim().Count() == 0) continue;
				var words = sentence.Words.Cast<Range>().Where(x => x.Text.Trim().Length > 0 && !char.IsPunctuation(x.Text.Trim()[0])).Select(x => x);
				var count = words.Count();
				if (count == 0) continue;
				if (IsIncorrctRhythm(previousSenteceWordCount, count))
				{
					//ChangeBackgroundColor(sentence);
					object text = $"{count.ToString()} - {sentence.Text}";
					var comment = Globals.ThisAddIn.Application.ActiveDocument.Comments.Add(sentence, ref text);
					comment.ShowTip = true;
					comment.Author = "Light Feather";
					ChangedSentences.Add(new ChangedSentence() { Sentence = sentence, PreviousBackgroundColor = sentence.Shading.BackgroundPatternColor, Comment = comment });
					//var indexRange = Globals.ThisAddIn.Application.ActiveDocument.Range(sentence.End, sentence.End);
					//indexRange.Font.Superscript = 1;
					//indexRange.Select();
					//indexRange.InsertAfter(count.ToString());
					//indexRange.Font.Superscript = 0;
				}
				previousSenteceWordCount = count;
			}
		}

		private static bool IsIncorrctRhythm(int previousSenteceWordCount, int count)
		{
			if (previousSenteceWordCount == 0) return false;
			var differencce = Math.Abs(previousSenteceWordCount - count);
			if (differencce <= 2)
			{
				return true;
			}
			return false;
		}

		private static void ChangeBackgroundColor(Range sentence)
		{
			sentence.Shading.BackgroundPatternColor = (WdColor)ColorTranslator.ToOle(Color.FromArgb(128, 255, 255, 0));
		}

		public static void DisableCheckRhtyhm()
		{
			CleanupChangedSentences();
		}

		private static void CleanupChangedSentences()
		{
			Debug.WriteLine("[Rhythm] Cleanup started.");
			foreach (var changedSentence in ChangedSentences)
			{
				try
				{
					changedSentence.Sentence.Shading.BackgroundPatternColor = changedSentence.PreviousBackgroundColor;
				} catch (Exception e)
				{
					Debug.WriteLine($"Error: {e}");
				}
				try
				{
					changedSentence.Comment?.DeleteRecursively();
				}
				catch (Exception e)
				{
					Debug.WriteLine($"Error: {e}");
				}
			}

			ChangedSentences.Clear();
		}
	}

	public class ChangedSentence
	{
		public Range Sentence;
		public WdColor PreviousBackgroundColor;
		public Comment Comment;
	}
}
