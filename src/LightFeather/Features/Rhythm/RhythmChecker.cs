using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using LightFeather.Extensions;
using Microsoft.Office.Interop.Word;

namespace LightFeather.Features.Rhythm {
	public class RhythmChecker {
		public static List<ChangedSentence> ChangedSentences = new List<ChangedSentence>();
		public static string PreviousParagraphText;
		public static bool UseComments;
		public static bool UseBackgroundChange;
		private static Timer Timer;

		public static void CheckRhythm() {
			Debug.WriteLine("[Rhythm] Check rhythm started.");
			var timer = new Timer
			{
				Interval = 300
			};
			timer.Tick += CheckRhythmInternal;
			timer.Start();
			Timer = timer;
			Debug.WriteLine("[Rhythm] Timer started.");
		}

		public static void DisableCheckRhythm() {
			if (Timer != null) {
				Timer.Tick -= CheckRhythmInternal;
				Timer.Stop();
				Timer.Dispose();
			}
			CleanupChangedSentences();
		}

		private static void CheckRhythmInternal(object sender, EventArgs e) {
			var currentSelection = Globals.ThisAddIn.Application.Selection;
			if (currentSelection == null) return;

			var currentParagraph = currentSelection.Paragraphs[1];
			if (PreviousParagraphText == null) {
				PreviousParagraphText = currentParagraph.Range.Text;
			}

			if (PreviousParagraphText != currentParagraph.Range.Text) {
				Debug.WriteLine("[Rhythm] Paragraph changed");
				CleanupChangedSentences();
				CheckRhythmForParagraph(currentParagraph);
				PreviousParagraphText = currentParagraph.Range.Text;
			}
		}

		private static void CheckRhythmForWholeDocument() {
			var document = Globals.ThisAddIn.Application.ActiveDocument;

			foreach (Paragraph paragraph in document.Paragraphs) {
				CheckRhythmForParagraph(paragraph);
			}
		}

		private static void CheckRhythmForParagraph(Paragraph paragraph) {
			int previousSentenceWordCount = 0;
			foreach (Range sentence in paragraph.Range.Sentences) {
				if (sentence?.Text == null || !sentence.Text.Trim().Any())
					continue;

				var words = sentence.GetActualWordsFromSentence();
				var count = words.Count();
				if (count == 0)
					continue;

				if (IsIncorrectRhythm(previousSentenceWordCount, count)) {
					var sentenceToEdit = sentence.Trim();
					MarkSentenceAsIncorrectRhythm(sentenceToEdit, count);
				}

				previousSentenceWordCount = count;
			}

			Debug.WriteLine("[Rhythm] Internal check. Sentences changed: " + ChangedSentences.Count);
		}

		private static void MarkSentenceAsIncorrectRhythm(Range sentenceToEdit, int count) {
			var changedSentence = new ChangedSentence() { Sentence = sentenceToEdit };
			if (UseBackgroundChange) changedSentence.PreviousBackgroundColor = ChangeBackgroundColor(sentenceToEdit);
			if (UseComments) changedSentence.Comment = AddComment(sentenceToEdit, count);
			ChangedSentences.Add(changedSentence);
		}

		private static Comment AddComment(Range sentence, int count) {
			object text = $"{count.ToString()} - {sentence.Text}";
			var comment = GetActiveDocument().Comments.Add(sentence, ref text);
			comment.ShowTip = true;
			comment.Author = GetCommentAuthor();
			return comment;
		}

		private static Document GetActiveDocument() {
			return Globals.ThisAddIn.Application.ActiveDocument;
		}

		private static string GetCommentAuthor() {
			return"Light Feather";
		}

		private static bool IsIncorrectRhythm(int previousSentenceWordCount, int count) {
			if (previousSentenceWordCount == 0) return false;
			var difference = Math.Abs(previousSentenceWordCount - count);
			return difference <= 2;
		}

		private static WdColor ChangeBackgroundColor(Range sentence) {
			var previousBackgroundColor = sentence.Shading.BackgroundPatternColor;
			sentence.Shading.BackgroundPatternColor = GetRhythmCheckerBackgroundColor();
			return previousBackgroundColor;
		}

		private static WdColor GetRhythmCheckerBackgroundColor() {
			return(WdColor)ColorTranslator.ToOle(Color.FromArgb(128, 255, 255, 0));
		}

		private static void CleanupChangedSentences() {
			Debug.WriteLine("[Rhythm] Cleanup started.");

			foreach (var changedSentence in ChangedSentences) {
				SafeCleanBackgroundColor(changedSentence);
				SafeDeleteComment(changedSentence.Comment);
			}

			ChangedSentences.Clear();
		}

		private static void SafeDeleteComment(Comment comment) {
			if (comment is null) return;

			try {
				comment?.DeleteRecursively();
			}
			catch (Exception e) {
				Debug.WriteLine($"Error: {e}");
			}
		}

		private static void SafeCleanBackgroundColor(ChangedSentence changedSentence) {
			if (changedSentence.PreviousBackgroundColor == null) return;
			try {
				changedSentence.Sentence.Shading.BackgroundPatternColor = changedSentence.PreviousBackgroundColor.Value;
			}
			catch (Exception e) {
				Debug.WriteLine($"Error: {e}");
			}
		}

		public static void CleanAllLeftovers() {
			var document = Globals.ThisAddIn.Application.ActiveDocument;

			foreach (Paragraph paragraph in document.Paragraphs) {
				if (paragraph.Shading.BackgroundPatternColor == GetRhythmCheckerBackgroundColor())
					paragraph.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;

				foreach (Range sentence in paragraph.Range.Sentences) {
					var trimmedSentence = sentence.Trim();
					if (trimmedSentence.Shading.BackgroundPatternColor == GetRhythmCheckerBackgroundColor())
						trimmedSentence.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
				}
			}

			CleanAllLeftoverComments();
		}

		private static void CleanAllLeftoverComments() {
			foreach (var comment in Globals.ThisAddIn.Application.ActiveDocument.Comments.Cast<Comment>()
				         .Where(comment => comment.Author == GetCommentAuthor())) {
				SafeDeleteComment(comment);
			}
		}
	}
}