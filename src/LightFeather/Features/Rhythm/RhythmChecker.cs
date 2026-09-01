using LightFeather.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using LightFeather.Features.Rhythm.Panel;
using LightFeather.Log;
using LightFeather.Shared;
using Microsoft.Office.Interop.Word;

namespace LightFeather.Features.Rhythm {
	public class RhythmChecker {
		public static string PreviousParagraphText;
		public static Paragraph PreviousParagraph;
		public static bool UsePanel;

		private static readonly RhythmTimer Timer = new RhythmTimer();

		private static RhythmPanel _panel;

		public static void CheckRhythm() {
			Debug.WriteLine("Check rhythm started.", LogCategories.Rhythm);

			if (UsePanel) {
				OpenPanel();
			}

			Timer.Start(CheckRhythmInternal);
		}

		public static void DisableCheckRhythm() {
			Timer.Stop();

			ClosePanel();
		}

		public static void OpenPanel() {
			if (_panel != null)
				return;

			_panel = new RhythmPanel();
			_panel.FormClosed += (sender, args) => _panel = null;

			_panel.Show();
			_panel.AttachTo(CurrentDocument.GetApplication().ActiveWindow.GetOwnerHandle());

			PreviousParagraphText = null;
		}

		public static void ClosePanel() {
			if (_panel == null)
				return;

			_panel?.Close();
			_panel?.Dispose();
			_panel = null;
		}

		private static void CheckRhythmInternal(object sender, EventArgs e) {
			var stopWatch = Stopwatch.StartNew();
			var currentSelection = CurrentDocument.GetCurrentSelection();
			if (currentSelection == null) return;

			var currentParagraph = currentSelection.Paragraphs[1];
			if (currentParagraph.Range.Text.Trim().Length == 0) return;

			if (PreviousParagraph == null) PreviousParagraph = currentParagraph;

			if (PreviousParagraph.ParaID == currentParagraph.ParaID) {
				if (currentParagraph.Range.Text == PreviousParagraphText) return;

				Debug.WriteLine("Same paragraph, but changed.", LogCategories.Rhythm);
				CheckRhythmForParagraph(currentParagraph);
			}
			else {
				Debug.WriteLine("Selection switched to different paragraph", LogCategories.Rhythm);
				CheckRhythmForParagraph(currentParagraph);
				PreviousParagraph = currentParagraph;
			}

			PreviousParagraphText = currentParagraph.Range.Text;
			stopWatch.Stop();
			Debug.WriteLine($"Check rhythm took {stopWatch.ElapsedMilliseconds}ms", LogCategories.Rhythm);
		}

		private static void CheckRhythmForParagraph(Paragraph paragraph) {
			var stopWatch = Stopwatch.StartNew();
			var previousSentenceWordCount = 0;
			var sentenceRhythms = new List<SentenceRhythm>();

			foreach (Range sentence in paragraph.Range.Sentences) {
				if (sentence?.Text == null || !sentence.Text.Trim().Any())
					continue;

				var wordCount = sentence.GetActualWordsFromSentence().Count();
				if (wordCount == 0)
					continue;

				var incorrectRhythm = IsIncorrectRhythm(previousSentenceWordCount, wordCount);

				sentenceRhythms.Add(new SentenceRhythm(sentence.Trim(), wordCount, incorrectRhythm));

				previousSentenceWordCount = wordCount;
			}

			if (UsePanel) _panel?.SetSentences(sentenceRhythms);

			stopWatch.Stop();
			Debug.WriteLine(
				$"Internal check. Sentences: {sentenceRhythms.Count}. Took: {stopWatch.ElapsedMilliseconds}ms",
				LogCategories.Rhythm);
		}

		private static bool IsIncorrectRhythm(int previousSentenceWordCount, int count) {
			if (previousSentenceWordCount == 0) return false;
			var difference = Math.Abs(previousSentenceWordCount - count);
			return difference <= 2;
		}
	}
}
