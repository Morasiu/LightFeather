using LightFeather.Extensions;
using System;
using System.Diagnostics;
using System.Linq;
using LightFeather.Log;
using LightFeather.Shared;
using Microsoft.Office.Interop.Word;

namespace LightFeather.Features.Rhythm {
	public class RhythmChecker {
		public static ChangedSentences ChangedSentences = new ChangedSentences();
		public static string PreviousParagraphText;
		public static Paragraph PreviousParagraph;
		public static bool UseComments;

		private static readonly RhythmTimer Timer = new RhythmTimer();

		public static void CheckRhythm() {
			Debug.WriteLine("Check rhythm started.", LogCategories.Rhythm);

			Timer.Start(CheckRhythmInternal);
		}

		public static void DisableCheckRhythm() {
			Timer.Stop();

			ChangedSentences.CleanupChangedSentences();
		}

		public static void CleanAllLeftovers() {
			var undoAction = new UndoAction("Light feather - leftover cleanup");
			RhythmCleaner.CleanAll();
			undoAction.Dispose();
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
				ChangedSentences.CleanupChangedSentences();
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

			var undoAction = new UndoAction("Light Feather - Rhythm check");

			foreach (Range sentence in paragraph.Range.Sentences) {
				if (sentence?.Text == null || !sentence.Text.Trim().Any())
					continue;

				var wordCount = sentence.GetActualWordsFromSentence().Count();
				if (wordCount == 0)
					continue;

				var sentenceToEdit = sentence.Trim();
				if (IsIncorrectRhythm(previousSentenceWordCount, wordCount)) {
					MarkSentenceAsIncorrectRhythm(sentenceToEdit, wordCount);
					
				}
				else {
					if (UseComments) {
						var changedSentence = new ChangedSentence {
							Sentence = sentenceToEdit,
							Comment = CommentFactory.AddNeutralComment(sentenceToEdit, wordCount)
						};
						ChangedSentences.AddOrUpdate(changedSentence);
					}
				}

				previousSentenceWordCount = wordCount;
			}

			undoAction.Dispose();
			stopWatch.Stop();
			Debug.WriteLine(
				$"Internal check. Sentences changed: {ChangedSentences.Log.Count}. Took: {stopWatch.ElapsedMilliseconds}ms",
				LogCategories.Rhythm);
		}

		private static void MarkSentenceAsIncorrectRhythm(Range sentence, int count) {
			var changedSentence = new ChangedSentence {
				Sentence = sentence
			};

			if (UseComments) {
				changedSentence.PreviousUnderline = sentence.Underline;
				sentence.Underline = WdUnderline.wdUnderlineWavyHeavy;
				changedSentence.Comment = CommentFactory.AddIncorrectRhythmComment(sentence, count);
			}


			ChangedSentences.AddOrUpdate(changedSentence);
		}

		private static bool IsIncorrectRhythm(int previousSentenceWordCount, int count) {
			if (previousSentenceWordCount == 0) return false;
			var difference = Math.Abs(previousSentenceWordCount - count);
			return difference <= 2;
		}
	}
}