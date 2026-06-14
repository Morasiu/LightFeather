using LightFeather.Extensions;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using LightFeather.Shared;

namespace LightFeather.Features.Rhythm {
	public class RhythmChecker {
		public static List<ChangedSentence> ChangedSentences = new List<ChangedSentence>();
		public static string PreviousParagraphText;
		public static Paragraph PreviousParagraph;
		public static bool UseComments;
		public static bool UseBackgroundChange;

		private static readonly RhythmTimer Timer = new RhythmTimer();

		public static void CheckRhythm() {
			Debug.WriteLine("[Rhythm] Check rhythm started.");

			Timer.Start(CheckRhythmInternal);

			Debug.WriteLine("[Rhythm] Timer started.");
		}

		public static void DisableCheckRhythm() {
			Timer.Stop();

			CleanupChangedSentences();
		}

		public static void CleanAllLeftovers() {
			CleanLeftoverTextStyleManipulation();
			CleanAllLeftoverComments();
		}

		private static void CleanLeftoverTextStyleManipulation() {
			var document = CurrentDocument.GetActiveDocument();
			foreach (Paragraph paragraph in document.Paragraphs) {
				CleanLeftoverTextStyleManipulationInParagraph(paragraph);
			}
		}

		private static void CleanLeftoverTextStyleManipulationInParagraph(Paragraph paragraph) {
			foreach (Range sentence in paragraph.Range.Sentences) {
				var trimmedSentence = sentence.Trim();

				CleanLeftoverBackgroundColorManipulation(trimmedSentence);
				CleanLeftoverUnderlineManipulation(trimmedSentence);
			}
		}

		private static void CleanLeftoverUnderlineManipulation(Range trimmedSentence) {
			if (trimmedSentence.Underline != WdUnderline.wdUnderlineWavyHeavy) return;

			trimmedSentence.Underline = WdUnderline.wdUnderlineNone;
		}

		private static void CleanLeftoverBackgroundColorManipulation(Range trimmedSentence) {
			if (trimmedSentence.Shading.BackgroundPatternColor != RhythmConsts.IncorrectRhythmBackgroundColor) return;

			trimmedSentence.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
		}

		private static void CheckRhythmInternal(object sender, EventArgs e) {
			var currentSelection = CurrentDocument.GetCurrentSelection();
			if (currentSelection == null) return;

			var currentParagraph = currentSelection.Paragraphs[1];
			if (currentParagraph.Range.Text.Trim().Length == 0) return;

			if (PreviousParagraph == null) PreviousParagraph = currentParagraph;

			if (PreviousParagraph.ParaID == currentParagraph.ParaID) {
				if (currentParagraph.Range.Text == PreviousParagraphText) return;

				Debug.WriteLine("[Rhythm] Same paragraph, but changed.");
				CheckRhythmForParagraph(currentParagraph);
			}
			else {
				Debug.WriteLine("[Rhythm] Selection switched to different paragraph");
				CleanupChangedSentences();
				CheckRhythmForParagraph(currentParagraph);
				PreviousParagraph = currentParagraph;
			}

			PreviousParagraphText = currentParagraph.Range.Text;
		}

		private static void CheckRhythmForParagraph(Paragraph paragraph) {
			var previousSentenceWordCount = 0;

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
						AddOrUpdateChangedSentence(changedSentence);
					}
				}

				previousSentenceWordCount = wordCount;
			}

			Debug.WriteLine("[Rhythm] Internal check. Sentences changed: " + ChangedSentences.Count);
		}

		private static void AddOrUpdateChangedSentence(ChangedSentence changedSentence) {
			var existingChangedSentence = ChangedSentences.
				FirstOrDefault(x => x.Sentence.TextEqualTo(changedSentence.Sentence));
			if (existingChangedSentence == null) {
				ChangedSentences.Add(changedSentence);
			}
			else {
				if (changedSentence.PreviousBackgroundColor == null) changedSentence.SafeCleanBackgroundColor();
				existingChangedSentence.PreviousBackgroundColor = changedSentence.PreviousBackgroundColor;

				if (changedSentence.PreviousUnderline == null) existingChangedSentence.SafeCleanUnderline();
				existingChangedSentence.PreviousUnderline = changedSentence.PreviousUnderline;

				if (changedSentence.Comment == null) existingChangedSentence.Comment.SafeDelete();
				existingChangedSentence.Comment = changedSentence.Comment;
			}
		}

		private static void MarkSentenceAsIncorrectRhythm(Range sentence, int count) {
			var changedSentence = new ChangedSentence() {
				Sentence = sentence
			};
			if (UseBackgroundChange) changedSentence.PreviousBackgroundColor = ChangeBackgroundColor(sentence);

			if (UseComments) {
				changedSentence.PreviousUnderline = sentence.Underline;
				sentence.Underline = WdUnderline.wdUnderlineWavyHeavy;
				changedSentence.Comment = CommentFactory.AddIncorrectRhythmComment(sentence, count);
			}

			AddOrUpdateChangedSentence(changedSentence);
		}

		
		private static bool IsIncorrectRhythm(int previousSentenceWordCount, int count) {
			if (previousSentenceWordCount == 0) return false;
			var difference = Math.Abs(previousSentenceWordCount - count);
			return difference <= 2;
		}

		private static WdColor ChangeBackgroundColor(Range sentence) {
			var previousBackgroundColor = sentence.Shading.BackgroundPatternColor;
			sentence.Shading.BackgroundPatternColor = RhythmConsts.IncorrectRhythmBackgroundColor;
			return previousBackgroundColor;
		}

		private static void CleanupChangedSentences() {
			Debug.WriteLine("[Rhythm] Cleanup started.");

			foreach (var changedSentence in ChangedSentences) {
				changedSentence.SafeCleanBackgroundColor();
				changedSentence.SafeCleanUnderline();
				changedSentence.Comment.SafeDelete();
			}

			ChangedSentences.Clear();
		}



		private static void CleanAllLeftoverComments() {
			foreach (var comment in CurrentDocument.GetActiveDocument().Comments.FilterMadeByLightFeather()) {
				comment.SafeDelete();
			}
		}
	}
}