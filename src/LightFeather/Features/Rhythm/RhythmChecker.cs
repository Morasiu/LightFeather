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
		private static Timer _timer;

		public static void CheckRhythm() {
			Debug.WriteLine("[Rhythm] Check rhythm started.");
			var timer = new Timer {
				Interval = 500
			};
			timer.Tick += CheckRhythmInternal;
			timer.Start();
			_timer = timer;
			Debug.WriteLine("[Rhythm] Timer started.");
		}

		public static void DisableCheckRhythm() {
			if (_timer != null) {
				_timer.Tick -= CheckRhythmInternal;
				_timer.Stop();
				_timer.Dispose();
			}

			CleanupChangedSentences();
		}

		private static void CheckRhythmInternal(object sender, EventArgs e) {
			var currentSelection = Globals.ThisAddIn.Application.Selection;
			if (currentSelection == null) return;

			var currentParagraph = currentSelection.Paragraphs[1];
			if (currentParagraph.Range.Text.Trim().Length == 0) return;
			if (PreviousParagraph == null) {
				PreviousParagraph = currentParagraph;
			}

			if (PreviousParagraph.ParaID == currentParagraph.ParaID) {
				if (currentParagraph.Range.Text == PreviousParagraphText) {
					return;
				}

				Debug.WriteLine("[Rhythm] Same paragraph, but changed.");
				CheckRhythmForParagraph(currentParagraph);

				PreviousParagraphText = currentParagraph.Range.Text;
			}
			else {
				Debug.WriteLine("[Rhythm] Selection switched to different paragraph");
				CleanupChangedSentences();
				CheckRhythmForParagraph(currentParagraph);
				PreviousParagraph = currentParagraph;
				PreviousParagraphText = currentParagraph.Range.Text;
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

				var sentenceToEdit = sentence.Trim();
				if (IsIncorrectRhythm(previousSentenceWordCount, count)) {
					MarkSentenceAsIncorrectRhythm(sentenceToEdit, count);
				}
				else {
					if (UseComments) {
						var changedSentence = new ChangedSentence {
							Sentence = sentenceToEdit,
							Comment = AddNeutralComment(sentenceToEdit, count)
						};
						AddOrUpdateChangedSentence(changedSentence);
					}
				}

				previousSentenceWordCount = count;
			}

			Debug.WriteLine("[Rhythm] Internal check. Sentences changed: " + ChangedSentences.Count);
		}

		private static void AddOrUpdateChangedSentence(ChangedSentence changedSentence) {
			var existingChangedSentence =
				ChangedSentences.FirstOrDefault(x => x.Sentence.Trim().Text == changedSentence.Sentence.Trim().Text);
			if (existingChangedSentence == null) {
				ChangedSentences.Add(changedSentence);
			}
			else {
				if (changedSentence.PreviousBackgroundColor == null) {
					SafeCleanBackgroundColor(existingChangedSentence);
				}

				existingChangedSentence.PreviousBackgroundColor = changedSentence.PreviousBackgroundColor;

				if (changedSentence.PreviousUnderline == null) {
					SafeCleanUnderline(existingChangedSentence);
				}

				existingChangedSentence.PreviousUnderline = changedSentence.PreviousUnderline;

				if (changedSentence.Comment == null) SafeDeleteComment(existingChangedSentence.Comment);

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
				changedSentence.Comment = AddIncorrectRhythmComment(sentence, count);
			}

			AddOrUpdateChangedSentence(changedSentence);
		}

		private static Comment AddNeutralComment(Range sentence, int count) {
			var text = $"{count.ToString()} - {sentence.Text}";
			return AddComment(sentence, text);
		}

		private static Comment AddIncorrectRhythmComment(Range sentence, int count) {
			var text = $"⚠️ {count.ToString()} - {sentence.Text}";
			return AddComment(sentence, text);
		}

		private static Comment AddComment(Range sentence, string text) {
			var weirdComments = sentence.Comments.GetMadeByLightFeather().ToList();
			var startIndexes = weirdComments.Select(x => new { x.Scope.Start, x.Scope.Text });
			var currentStart = sentence.Start;
			var oldComments = sentence.Comments
				.GetMadeByLightFeather()
				.Where(x => x.Scope.Start == sentence.Start)
				.ToList();
			if (!oldComments.Any()) {
				object textObject = text;
				var comment = GetActiveDocument().Comments.Add(sentence, ref textObject);
				comment.ShowTip = true;
				comment.Author = CommentConsts.AuthorName;
				comment.Initial = "LF";
				return comment;
			}
			else {
				var previousComment = oldComments.First();
				previousComment.Range.Text = text;

				foreach (Comment comment in sentence.Comments) {
					var allSentencesInParagraph =
						sentence.Paragraphs.First.Range.Sentences.OfType<Range>().Select(x => x.Trim());
					if (allSentencesInParagraph.All(x => x.Start != comment.Scope.Start)) {
						SafeDeleteComment(comment);
					}
				}

				return previousComment;
			}
		}

		private static Document GetActiveDocument() {
			return Globals.ThisAddIn.Application.ActiveDocument;
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
				SafeCleanUnderline(changedSentence);
				SafeDeleteComment(changedSentence.Comment);
			}

			ChangedSentences.Clear();
		}

		private static void SafeDeleteComment(Comment comment) {
			if (comment is null)
				return;
			try {
				if (comment.Done) return;
				comment?.Delete();
			}
			catch (COMException e) {
				Debug.WriteLine($"Error: {e}");
			}
		}

		private static void SafeCleanUnderline(ChangedSentence changedSentence) {
			if (changedSentence.PreviousUnderline == null) return;

			try {
				if (changedSentence.Sentence.Underline == WdUnderline.wdUnderlineWavyHeavy) {
					changedSentence.Sentence.Underline = changedSentence.PreviousUnderline.Value;
				}
			}
			catch (COMException e) {
				Debug.WriteLine($"Error: {e}");
			}
		}

		private static void SafeCleanBackgroundColor(ChangedSentence changedSentence) {
			if (changedSentence.PreviousBackgroundColor == null) return;
			try {
				changedSentence.Sentence.Shading.BackgroundPatternColor = changedSentence.PreviousBackgroundColor.Value;
				if (changedSentence.Sentence.Underline == WdUnderline.wdUnderlineWavyHeavy) {
					changedSentence.Sentence.Underline = WdUnderline.wdUnderlineNone;
				}
			}
			catch (COMException e) {
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
					if (trimmedSentence.Shading.BackgroundPatternColor == GetRhythmCheckerBackgroundColor()) {
						trimmedSentence.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
					}

					if (trimmedSentence.Underline == WdUnderline.wdUnderlineWavyHeavy) {
						trimmedSentence.Underline = WdUnderline.wdUnderlineNone;
					}
				}
			}

			CleanAllLeftoverComments();
		}

		private static void CleanAllLeftoverComments() {
			foreach (var comment in Globals.ThisAddIn.Application.ActiveDocument.Comments.GetMadeByLightFeather()) {
				SafeDeleteComment(comment);
			}
		}
	}
}