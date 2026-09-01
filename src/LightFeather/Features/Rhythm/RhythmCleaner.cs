using LightFeather.Extensions;
using LightFeather.Shared;
using Microsoft.Office.Interop.Word;

namespace LightFeather.Features.Rhythm {
	public class RhythmCleaner {
		public static void CleanAll() {
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

				CleanLeftoverUnderlineManipulation(trimmedSentence);
			}
		}

		private static void CleanLeftoverUnderlineManipulation(Range trimmedSentence) {
			if (trimmedSentence.Underline != WdUnderline.wdUnderlineWavyHeavy)
				return;

			trimmedSentence.Underline = WdUnderline.wdUnderlineNone;
		}

		private static void CleanAllLeftoverComments() {
			foreach (var comment in CurrentDocument.GetActiveDocument().Comments.FilterMadeByLightFeather()) {
				comment.SafeDelete();
			}
		}
	}
}