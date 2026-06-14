using LightFeather.Features.Rhythm;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;

namespace LightFeather.Extensions {
	public static class SentenceExtensions {
		public static IEnumerable<Range> GetActualWordsFromSentence(this Range sentence)
		{
			return sentence?.Words.Cast<Range>().Where(x => x.Text.Trim().Length > 0 && !char.IsPunctuation(x.Text.Trim()[0])).Select(x => x);
		}

		public static Range Trim(this Range sentence)
		{
			if (sentence == null) return null;
			if (sentence.Text == null) return sentence;
			var sentenceStart = sentence.Start;
			var sentenceEnd = sentence.End;

			if (sentence.Text.TrimEnd() != sentence.Text) {
				var sentenceLength = sentence.Text.TrimEnd().Length;
				sentenceEnd = sentenceStart + sentenceLength;
			}

			var activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
			var trimmedSentence = activeDocument.Range(sentenceStart, sentenceEnd);
			return trimmedSentence;
		}

		public static bool TextEqualTo(this Range sentence, Range otherSentence) {
			return sentence.Trim().Text == otherSentence.Trim().Text;
		}
	}
}