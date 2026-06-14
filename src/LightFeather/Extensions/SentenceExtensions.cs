using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;

namespace LightFeather.Extensions {
	public static class SentenceExtensions {
		public static IEnumerable<Range> GetActualWordsFromSentence(this Range sentence)
		{
			return sentence.Words.Cast<Range>().Where(x => x.Text.Trim().Length > 0 && !char.IsPunctuation(x.Text.Trim()[0])).Select(x => x);
		}

		public static Range Trim(this Range sentence)
		{
			var sentenceToEdit = sentence;
			if (sentence.Text.EndsWith("\r"))
			{
				var activeDocument = Globals.ThisAddIn.Application.ActiveDocument;
				var trimmedRange = activeDocument.Range(sentence.Start, sentence.End - 1);
				sentenceToEdit = trimmedRange;
			}
			return sentenceToEdit;
		}
	}
}