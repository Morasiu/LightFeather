using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;

namespace LightFeather.Extensions {
	public static class SentenceExtensions {
		public static IEnumerable<Range> GetActualWordsFromSentence(this Range sentence)
		{
			return sentence.Words.Cast<Range>().Where(x => x.Text.Trim().Length > 0 && !char.IsPunctuation(x.Text.Trim()[0])).Select(x => x);
		}
	}
}