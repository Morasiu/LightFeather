using Word = Microsoft.Office.Interop.Word;

namespace LightFeather.Features.Rhythm {
	public sealed class SentenceRhythm {
		public SentenceRhythm(Word.Range sentence, int wordCount, bool incorrectRhythm) {
			Sentence = sentence;
			WordCount = wordCount;
			IncorrectRhythm = incorrectRhythm;
		}

		public Word.Range Sentence { get; }
		public int WordCount { get; }
		public bool IncorrectRhythm { get; }
	}
}
