using Microsoft.Office.Interop.Word;

namespace LightFeather.Features.Rhythm {
	public class ChangedSentence
	{
		public Range Sentence;
		public WdColor? PreviousBackgroundColor;
		public Comment Comment;
	}
}