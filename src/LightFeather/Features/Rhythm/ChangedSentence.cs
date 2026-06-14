using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace LightFeather.Features.Rhythm {
	public class ChangedSentence {
		public Range Sentence;
		public WdColor? PreviousBackgroundColor;
		public Comment Comment;
		public WdUnderline? PreviousUnderline { get; set; }

		public void SafeCleanBackgroundColor() {
			if (this.PreviousBackgroundColor == null)
				return;

			try {
				this.Sentence.Shading.BackgroundPatternColor = this.PreviousBackgroundColor.Value;
				if (this.Sentence.Underline == WdUnderline.wdUnderlineWavyHeavy) {
					this.Sentence.Underline = WdUnderline.wdUnderlineNone;
				}
			}
			catch (COMException e) {
				Debug.WriteLine($"Error: {e}");
			}
		}

		public void SafeCleanUnderline() {
			if (this.PreviousUnderline == null)
				return;

			try {
				if (this.Sentence.Underline == WdUnderline.wdUnderlineWavyHeavy) {
					this.Sentence.Underline = this.PreviousUnderline.Value;
				}
			}
			catch (COMException e) {
				Debug.WriteLine($"Error: {e}");
			}
		}
	}
}