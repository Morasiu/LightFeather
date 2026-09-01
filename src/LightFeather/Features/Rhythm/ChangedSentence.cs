using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Task = System.Threading.Tasks.Task;

namespace LightFeather.Features.Rhythm {
	public class ChangedSentence {
		public Range Sentence;
		public Comment Comment;
		public WdUnderline? PreviousUnderline { get; set; }

		public void SafeCleanUnderline() {
			if (this.PreviousUnderline == null)
				return;

			CleanUnderline();
		}

		private void CleanUnderline() {
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