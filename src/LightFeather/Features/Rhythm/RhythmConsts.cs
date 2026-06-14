using Microsoft.Office.Interop.Word;
using System.Drawing;

namespace LightFeather.Features.Rhythm {
	public class RhythmConsts {
		public static readonly WdColor IncorrectRhythmBackgroundColor = (WdColor)ColorTranslator.ToOle(Color.FromArgb(128, 255, 255, 0));
	}
}