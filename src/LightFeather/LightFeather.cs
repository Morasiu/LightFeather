using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using LightFeather.Features.Rhythm;

namespace LightFeather {
	public partial class LightFeather {
		private void LightFeather_Load(object sender, RibbonUIEventArgs e) { }

		private void rythmCheckbox_Click(object sender, RibbonControlEventArgs e) {
			if (rhythmCheckbox.Checked) {
				RhythmChecker.CheckRhythm();
			}
			else {
				RhythmChecker.DisableCheckRhythm();
			}
		}

		private void About_Click(object sender, RibbonControlEventArgs e) {
			string text = "O dodatku \n" +
			              "LightFeather \n " +
			              $"Wersja: {typeof(LightFeather).Assembly.GetName().Version}";
			MessageBox.Show(text, "O Light Feather");
		}

		private void useCommentCheckbox_Click(object sender, RibbonControlEventArgs e) {
			RhythmChecker.DisableCheckRhythm();
			RhythmChecker.UseComments = useCommentCheckbox.Checked;
			if (rhythmCheckbox.Checked) RhythmChecker.CheckRhythm();
		}

		private void checkBox1_Click(object sender, RibbonControlEventArgs e) {
			RhythmChecker.DisableCheckRhythm();
			RhythmChecker.UseBackgroundChange = useBackgroundChangeCheckbox.Checked;
			if (rhythmCheckbox.Checked) RhythmChecker.CheckRhythm();
		}

		private void cleanRhythmChecks_Click(object sender, RibbonControlEventArgs e) {
			RhythmChecker.CleanAllLeftovers();
		}
	}
}