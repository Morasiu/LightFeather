using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using LightFeather.Features.Rhythm;

namespace LightFeather {
	public partial class LightFeather {
		private void LightFeather_Load(object sender, RibbonUIEventArgs e) {
			if (rhythmCheckbox.Checked) RhythmChecker.CheckRhythm();
		}

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

		private void showPanelCheckbox_Click(object sender, RibbonControlEventArgs e) {
			RhythmChecker.UsePanel = showPanelCheckbox.Checked;

			if (!showPanelCheckbox.Checked) {
				RhythmChecker.ClosePanel();
				return;
			}

			if (rhythmCheckbox.Checked) RhythmChecker.OpenPanel();
		}
	}
}