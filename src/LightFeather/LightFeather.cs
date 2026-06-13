using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace LightFeather
{
	public partial class LightFeather
	{
		private void LightFeather_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void rythmCheckbox_Click(object sender, RibbonControlEventArgs e)
		{
			if (rhythmCheckbox.Checked)
			{
				RhythmChecker.CheckRhythm();
			}
			else
			{
				RhythmChecker.DisableCheckRhtyhm();
			}
		}

		private void About_Click(object sender, RibbonControlEventArgs e)
		{
			string text = "O dodatku \n" +
						  "LightFeather \n " +
						  $"Wersja: {typeof(LightFeather).Assembly.GetName().Version}";
			MessageBox.Show(text, "O Light Feather");
		}
	}
}
