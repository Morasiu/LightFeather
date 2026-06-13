namespace LightFeather
{
	partial class LightFeather : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Wymagana zmienna projektanta.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public LightFeather()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Wyczyść wszystkie używane zasoby.
		/// </summary>
		/// <param name="disposing">prawda, jeżeli zarządzane zasoby powinny zostać zlikwidowane; Fałsz w przeciwnym wypadku.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Kod wygenerowany przez Projektanta składników

		/// <summary>
		/// Metoda wymagana do obsługi projektanta — nie należy modyfikować
		/// jej zawartości w edytorze kodu.
		/// </summary>
		private void InitializeComponent()
		{
			this.tab1 = this.Factory.CreateRibbonTab();
			this.RhythmGroup = this.Factory.CreateRibbonGroup();
			this.rhythmCheckbox = this.Factory.CreateRibbonCheckBox();
			this.About = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.RhythmGroup.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Groups.Add(this.RhythmGroup);
			this.tab1.Label = "Light Feather";
			this.tab1.Name = "tab1";
			// 
			// RhythmGroup
			// 
			this.RhythmGroup.Items.Add(this.rhythmCheckbox);
			this.RhythmGroup.Items.Add(this.About);
			this.RhythmGroup.Label = "Rytm";
			this.RhythmGroup.Name = "RhythmGroup";
			// 
			// rhythmCheckbox
			// 
			this.rhythmCheckbox.Label = "Włącz sprawdzenie rytmu";
			this.rhythmCheckbox.Name = "rhythmCheckbox";
			this.rhythmCheckbox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rythmCheckbox_Click);
			// 
			// About
			// 
			this.About.Label = "O LightFeather";
			this.About.Name = "About";
			this.About.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.About_Click);
			// 
			// LightFeather
			// 
			this.Name = "LightFeather";
			this.RibbonType = "Microsoft.Word.Document";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.LightFeather_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.RhythmGroup.ResumeLayout(false);
			this.RhythmGroup.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup RhythmGroup;
		internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox rhythmCheckbox;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton About;
	}

	partial class ThisRibbonCollection
	{
		internal LightFeather LightFeather
		{
			get { return this.GetRibbon<LightFeather>(); }
		}
	}
}
