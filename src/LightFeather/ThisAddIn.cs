using LightFeather.Features.Rhythm;

namespace LightFeather
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

		}


		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            RhythmChecker.DisableCheckRhythm();
        }

        #region Kod wygenerowany przez program VSTO

        /// <summary>
        /// Metoda wymagana do obsługi projektanta — nie należy modyfikować
        /// jej zawartości w edytorze kodu.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
