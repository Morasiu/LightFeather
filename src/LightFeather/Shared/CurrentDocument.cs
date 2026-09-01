using Microsoft.Office.Interop.Word;

namespace LightFeather.Shared {
	public class CurrentDocument {
		public static Application GetApplication() => Globals.ThisAddIn.Application;

		public static Document GetActiveDocument() => Globals.ThisAddIn.Application.ActiveDocument;

		public static Selection GetCurrentSelection() => Globals.ThisAddIn.Application.Selection;

	}
}