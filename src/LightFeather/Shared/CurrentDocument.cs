using Microsoft.Office.Interop.Word;

namespace LightFeather.Shared {
	public class CurrentDocument {
		public static Document GetActiveDocument() => Globals.ThisAddIn.Application.ActiveDocument;

		public static Selection GetCurrentSelection() => Globals.ThisAddIn.Application.Selection;

	}
}