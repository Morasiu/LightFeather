using System;
using System.Diagnostics;

namespace LightFeather.Shared {
	public class UndoAction : IDisposable {
		private readonly string _undoName;

		public UndoAction(string undoName = "Light Feather") {
			_undoName = undoName;
			Debug.WriteLine($"Undo started: {_undoName}", "Undo");
			Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord(_undoName);
		}

		public void Dispose() {
			var undoRecord = Globals.ThisAddIn.Application.UndoRecord;
			if (!undoRecord.IsRecordingCustomRecord)
				Debug.WriteLine($"No custom undo action in progress ({_undoName})", "Undo");

			undoRecord.EndCustomRecord();
			Debug.WriteLine($"Undo ended: {_undoName}", "Undo");
		}
	}
}