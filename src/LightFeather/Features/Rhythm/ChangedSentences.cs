using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using LightFeather.Extensions;
using LightFeather.Shared;

namespace LightFeather.Features.Rhythm {
	public class ChangedSentences {
		private readonly List<ChangedSentence> _changedSentenceList = new List<ChangedSentence>();

		public IReadOnlyCollection<ChangedSentence> Log => _changedSentenceList;

		public void AddOrUpdate(ChangedSentence changedSentence) {
			var existingChangedSentence =
				_changedSentenceList.FirstOrDefault(x => x.Sentence.TextEqualTo(changedSentence.Sentence));
			if (existingChangedSentence == null) {
				_changedSentenceList.Add(changedSentence);
			}
			else {
				if (changedSentence.PreviousBackgroundColor == null)
					changedSentence.SafeCleanBackgroundColor();
				existingChangedSentence.PreviousBackgroundColor = changedSentence.PreviousBackgroundColor;

				if (changedSentence.PreviousUnderline == null)
					existingChangedSentence.SafeCleanUnderline();
				existingChangedSentence.PreviousUnderline = changedSentence.PreviousUnderline;

				if (changedSentence.Comment == null)
					existingChangedSentence.Comment.SafeDelete();
				existingChangedSentence.Comment = changedSentence.Comment;
			}
		}


		public void CleanupChangedSentences() {
			CleanChangedSentencesInternal(_changedSentenceList);
		}

		private void CleanChangedSentencesInternal(List<ChangedSentence> sentences) {
			Debug.WriteLine("Cleanup started.", "[Rhythm]");
			var stopWatch = Stopwatch.StartNew();

			var undoAction = new UndoAction("Light feather - cleanup sentences");

			Parallel.ForEach(sentences, CleanSentence);

			undoAction.Dispose();

			stopWatch.Stop();
			Debug.WriteLine($"Cleanup took: {stopWatch.ElapsedMilliseconds}ms.", "[Rhythm]");
		}

		private static void CleanSentence(ChangedSentence changedSentence) {
			changedSentence.SafeCleanBackgroundColor();
			changedSentence.SafeCleanUnderline();
			changedSentence.Comment.SafeDelete();
		}
	}
}