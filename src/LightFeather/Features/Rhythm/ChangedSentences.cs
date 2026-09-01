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
			var existingChangedSentence = _changedSentenceList
				.FirstOrDefault(x => x.Sentence.TextEqualTo(changedSentence.Sentence));
			if (existingChangedSentence == null) {
				_changedSentenceList.Add(changedSentence);
			}
			else {
				UpdateSentence(existingChangedSentence, changedSentence);
			}
		}

		private static void UpdateSentence(ChangedSentence currentSentence, ChangedSentence updatedSentence) {
			if (updatedSentence.PreviousUnderline == null) {
				currentSentence.SafeCleanUnderline();
			}
			currentSentence.PreviousUnderline = updatedSentence.PreviousUnderline;

			if (updatedSentence.Comment == null) {
				currentSentence.Comment.SafeDelete();
			}
			currentSentence.Comment = updatedSentence.Comment;
		}


		public void CleanupChangedSentences() {
			CleanChangedSentencesInternal(_changedSentenceList);
			_changedSentenceList.Clear();
		}

		private static void CleanChangedSentencesInternal(List<ChangedSentence> sentences) {
			Debug.WriteLine("Cleanup started.", "[Rhythm]");
			var stopWatch = Stopwatch.StartNew();

			var undoAction = new UndoAction("Light feather - cleanup sentences");

			Parallel.ForEach(sentences, CleanSentence);
			//sentences.ForEach(CleanSentence);

			undoAction.Dispose();

			stopWatch.Stop();
			Debug.WriteLine($"Cleanup took: {stopWatch.ElapsedMilliseconds}ms.", "[Rhythm]");
		}

		private static void CleanSentence(ChangedSentence changedSentence) {
			var stopWatch = Stopwatch.StartNew();
			changedSentence.SafeCleanUnderline();
			Debug.WriteLine($"Underline: {stopWatch.ElapsedMilliseconds}ms");
			stopWatch.Restart();
			changedSentence.Comment.SafeDelete();
			Debug.WriteLine($"Comment: {stopWatch.ElapsedMilliseconds}ms");
		}
	}
}