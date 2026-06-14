using LightFeather.Shared;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;
using LightFeather.Extensions;

namespace LightFeather.Features.Rhythm {
	public class CommentFactory {
		public static Comment AddNeutralComment(Range sentence, int count) {
			var text = $"{count} - {sentence.Text}";
			return AddComment(sentence, text);
		}

		public static Comment AddIncorrectRhythmComment(Range sentence, int count) {
			var text = $"⚠️ {count} - {sentence.Text}";
			return AddComment(sentence, text);
		}

		private static Comment AddComment(Range sentence, string text) {
			var oldComments = sentence.Comments
				.GetMadeByLightFeather()
				.Where(x => x.Scope.Start == sentence.Start)
				.ToList();
			if (!oldComments.Any()) {
				return AddNewComment(text, sentence);
			}
			else {
				return UpdateExistingComments(sentence, text, oldComments);
			}
		}

		private static Comment UpdateExistingComments(Range sentence, string text, List<Comment> oldComments) {
			var previousComment = oldComments.First();
			if (previousComment.Range.Text != text)
				previousComment.Range.Text = text;

			RemoveOrphanComments(sentence);

			return previousComment;
		}

		private static Comment AddNewComment(string text, Range sentence) {
			object textObject = text;
			var comment = CurrentDocument.GetActiveDocument().Comments.Add(sentence, ref textObject);
			comment.ShowTip = true;
			comment.Author = CommentConsts.AuthorName;
			comment.Initial = "LF";
			return comment;
		}

		private static void RemoveOrphanComments(Range sentence) {
			var allSentencesInParagraph = sentence.Paragraphs.First.Range.Sentences
				.OfType<Range>()
				.Select(x => x.Trim())
				.ToList();
			foreach (Comment comment in sentence.Comments) {
				if (allSentencesInParagraph.All(x => x.Start != comment.Scope.Start)) {
					comment.SafeDelete();
				}
			}
		}
	}
}