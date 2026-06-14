using System.Collections.Generic;
using System.Linq;
using LightFeather.Shared;
using Microsoft.Office.Interop.Word;

namespace LightFeather.Extensions {
	public static class CommentsExtensions {
		public static IEnumerable<Comment> GetMadeByLightFeather(this Comments comments) {
			return comments
				.OfType<Comment>()
				.Where(comment => comment.Author == CommentConsts.AuthorName);
		}
	}
}