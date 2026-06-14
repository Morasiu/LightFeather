using LightFeather.Shared;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace LightFeather.Extensions {
	public static class CommentsExtensions {
		public static IEnumerable<Comment> GetMadeByLightFeather(this Comments comments) {
			return comments
				.OfType<Comment>()
				.Where(comment => comment.Author == CommentConsts.AuthorName);
		}

		public static void SafeDelete(this Comment comment) {
			if (comment is null)
				return;
			try {
				if (comment.Done)
					return;
				comment?.Delete();
			}
			catch (COMException e) {
				Debug.WriteLine($"Error: {e}");
			}
		}
	}
}