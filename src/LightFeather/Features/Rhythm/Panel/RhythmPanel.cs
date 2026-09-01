using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using LightFeather.Shared;
using Word = Microsoft.Office.Interop.Word;

namespace LightFeather.Features.Rhythm.Panel {
	public sealed class RhythmPanel : Form {
		private const int MaxSentenceLength = 400;

		private static readonly Color WarningBackColor = Color.FromArgb(255, 244, 222);
		private static readonly Color WarningForeColor = Color.FromArgb(140, 74, 0);

		private static Point? _lastLocation;
		private static Size? _lastSize;

		private readonly ListView _list;
		private readonly Label _summary;

		public RhythmPanel() {
			Text = "Light Feather - rytm akapitu";
			FormBorderStyle = FormBorderStyle.SizableToolWindow;
			ShowInTaskbar = false;
			StartPosition = FormStartPosition.Manual;
			MinimumSize = new Size(280, 180);
			ClientSize = _lastSize ?? new Size(420, 280);
			Font = new Font("Segoe UI", 9f);

			_list = new ListView {
				Dock = DockStyle.Fill,
				View = View.Details,
				FullRowSelect = true,
				HideSelection = false,
				MultiSelect = false,
				HeaderStyle = ColumnHeaderStyle.Nonclickable,
				ShowItemToolTips = true,
				UseCompatibleStateImageBehavior = false,
				BorderStyle = BorderStyle.None
			};

			_list.Columns.Add("#", 32, HorizontalAlignment.Right);
			_list.Columns.Add("Słowa", 48, HorizontalAlignment.Right);
			_list.Columns.Add("Zdanie", 320, HorizontalAlignment.Left);

			_list.DoubleClick += (sender, args) => SelectInDocument();
			_list.KeyDown += OnListKeyDown;

			_summary = new Label {
				Dock = DockStyle.Bottom,
				Height = 48,
				Padding = new Padding(8, 6, 8, 6),
				TextAlign = ContentAlignment.MiddleLeft,
				BackColor = SystemColors.Control
			};

			Controls.Add(_list);
			Controls.Add(_summary);

			SetSummary(0, 0);
		}

		protected override bool ShowWithoutActivation => true;


		public void AttachTo(IntPtr ownerHwnd) {
			if (ownerHwnd == IntPtr.Zero)
				return;

			WindowNative.SetWindowLong(Handle, WindowNative.GWL_HWNDPARENT, ownerHwnd);

			if (_lastLocation.HasValue) {
				Location = _lastLocation.Value;
				return;
			}

			MoveToDefaultCorner(ownerHwnd);
		}

		private void MoveToDefaultCorner(IntPtr ownerHwnd) {
			if (!WindowNative.GetWindowRect(ownerHwnd, out var owner))
				return;

			var margin = (int)Math.Round(24 * WindowNative.GetScale(ownerHwnd));

			Location = new Point(
				owner.Right - Width - margin,
				owner.Bottom - Height - margin);
		}

		public void SetSentences(IList<SentenceRhythm> sentences) {
			_list.BeginUpdate();

			try {
				while (_list.Items.Count > sentences.Count)
					_list.Items.RemoveAt(_list.Items.Count - 1);

				var warnings = 0;

				for (var index = 0; index < sentences.Count; index++) {
					var sentence = sentences[index];

					if (sentence.IncorrectRhythm)
						warnings++;

					var row = index < _list.Items.Count
						? _list.Items[index]
						: AddRow();

					Fill(row, index, sentence);
				}

				SetSummary(sentences.Count, warnings);
			}
			finally {
				_list.EndUpdate();
			}
		}

		private ListViewItem AddRow() {
			var row = new ListViewItem(string.Empty);

			row.SubItems.Add(string.Empty);
			row.SubItems.Add(string.Empty);

			_list.Items.Add(row);

			return row;
		}


		private static void Fill(ListViewItem row, int index, SentenceRhythm sentence) {
			var text = Flatten(sentence.Sentence);

			SetCell(row, 0, (index + 1).ToString());
			SetCell(row, 1, sentence.WordCount.ToString());
			SetCell(row, 2, text);

			var back = sentence.IncorrectRhythm ? WarningBackColor : SystemColors.Window;
			var fore = sentence.IncorrectRhythm ? WarningForeColor : SystemColors.WindowText;

			if (row.BackColor != back)
				row.BackColor = back;

			if (row.ForeColor != fore)
				row.ForeColor = fore;

			var toolTip = sentence.IncorrectRhythm
				? text + Environment.NewLine + Environment.NewLine +
				  "Podobna długość jak poprzednie zdanie."
				: text;

			if (row.ToolTipText != toolTip)
				row.ToolTipText = toolTip;

			var offsets = OffsetsOf(sentence.Sentence);

			if (offsets != null)
				row.Tag = offsets;
		}

		private static void SetCell(ListViewItem row, int column, string value) {
			if (row.SubItems[column].Text != value)
				row.SubItems[column].Text = value;
		}

		private void SetSummary(int sentenceCount, int warningCount) {
			var sentences = sentenceCount + " " + Plural(sentenceCount, "zdanie", "zdania", "zdań");

			if (warningCount == 0) {
				_summary.Text = sentences + ". Rytm w porządku.";
				_summary.ForeColor = SystemColors.GrayText;
				return;
			}

			_summary.Text = sentences + ". Ostrzeżenia: " + warningCount + Environment.NewLine +
			                "Kliknij wiersz dwukrotnie, aby przejść do zdania.";

			_summary.ForeColor = WarningForeColor;
		}

		private static string Plural(int count, string one, string few, string many) {
			if (count == 1)
				return one;

			var lastTwoDigits = count % 100;
			var lastDigit = count % 10;

			var isFew = lastDigit >= 2
			            && lastDigit <= 4
			            && (lastTwoDigits < 12 || lastTwoDigits > 14);

			return isFew ? few : many;
		}

		private static string Flatten(Word.Range sentence) {
			string text;

			try {
				text = sentence.Text;
			}
			catch (COMException) {
				return string.Empty;
			}

			if (string.IsNullOrEmpty(text))
				return string.Empty;

			var builder = new StringBuilder(text.Length);
			var lastWasSpace = false;

			foreach (var character in text) {
				var isSpace = char.IsWhiteSpace(character) || char.IsControl(character);

				if (isSpace) {
					if (!lastWasSpace && builder.Length > 0)
						builder.Append(' ');
				}
				else {
					builder.Append(character);
				}

				lastWasSpace = isSpace;

				if (builder.Length >= MaxSentenceLength)
					return builder.ToString().TrimEnd() + "...";
			}

			return builder.ToString().TrimEnd();
		}

		private static int[] OffsetsOf(Word.Range sentence) {
			try {
				return new[] { sentence.Start, sentence.End };
			}
			catch (COMException) {
				return null;
			}
		}

		private void OnListKeyDown(object sender, KeyEventArgs args) {
			if (args.KeyCode != Keys.Enter)
				return;

			args.Handled = true;
			SelectInDocument();
		}

		private void SelectInDocument() {
			if (_list.SelectedItems.Count == 0)
				return;

			var offsets = _list.SelectedItems[0].Tag as int[];

			if (offsets == null)
				return;

			try {
				var document = CurrentDocument.GetActiveDocument();

				if (offsets[1] > document.Content.End)
					return;

				document.Range(offsets[0], offsets[1]).Select();
				CurrentDocument.GetApplication().Activate();
			}
			catch (COMException e) {
				Debug.WriteLine($"Jump to sentence failed: {e.Message}", "[Rhythm]");
			}
		}

		protected override void OnResizeEnd(EventArgs args) {
			base.OnResizeEnd(args);
			Remember();
		}

		protected override void OnMove(EventArgs args) {
			base.OnMove(args);
			Remember();
		}

		protected override void OnFormClosing(FormClosingEventArgs args) {
			Remember();
			base.OnFormClosing(args);
		}

		private void Remember() {
			if (WindowState != FormWindowState.Normal || !IsHandleCreated)
				return;

			_lastLocation = Location;
			_lastSize = ClientSize;
		}
	}
}
