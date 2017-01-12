using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReplaceText
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();

			dataGridView1.Rows.Add(new object[] { "統一商品コード", "高山商品コード" });
			dataGridView1.Rows.Add(new object[] { "商品統一コード", "高山商品コード" });
			dataGridView1.Rows.Add(new object[] { "JAN", "ＪＡＮコード" });
			dataGridView1.Rows.Add(new object[] { "ＪＡＮ", "ＪＡＮコード" });
			dataGridView1.Rows.Add(new object[] { "メーカー", "仕入先" });
		}

		XLWorkbook workbook = null;
		IXLWorksheet ws = null;

		private void btnReplace_Click(object sender, EventArgs e)
		{
			if (ws == null)
			{
				return;
			}

			var selectedItems = listBox1.SelectedItems;
			List<MatchValue> toBeRemoved = new List<MatchValue>();
			IXLRange printArea = ws.PageSetup.PrintAreas.FirstOrDefault();
			int colCount = 0;
			if (printArea != null)
			{
				colCount = printArea.ColumnCount();
			}

			foreach (MatchValue item in selectedItems)
			{
				ws.Cell(item.CellPosition).Value = (ws.Cell(item.CellPosition).Value as string)
					.Remove(item.StartIndex, item.ToBeReplacedString.Length)
					.Insert(item.StartIndex, item.ReplacementString);

				int offset = item.ReplacementString.Length - item.ToBeReplacedString.Length;
				foreach (MatchValue itemBehind in selectedItems.Cast<MatchValue>()
					.Where(x => x.CellPosition == item.CellPosition && x.StartIndex > item.StartIndex))
				{
					itemBehind.StartIndex += offset;
				}

				if (colCount != 0)
				{
					ws.Cell(item.CellPosition.RowNumber, colCount + 1).Value = DateTime.Now.ToString("yyyy/MM/dd") + " 変更";
					ws.Cell(item.CellPosition.RowNumber, colCount + 1).Style.Font.FontColor = XLColor.Red;
				}

				toBeRemoved.Add(item);
			}

			foreach (MatchValue item in selectedItems)
			{
				ws.Cell(item.CellPosition).RichText.Substring(item.StartIndex, item.ReplacementString.Length).FontColor = XLColor.Red;
			}

			foreach (var item in toBeRemoved)
			{
				listBox1.Items.Remove(item);
			}

			foreach (var worksheet in workbook.Worksheets)
			{
				worksheet.SheetView.View = XLSheetViewOptions.PageBreakPreview;
				worksheet.Cell(1, 1).SetActive();
			}

			if (checkBox1.Checked)
			{
				workbook.Save();
			}
			else
			{
				string saveAsName = Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName) + "_Edit" + Path.GetExtension(fileName));
				workbook.SaveAs(saveAsName);
			}
		}

		private static string fileName = null;

		class MatchValue
		{
			public IXLAddress CellPosition { get; set; }
			public int MatchIndex { get; set; }
			public int StartIndex { get; set; }
			public string CellValue { get; set; }
			public string ToBeReplacedString { get; set; }
			public string ReplacementString { get; set; }
		}

		private void btnBrowse_Click(object sender, EventArgs e)
		{
			OpenFileDialog saveFileDialog = new OpenFileDialog();
			saveFileDialog.Filter = "Excel|*.xlsx";
			saveFileDialog.Title = "Choose a location";
			saveFileDialog.FileName = Path.GetFileName(fileName);

			if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				btnClose_Click(sender, e);
				fileName = saveFileDialog.FileName;
				using (workbook = new XLWorkbook(fileName, XLEventTracking.Disabled))
				{
					foreach (var worksheet in workbook.Worksheets)
					{
						cmbSheets.Items.Add(worksheet.Name);
					}

					if (cmbSheets.Items.Contains("3章 処理説明書"))
					{
						cmbSheets.SelectedItem = "3章 処理説明書";
					}
					else if (cmbSheets.Items.Count > 0)
					{
						cmbSheets.SelectedIndex = 0;
					}
				}
			}
		}

		private void SetActiveWorksheet(string worksheetName)
		{
			ws = workbook.Worksheet(worksheetName);
			if (ws != null)
			{
				listBox1.Items.Clear();
				var valueList = dataGridView1.Rows.Cast<DataGridViewRow>().Select(r => r.Cells[0] == null ? string.Empty : r.Cells[0].Value as string).Where(v => v != null).ToList();
				var replaceList = dataGridView1.Rows.Cast<DataGridViewRow>().Select(r => r.Cells[1] == null ? string.Empty : r.Cells[1].Value as string).Where(v => v != null).ToList();

				var cells = ws.CellsUsed(x => CellContainsList(x, valueList) && !x.Style.Font.Strikethrough);
				foreach (var cell in cells)
				{
					foreach (var compareToValue in valueList)
					{
						if (CellContains(cell, compareToValue))
						{
							//var indexes = AllIndexesOf(cell.Value as string, compareToValue);
							//foreach (var index in indexes)
							//{
							//	MatchValue itemValue = new MatchValue()
							//	{
							//		CellPosition = cell.Address,
							//		MatchIndex = index,
							//		CellValue = cell.Value as string,
							//		ToBeReplacedString = compareToValue,
							//		ReplacementString = replaceList[valueList.IndexOf(compareToValue)],
							//	};

							//	listBox1.Items.Add(itemValue);
							//}

							var indexes = AllIndexesOf(cell.Value as string, compareToValue);

							for (int i = 0; i < indexes.Count; i++)
							{
								MatchValue itemValue = new MatchValue()
								{
									CellPosition = cell.Address,
									MatchIndex = i,
									StartIndex = indexes[i],
									CellValue = cell.Value as string,
									ToBeReplacedString = compareToValue,
									ReplacementString = replaceList[valueList.IndexOf(compareToValue)],
								};

								listBox1.Items.Add(itemValue);
							}
						}
					}
				}
			}
		}

		private static bool CellContainsList(IXLCell value, IEnumerable<string> compareToList)
		{
			foreach (var compareToValue in compareToList)
			{
				if (CellContains(value, compareToValue))
				{
					return true;
				}
			}

			return false;
		}

		private static bool CellContains(IXLCell value, string compareToValue)
		{
			if (value == null || string.IsNullOrEmpty(compareToValue))
			{
				return false;
			}

			if (value.HasFormula)
			{
				return false;
			}

			string stringValue = null;
			try
			{
				stringValue = value.Value as string;
			}
			catch
			{
				// Formula error
				return false;
			}

			if (stringValue == null)
			{
				return false;
			}

			if (stringValue.IndexOf(compareToValue, StringComparison.OrdinalIgnoreCase) != -1)
			{
				return true;
			}

			return value.Equals(compareToValue);
		}

		public static List<int> AllIndexesOf(string str, string value)
		{
			if (String.IsNullOrEmpty(value))
				throw new ArgumentException("the string to find may not be empty", "value");
			List<int> indexes = new List<int>();
			for (int index = 0; ; index += value.Length)
			{
				index = str.IndexOf(value, index);
				if (index == -1)
					return indexes;
				indexes.Add(index);
			}
		}

		private void listBox1_DrawItem(object sender, DrawItemEventArgs e)
		{
			Brush brush = ((e.State & DrawItemState.Selected) == DrawItemState.Selected) ?
				  Brushes.LightBlue : new SolidBrush(e.BackColor);
			e.Graphics.FillRectangle(brush, e.Bounds);

			if (e.Index == -1) return;

			MatchValue value = (MatchValue)listBox1.Items[e.Index];
			int length = 0;
			Point pt = new Point(e.Bounds.X, e.Bounds.Y);

			var normalHeight = TextRenderer.MeasureText("Y", Font).Height;

			string location = string.Format("({0}{1})", value.CellPosition.ColumnLetter, value.CellPosition.RowNumber);
			TextRenderer.DrawText(e.Graphics, location, Font, pt, Color.Black);
			pt.X += TextRenderer.MeasureText(location, Font).Width - 2;

			if (value.StartIndex != 0)
			{
				string beforeString = new string(value.CellValue.Take(value.StartIndex).ToArray());
				TextRenderer.DrawText(e.Graphics, beforeString, Font, pt, Color.Black);

				// Last line width
				string beforeStringLastLine = beforeString.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None).LastOrDefault();
				pt.X += TextRenderer.MeasureText(beforeStringLastLine, Font).Width - 2;

				pt.Y += TextRenderer.MeasureText(beforeString, Font).Height - normalHeight;
				length += beforeString.Length;
			}

			var boldFont = new Font(Font.FontFamily, Font.Size, FontStyle.Bold);
			TextRenderer.DrawText(e.Graphics, value.ToBeReplacedString, boldFont, pt, Color.Black);
			pt.X += TextRenderer.MeasureText(value.ToBeReplacedString, Font).Width + 3;
			pt.Y += TextRenderer.MeasureText(value.ToBeReplacedString, Font).Height - normalHeight;

			length += value.ToBeReplacedString.Length;
			if (length < value.CellValue.Length)
			{
				string afterString = new string(value.CellValue.Skip(length).ToArray());
				TextRenderer.DrawText(e.Graphics, afterString, Font, pt, Color.Black);
			}

			e.DrawFocusRectangle();
		}

		private void listBox1_MeasureItem(object sender, MeasureItemEventArgs e)
		{
			if (e.Index == -1) return;
			MatchValue value = (MatchValue)listBox1.Items[e.Index];

			int count = value.CellValue.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None).Length;
			if (count > 1)
				e.ItemHeight = e.ItemHeight * count;
		}

		private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Control && e.KeyCode == Keys.V)
			{
				DataObject o = (DataObject)Clipboard.GetDataObject();
				if (o.GetDataPresent(DataFormats.Text))
				{
					string[] pastedRows = Regex.Split(o.GetData(DataFormats.Text).ToString().TrimEnd("\r\n".ToCharArray()), "\r\n");
					foreach (string pastedRow in pastedRows)
					{
						string[] pastedRowCells = pastedRow.Split(new char[] { '\t' }, StringSplitOptions.RemoveEmptyEntries);
						dataGridView1.Rows.Add(pastedRowCells);
					}
				}

				var selectedSheet = cmbSheets.SelectedItem as string;
				if (string.IsNullOrEmpty(selectedSheet))
				{
					return;
				}

				SetActiveWorksheet(selectedSheet);
			}
		}

		private void cmbSheets_SelectedIndexChanged(object sender, EventArgs e)
		{
			SetActiveWorksheet(cmbSheets.SelectedItem as string);
		}

		private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
		{
			if (e.RowIndex < 0)
			{
				return;
			}

			if (dataGridView1.Rows[e.RowIndex].Cells[0].Value == null || dataGridView1.Rows[e.RowIndex].Cells[1].Value == null)
			{
				return;
			}

			var selectedSheet = cmbSheets.SelectedItem as string;
			if (string.IsNullOrEmpty(selectedSheet))
			{
				return;
			}

			SetActiveWorksheet(selectedSheet);
		}

		private void btnClose_Click(object sender, EventArgs e)
		{
			ws = null;
			workbook = null;
			cmbSheets.Items.Clear();
			listBox1.Items.Clear();
		}

		private void checkBox1_CheckedChanged(object sender, EventArgs e)
		{

		}
	}
}
