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


		}

		IXLWorksheet ws = null;
		const string findString = "統一商品コード";
		const string replaceWithString = "高山商品コード";

		private void btnReplace_Click(object sender, EventArgs e)
		{
			if (ws == null)
			{
				return;
			}

			var selectedItems = listBox1.SelectedItems;
			List<MatchValue> toBeRemoved = new List<MatchValue>();
			IXLRange printArea = ws.PageSetup.PrintAreas.FirstOrDefault();
			int colCount = printArea.ColumnCount();
			foreach (MatchValue item in selectedItems)
			{
				string newValue = item.CellValue
					.Remove(item.MatchIndex, item.ReplaceString.Length)
					.Insert(item.MatchIndex, replaceWithString);
				ws.Cell(item.CellPosition).Value = newValue;
				ws.Cell(item.CellPosition).RichText.Substring(item.MatchIndex, replaceWithString.Length).FontColor = XLColor.Red;

				ws.Cell(item.CellPosition.RowNumber, colCount + 1).Value = DateTime.Now.ToString("yyyy/MM/dd") + " 変更";
				ws.Cell(item.CellPosition.RowNumber, colCount + 1).Style.Font.FontColor = XLColor.Red;
				toBeRemoved.Add(item);
			}

			foreach (var item in toBeRemoved)
			{
				listBox1.Items.Remove(item);
			}

			string saveAsName = Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName) + "_Edit" + Path.GetExtension(fileName));
			ws.Workbook.SaveAs(saveAsName);
		}

		private static string fileName = null;

		class MatchValue
		{
			public IXLAddress CellPosition { get; set; }
			public int MatchIndex { get; set; }
			public string CellValue { get; set; }
			public string ReplaceString { get; set; }
		}

		private void btnBrowse_Click(object sender, EventArgs e)
		{
			OpenFileDialog saveFileDialog = new OpenFileDialog();
			saveFileDialog.Filter = "Excel|*.xlsx";
			saveFileDialog.Title = "Choose a location";
			saveFileDialog.FileName = Path.GetFileName(fileName);

			if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
			{
				fileName = saveFileDialog.FileName;
				using (var workbook = new XLWorkbook(saveFileDialog.OpenFile(), XLEventTracking.Disabled))
				{
					ws = workbook.Worksheet("3章 処理説明書");
					if (ws != null)
					{
						var cells = ws.CellsUsed(x => CompareValueTo(x, findString));
						foreach (var cell in cells)
						{
							var indexes = AllIndexesOf(cell.Value as string, findString);
							foreach (var index in indexes)
							{
								MatchValue itemValue = new MatchValue()
								{
									CellPosition = cell.Address,
									MatchIndex = index,
									CellValue = cell.Value as string,
									ReplaceString = findString,
								};

								listBox1.Items.Add(itemValue);
							}
						}
					}
				}
			}
		}

		private static bool CompareValueTo(IXLCell value, string compareToValue)
		{
			if (value == null || compareToValue == null)
			{
				return false;
			}

			string stringValue = value.Value as string;
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
			e.DrawBackground();

			MatchValue value = (MatchValue)listBox1.Items[e.Index];
			int length = 0;
			Point pt = new Point(e.Bounds.X, e.Bounds.Y);
			if (value.MatchIndex != 0)
			{
				string beforeString = new string(value.CellValue.Take(value.MatchIndex).ToArray());
				TextRenderer.DrawText(e.Graphics, beforeString, Font, pt, Color.Black);
				pt.X += TextRenderer.MeasureText(beforeString, Font).Width - 5;
				length += beforeString.Length;
			}

			var boldFont = new Font(Font.FontFamily, Font.Size, FontStyle.Bold);
			TextRenderer.DrawText(e.Graphics, value.ReplaceString, boldFont, pt, Color.Black);
			pt.X += TextRenderer.MeasureText(value.ReplaceString, Font).Width + 3;

			length += value.ReplaceString.Length;
			if (length < value.CellValue.Length)
			{
				string afterString = new string(value.CellValue.Skip(length).ToArray());
				TextRenderer.DrawText(e.Graphics, afterString, Font, pt, Color.Black);
			}

			e.DrawFocusRectangle();
		}

		private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.Control && e.KeyCode == Keys.V)
			{
				DataObject o = (DataObject)Clipboard.GetDataObject();
				if (o.GetDataPresent(DataFormats.Text))
				{
					if (dataGridView1.RowCount > 0)
						dataGridView1.Rows.Clear();

					if (dataGridView1.ColumnCount > 0)
						dataGridView1.Columns.Clear();

					bool columnsAdded = false;
					string[] pastedRows = Regex.Split(o.GetData(DataFormats.Text).ToString().TrimEnd("\r\n".ToCharArray()), "\r\n");
					foreach (string pastedRow in pastedRows)
					{
						string[] pastedRowCells = pastedRow.Split(new char[] { '\t' });

						if (!columnsAdded)
						{
							for (int i = 0; i < pastedRowCells.Length; i++)
								dataGridView1.Columns.Add("col" + i, pastedRowCells[i]);

							columnsAdded = true;
							continue;
						}

						dataGridView1.Rows.Add();
						int myRowIndex = dataGridView1.Rows.Count - 1;

						using (DataGridViewRow dataGridView1Row = dataGridView1.Rows[myRowIndex])
						{
							for (int i = 0; i < pastedRowCells.Length; i++)
								dataGridView1Row.Cells[i].Value = pastedRowCells[i];
						}
					}
				}
			}
		}
	}
}
