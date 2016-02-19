/*
	© 2016 Florida State University. All rights reserved.

	TODO:
			- Change this to support multiple file formats.

	History
	==============================================================================================
	2016/02/03	G.K.	Created.

*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.IO;
using System.Xml.Linq;

namespace Read_XLSX
{
	class DataCellValue
	{
		public Field field { get; set; }
		public string CellReference { get; set; }
		public string Value { get; set; }
		public int rowNumber { get; set; }
		public DataRow Row { get; set; }
	}

	class DataRow
	{
		public int Row { get; set; }
		public Dictionary<int, DataCellValue> Cells { get; set; }
		public DataSet Sheet { get; set; }
		public DataRow(int row)
		{
			Row = row;
		}

		public DataCellValue AddDataCell(DataCellValue dc)
		{
			if (Cells == null) Cells = new Dictionary<int, DataCellValue>();

			if (Cells.Where(c => c.Value.CellReference == dc.CellReference).Count() > 0)
				throw new Exception("Cell has already been recorded");
			dc.Row = this;
			Cells.Add(dc.field.OutputOrder, dc);

			return dc;
		}

		public bool HasRequiredCol(List<int> reqCol)
		{
			var cols = Cells.Select(c => c.Key).ToList();
			var missing = reqCol.Where(r => !cols.Contains(r));
			return missing.Count() == 0;
		}

		public StringBuilder DelimitedRow(StringBuilder sb, List<Field> datacols, string delim)
		{
			int lastCol = datacols.Max(d => d.OutputOrder);
			foreach (var v in datacols.OrderBy(d => d.OutputOrder))
			{
				try
				{
					string val = null;

					switch (v.fldType)
					{
						case FieldType.column:
							val = Cells.ContainsKey(v.OutputOrder) ? Cells[v.OutputOrder].Value ?? "" : "";
							break;
						case FieldType.cell:
						case FieldType.fileName:
							var df = this.Sheet.wsLayout.fieldCellMap.fldmaps.FirstOrDefault(fm => fm.field.OutputOrder == v.OutputOrder);
							val = (df != null ? df.Value ?? "" : "");
							break;
					}

					// Don't want a delimiter after last column.
					sb.Append(val + (v.OutputOrder < lastCol ? delim : ""));
				}
				catch(Exception ex)
				{
					Log.New.Msg(ex);
				}
			}

			Cells.Clear();
			return sb;
		}
	}

	class DataSet
	{
//		public List<FieldCellMap> CellData { get; set; }
		public Dictionary<int, DataRow> Rows { get; set; }
		private List<int> RequiredCols { get; set; }
		public WorkSheetLayout wsLayout { get; set; }

		public DataSet(WorkSheetLayout layout)
		{
			wsLayout = layout;
			RequiredCols = layout.fields.Where(c => c.isRequired && c.fldType == FieldType.column).Select(c => c.OutputOrder).ToList();
		}

		public DataRow AddCell(DataCellValue dc)
		{
			if (Rows == null) Rows = new Dictionary<int, DataRow>();
			DataRow dr;
			if(Rows.ContainsKey(dc.rowNumber))
				dr = Rows[dc.rowNumber];
			else
			{
				dr = new DataRow(dc.rowNumber);
				dr.Sheet = this;
				Rows.Add(dc.rowNumber, dr);
			}
			dr.AddDataCell(dc);
			return dr;
		}

		public void DropRows()
		{
			if (Rows == null) return;

			var drop = Rows.Where(r => !r.Value.HasRequiredCol(RequiredCols));
			drop.ToList().ForEach(d => Rows.Remove(d.Key));
		}

		public int RecCount()
		{
			if (Rows == null) return 0;

			return Rows.Count();
		}

		public StringBuilder GetDelimitedRows(StringBuilder sb, string fldDelimiter, string rowDelimiter)
		{
			Rows.ToList().ForEach(r => { r.Value.DelimitedRow(sb, wsLayout.fields, fldDelimiter); sb.Append(rowDelimiter); });
			Rows.Clear();
			return sb;
		}

		public string GetColumnHeaders(StringBuilder sb, string fldDelimiter, string rowDelimieter)
		{
			int idx = 0;
			foreach (var c in wsLayout.fields)
			{
				sb.Append(c.Name);
				if (++idx < wsLayout.fields.Count())
					sb.Append(fldDelimiter);
			}
			sb.Append(rowDelimieter);
			return sb.ToString();
		}

		public void Write()
		{
			if (Rows == null) return;

			string fn = $"{wsLayout.OutputFileName}_{wsLayout.dst.timeStamp.ToString("yyyyMMdd_HHmmss")}.txt";

			string fp = Path.Combine(wsLayout.dst.RootFolder, fn);
			StringBuilder sb = new StringBuilder();

			if (!File.Exists(fp))
				this.GetColumnHeaders(sb, wsLayout.fldDelim, wsLayout.recDelim);

			this.GetDelimitedRows(sb, wsLayout.fldDelim, wsLayout.recDelim);

			File.AppendAllText(fp, sb.ToString());
		}
	}

	class Spreadsheet
	{
		public DataSourceTypes _dsts;
		private SharedStringTablePart stringTable;
		private CellFormats cellFormats;

		public Spreadsheet(DataSourceTypes dsts)
		{
			_dsts = dsts;
		}

		public int ProcessFile(FileInfo file)
		{
			try
			{
				using (SpreadsheetDocument ss = SpreadsheetDocument.Open(file.FullName, false))
				{
					var ssLayout = _dsts.DetermineLayout(ss, file);

					if (ssLayout == null)
					{
						return 0;
					}

					WorkbookPart wbp = ss.WorkbookPart;
					stringTable = wbp.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

					cellFormats = wbp.WorkbookStylesPart.Stylesheet.CellFormats;
					var numberingFormats = wbp.WorkbookStylesPart.Stylesheet.NumberingFormats;

					foreach (var sheetLayout in ssLayout.sLayouts.Where(ssl => ssl.wsLayout != null && ssl.srcWorksheets != null && ssl.srcWorksheets.Count() > 0))
					// Iterate through the worksheets for this data type
					{
						foreach(var sht in sheetLayout.srcWorksheets)
						{
							sheetLayout.dataSet = new DataSet(sheetLayout.wsLayout);
							ProcessCells(sheetLayout, sht.WorksheetPart);
							sheetLayout.dataSet.DropRows();
						}
					}

					int recCnt = ssLayout.sLayouts.Where(s => s.dataSet != null).Sum(s => s.dataSet.RecCount());

					ssLayout.Write();

					Log.New.Msg($"processed { recCnt } records from file: {file.Name}");

					ss.Close();

					ssLayout = null;

					return recCnt;
				}
			}
			catch (Exception ex)
			{
				Log.New.Msg(ex, $"loading file: {file.Name}");
				return 0;
			}
		}

		private void ProcessCells(SheetLayout sLayout, WorksheetPart wsp)
		{
			Dictionary<int, Field> cols = new Dictionary<int, Field>();
			sLayout.wsLayout.fieldColMap.colmaps.ForEach(cm => cols.Add(cm.column, cm.field));
			
			var tcs = wsp.Worksheet.Descendants<Cell>()
										.Where(c => c.InnerText.Length > 0)
										.Select(t => new { cell = t, row = GetRowNum(t.CellReference.InnerText), col = GetColumn(t.CellReference.InnerText) })
										.Where(k => k.row >= sLayout.wsLayout.fieldColMap.colLayout.FirstRow && cols.ContainsKey(k.col));

			foreach (var tc in tcs)
			{
				string sval = GetCellValue(tc.cell, stringTable.SharedStringTable, cellFormats, cols[tc.col]);
				var dataCell = new DataCellValue { CellReference = tc.cell.CellReference.InnerText, rowNumber = tc.row, field = cols[tc.col], Value = sval };
				sLayout.dataSet.AddCell(dataCell);
			}
		}

		public static int GetRowNum(string address)
		{
			var rwx = Regex.Replace(address, "[^0-9.]", "");
			return int.Parse(rwx);
		}

		public static int GetColumn(string address)
		{
			var rwx = Regex.Replace(address, "[0-9.]", "");
			var cls = rwx.ToLower().ToCharArray();
			int mult = 1;
			int col = 0;
			foreach(var c in cls)
			{
				var n = c - 'a' + 1;
				col += n * mult;
				mult *= 26;
			}
			return col;
		}

		public static string GetCellValue(Cell c, SharedStringTable stringTable, CellFormats formats, Field colmn)
		{
			string sval = null;
			int styleIndex;
			CellFormat cellFormat = null;

			if (c == null) return null;

			if (c.StyleIndex != null)
			{
				styleIndex = (int)c.StyleIndex.Value;
				cellFormat = (CellFormat)formats.ElementAt(styleIndex);
			}

			if (c == null || c.CellValue == null)
				return sval;

			if (c.DataType != null && c.CellFormula == null)
			{
				switch (c.DataType.Value)
				{
					case CellValues.SharedString:
						sval = stringTable.ElementAt(int.Parse(c.InnerText)).InnerText;
						break;
					default:
						sval = "not a string";
						break;
				}
			}
			else
				sval = string.IsNullOrWhiteSpace(c.CellValue.InnerText) ? null : c.CellValue.InnerText;

			if (colmn != null)
			{
				switch (colmn.DataFormat)
				{
					case DataFormatType.Date:
						if (cellFormat != null && cellFormat.NumberFormatId != null && c.DataType == null)
						{
							var d = DateTime.FromOADate(double.Parse(c.CellValue.InnerText));
							sval = d.ToString("MM/dd/yy");
						}
						else
						{
							DateTime tv;
							if(sval != null && DateTime.TryParse(sval, out tv))
							{
								if (colmn.DataFormat == DataFormatType.Date)
									sval = tv.ToShortDateString();
								else
									sval = tv.ToString();
							}
							else
								sval = null;
						}
						break;
					case DataFormatType.DateTime:
						var dt = DateTime.FromOADate(double.Parse(c.CellValue.InnerText));
						sval = dt.ToString();
						break;
				}
			}

			if (sval != null && colmn != null && colmn.postProcRegex != null && colmn.postProcRegex.Count() == 2)
				sval = Regex.Replace(sval, colmn.postProcRegex[0], colmn.postProcRegex[1]);

			if (sval != null)
			{
				//if (sval.Length > 4000)
				//{
				//	Log.New.Msg($"truncating to length 4000 cell: {c.CellReference.InnerText} contents: {sval}");
				//	sval = sval.Substring(0, 4000);
				//}

				sval = sval.Replace("\t", "");
			}

			return sval;
		}
	}
}
