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
		public int colNumber { get; set; }
		public DataRow Row { get; set; }
	}

	class GroupCellValue
	{
		public DataCellValue cell { get; set; }
		public string grpValue { get; set; }
		public int row { get; set; }
		public Field field { get; set; }
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
			// list of required columns with a value
			var cols = Cells.Where(c => c.Value.field.isRequired && !string.IsNullOrWhiteSpace(c.Value.Value)).Select(c => c.Key).ToList();
			var missing = reqCol.Where(r => !cols.Contains(r));
			return missing.Count() == 0;
		}

		public StringBuilder DelimitedRow(StringBuilder sb, List<Field> datacols, string delim, SpreadSheetLayout ssl)
		{
			bool isFirst = true;
			datacols.OrderBy(d => d.OutputOrder).ToList().ForEach(d =>
			{
				try
				{
					string val = null;

					switch (d.fldType)
					{
						case FieldType.column:
							val = Cells.ContainsKey(d.OutputOrder) ? Cells[d.OutputOrder].Value ?? "" : "";
							break;
						case FieldType.cell:
						case FieldType.fileName:
						case FieldType.filePath:
							var df = this.Sheet.wsLayout.fieldCellMap.fldmaps.FirstOrDefault(fm => fm.field.OutputOrder == d.OutputOrder);
							val = (df != null ? df.Value ?? "" : "");
							break;
					}

					sb.Append((isFirst ? "" : delim) + val);
					isFirst = false;
				}
				catch (Exception ex)
				{
					Log.New.Msg(ex);
				}
			});

			ssl.sLayouts.Where(sl => sl.sheetType == SheetType.CommonData).ToList().ForEach(sl =>
			{
				sl.wsLayout.fieldCellMap.fldmaps.OrderBy(fm => fm.field.OutputOrder).ToList().ForEach(fm =>
				{
					sb.Append(delim);
					sb.Append(fm.Value ?? "");
				});
			});

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

		public void ProcessRows()
		{
			if (Rows == null) return;

			// Drop cells in ignore lists
			var ignflds = wsLayout.fields.Where(f => f.ignore != null && f.ignore.Count() > 0);

			var ignCells = Rows.SelectMany(r => r.Value.Cells.Where(ce => ignflds.Contains(ce.Value.field)).Select(ce => ce.Value));
			ignCells = ignCells.Where(ce => ce.field.ignore.Select(ig => ig.ToLower().Trim()).Contains(ce.Value.Trim().ToLower()));
			ignCells.ToList().ForEach(igc =>
			{
				igc.Row.Cells.Remove(igc.field.OutputOrder);
			});

			// locate any group rows.
			var grpCols = wsLayout.fieldColMap.colLayout.titleLocations.Where(tl => tl.isGroupData).Select(tl => tl.col).ToList();

			List<GroupCellValue> grpCells = new List<GroupCellValue>();

			if (grpCols.Count() > 0)
			{
				var flds = wsLayout.fields.Where(f => f.rowType == RowType.GroupData);

				var grpRows = Rows.Where(rw => rw.Value.Cells.Any(c => grpCols.Contains(c.Value.colNumber) && rw.Value.Cells.Count() == 1));

				// Find and map group rows to grp fields.
				foreach (var r in grpRows)
				{
					foreach (var c in r.Value.Cells)
					{
						foreach (var f in flds)
						{
							GroupCellValue grpCell = null;

							foreach (var t in f.titles)
							{
								if (c.Value.Value.ToLower().StartsWith(t.ToLower()))
								{
									grpCell = new GroupCellValue { cell = c.Value, grpValue = c.Value.Value.Substring(t.Length).Trim(), row = c.Value.rowNumber, field = f };
									break;
								}
							}

							if (grpCell != null)
							{
								grpCells.Add(grpCell);
								break;
							}
						}
					}
				}
			}

			// Drop rows with insufficient required fields.
			var drop = Rows.Where(r => !r.Value.HasRequiredCol(RequiredCols));
			drop.ToList().ForEach(d => Rows.Remove(d.Key));

			// Add groups to remaining non-group rows.
			foreach(var gc in grpCells)
			{
				int currIdx = grpCells.IndexOf(gc);
				int lastRow = currIdx + 1 < grpCells.Count() ? grpCells[currIdx + 1].row : Rows.Max(w => w.Key) + 1;

				Rows
					.Where(r => r.Key > gc.row && r.Key < lastRow)
					.ToList()
					.ForEach(r =>
					{
						int nxt = r.Value.Cells.Max(c => c.Key);
						r.Value.Cells.Add(nxt + 1, new DataCellValue { Value = gc.grpValue, CellReference = gc.cell.CellReference, field = gc.field, Row = r.Value, rowNumber = r.Key, colNumber = nxt });
					});
			}
		}

		public int RecCount()
		{
			if (Rows == null) return 0;

			return Rows.Count();
		}

		public StringBuilder GetDelimitedRows(StringBuilder sb, string fldDelimiter, string rowDelimiter, SpreadSheetLayout ssl)
		{
			Rows.ToList().ForEach(r => 
			{
				r.Value.DelimitedRow(sb, wsLayout.fields, fldDelimiter, ssl);
				sb.Append(rowDelimiter);
			});

			Rows.Clear();
			return sb;
		}

		public string GetColumnHeaders(StringBuilder sb, string fldDelimiter, string rowDelimieter, SpreadSheetLayout ssl)
		{
			bool isFirst = true;
			foreach (var c in wsLayout.fields)
			{
				if (!isFirst)
					sb.Append(fldDelimiter);

				sb.Append(c.Name);
				
				isFirst = false;
			}

			var fldtypes = new List<FieldType> { FieldType.cell, FieldType.fileName, FieldType.filePath };
			ssl.sLayouts.Where(sl => sl.sheetType == SheetType.CommonData).ToList().ForEach(sl =>
			{
				sl.wsLayout.fields.OrderBy(f => f.OutputOrder).Where(f => fldtypes.Contains(f.fldType)).ToList().ForEach(f =>
				{
					sb.Append(fldDelimiter);
					sb.Append(f.Name);
				});
			});

			sb.Append(rowDelimieter);
			return sb.ToString();
		}

		public void Write(SpreadSheetLayout ssl)
		{
			if (Rows == null) return;

			string fn = $"{wsLayout.OutputFileName}_{wsLayout.dst.timeStamp.ToString("yyyyMMdd_HHmmss")}.txt";

			string fp = Path.Combine(wsLayout.dst.RootFolder, fn);
			StringBuilder sb = new StringBuilder();

			if (!File.Exists(fp))
				this.GetColumnHeaders(sb, wsLayout.fldDelim, wsLayout.recDelim, ssl);

			this.GetDelimitedRows(sb, wsLayout.fldDelim, wsLayout.recDelim, ssl);

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
					List<LayoutType> allowedLT = new List<LayoutType> { LayoutType.Both, LayoutType.ColumnOnly };

					foreach (var sheetLayout in ssLayout.sLayouts.Where(ssl => ssl.wsLayout != null && ssl.srcWorksheets != null && ssl.srcWorksheets.Count() > 0 && allowedLT.Contains( ssl.wsLayout.layoutType)))
					// Iterate through the worksheets for this data type
					{
						foreach(var sht in sheetLayout.srcWorksheets)
						{
							sheetLayout.dataSet = new DataSet(sheetLayout.wsLayout);
							ProcessCells(sheetLayout, sht.WorksheetPart);
							sheetLayout.dataSet.ProcessRows();
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

			IEnumerable<CellRowCol> tcs = Enumerable.Empty<CellRowCol>();

			switch (sLayout.wsLayout.fieldColMap.colLayout.colLayoutType)
			{
				case ColLayoutType.Row_Col:
					tcs = wsp.Worksheet.Descendants<Cell>()
												.Where(c => c.InnerText.Length > 0)
												.Select(t => new CellRowCol { cell = t, row = GetRowNum(t.CellReference.InnerText), col = GetColumn(t.CellReference.InnerText) })
												.Where(k => k.row >= sLayout.wsLayout.fieldColMap.colLayout.FirstRow && cols.ContainsKey(k.col));
					break;
				case ColLayoutType.Col_Row:
					tcs = wsp.Worksheet.Descendants<Cell>()
												.Where(c => c.InnerText.Length > 0)
												.Select(t => new CellRowCol { cell = t, col = GetRowNum(t.CellReference.InnerText), row = GetColumn(t.CellReference.InnerText) })
												.Where(k => k.row >= sLayout.wsLayout.fieldColMap.colLayout.FirstRow && cols.ContainsKey(k.col));
					break;
			}


			foreach (var tc in tcs)
			{
				try
				{
					string sval = GetCellValue(tc.cell, stringTable.SharedStringTable, cellFormats, cols[tc.col]);
					var dataCell = new DataCellValue { CellReference = tc.cell.CellReference.InnerText, rowNumber = tc.row, colNumber = tc.col, field = cols[tc.col], Value = sval };
					sLayout.dataSet.AddCell(dataCell);
				}
				catch (Exception ex)
				{
					Log.New.Msg(ex);
				}
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

		public static string GetCellRef(int row, int col)
		{
			string colRef = GetColRef(col);
			return colRef + row.ToString();
		}

		/// <summary>
		/// col starts at 1
		/// </summary>
		/// <param name="col"></param>
		/// <returns></returns>
		public static string GetColRef(int col)
		{
			int plc = (col - 1) % 26;
			int bal = (col - 1) / 26;
			var plcStr = ((char)('A' + plc)).ToString();
			var plcRef = (bal > 0 ? GetColRef(bal) : "") + plcStr;
			return plcRef;
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
						var pdt = c.CellValue.InnerText;
						pdt = FixDate(pdt);
						if (cellFormat != null && cellFormat.NumberFormatId != null && c.DataType == null)
						{
							var d = DateTime.FromOADate(double.Parse(pdt));
							sval = d.ToString("MM/dd/yyyy");
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
					case DataFormatType.DateMixed:
						double aoDate;
						if (double.TryParse(c.CellValue.InnerText, out aoDate))
						{
							var aodt = DateTime.FromOADate(aoDate);
							if ((DateTime.Now - aodt).Days < 36500)
								sval = aodt.ToString("MM/dd/yyyy");
						}
						break;
				}
			}

			if (sval != null && colmn != null && colmn.postProcRegex != null)
				foreach(var d in colmn.postProcRegex)
					sval = Regex.Replace(sval, d.Item1, d.Item2);

			if (sval != null)
			{
				//if (sval.Length > 4000)
				//{
				//	Log.New.Msg($"truncating to length 4000 cell: {c.CellReference.InnerText} contents: {sval}");
				//	sval = sval.Substring(0, 4000);
				//}

				sval = sval.Replace("\t", "");
				sval.Trim();
			}

			return sval;
		}

		private static string FixDate(string dt)
		{
			var mnth = Config.Data.Months.Where(m => dt.StartsWith(m)).FirstOrDefault();
			if(mnth == null)
				return dt;

			var prts = dt.Split('-');
			if (prts.Count() != 2)
				return dt;

			int yr;
			if (!int.TryParse(prts[1], out yr))
				return dt;

			var mon = Config.Data.Months.IndexOf(mnth);

			if (yr < 1000)
				yr += 2000;

			dt = $"{mon}/{yr}";

			return dt;
		}
	}

	class CellRowCol
	{
		public Cell cell { get; set; }
		public int row { get; set; }
		public int col { get; set; }
	}
}
