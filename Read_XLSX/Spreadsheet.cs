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

namespace Read_XLSX
{


	class DataCell
	{
		public string CellReference { get; set; }
		public int Col { get; set; }
		public string Value { get; set; }
		public DataRow Row { get; set; }
	}

	class DataRow
	{
		public int Row { get; set; }
		public List<DataCell> Cells { get; set; }

		public DataSheet Sheet { get; set; }

		public DataRow(int row)
		{
			Row = row;
		}

		public DataCell AddDataCell(DataCell dc)
		{
			if (Cells == null) Cells = new List<DataCell>();

			if (Cells.Where(c => c.CellReference == dc.CellReference).Count() > 0)
				throw new Exception("Cell has already been recorded");
			dc.Row = this;
			Cells.Add(dc);

			return dc;
		}

		public bool HasRequiredCol(List<int> reqCol)
		{
			var cc = Cells.Select(c => c.Col);
			var missing = reqCol.Where(r => !cc.Contains(r));
			return missing.Count() == 0;
		}

		public StringBuilder DelimitedRow(StringBuilder sb, List<DataColumn> datacols, string delim)
		{
			// Build list of columns from column list and cell data. Missing cells will have blank column value.
			var cols = datacols.Select(dc => { var c = Cells.FirstOrDefault(cl => cl.Col == dc.col); return new DataCell { Value = (c == null ? "" : c.Value ?? "") }; }).ToList();
			cols.AddRange(Sheet.addedCells);
			// Don't want a delimiter after last column.
			int idx = 0;
			foreach(var v in cols)
			{
				sb.Append(v.Value ?? "");
				if (++idx < cols.Count())
					sb.Append(delim);
			}
			
			return sb;
		}
	}





	class DataSheet
	{
		public string Name { get; set; }
		public List<SpecialCell> SpecialCells { get; set; }
		public List<DataColumn> DataColumns { get; set; }
		public List<DataRow> Rows { get; set; }

		public List<DataCell> addedCells { get { return SpecialCells.Select(s => new DataCell { Value = s.Value }).ToList(); } }

		private List<int> RequiredColumns { get; set; }
		public DataFile File { get; set; }

		public DataSheet(string name, List<DataColumn> dataColumns, List<SpecialCell> specialCells)
		{
			Name = name;
			SpecialCells = specialCells;
			DataColumns = dataColumns;
			RequiredColumns = dataColumns.Where(c => c.isRequired).Select(c => c.col).ToList();
		}

		public DataRow AddCell(DataCell dc)
		{
			int row = Spreadsheet.GetRow(dc.CellReference);
			if (Rows == null) Rows = new List<DataRow>();
			var dr = Rows.Where(r => r.Row == row).SingleOrDefault();
			if (dr == null)
			{
				dr = new DataRow(row);
				dr.Sheet = this;
				Rows.Add(dr);
			}
			dr.AddDataCell(dc);
			return dr;
		}

		public void DropRows()
		{
			if (Rows == null) return;

			var drop = Rows.Where(r => !r.HasRequiredCol(RequiredColumns));
			drop.ToList().ForEach(d => Rows.Remove(d));
		}

		public int RecCount()
		{
			return Rows.Count();
		}

		public StringBuilder GetDelimitedRows(StringBuilder sb, string fldDelimiter, string rowDelimiter)
		{
			if (sb.Length == 0)
				GetColumnHeaders(sb, fldDelimiter, rowDelimiter);
			Rows.ForEach(r => { r.DelimitedRow(sb, DataColumns, fldDelimiter); sb.Append(rowDelimiter); });
			return sb;
		}

		public StringBuilder GetColumnHeaders(StringBuilder sb, string fldDelimiter, string rowDelimieter)
		{
			var cols = new List<DataColumn>();
			cols.AddRange(DataColumns);
			cols.AddRange(SpecialCells.Select(s => new DataColumn { Name = s.CellName }).ToList());
			int idx = 0;
			foreach(var c in cols)
			{
				sb.Append(c.Name);
				if (++idx < cols.Count())
					sb.Append(fldDelimiter);
			}
			sb.Append(rowDelimieter);
			return sb;
		}
	}

	class DataFile
	{
		public string FilePath { get; set; }
		public List<DataSheet> DataSheets { get; set; }

		public DataFile(string filePath)
		{
			FilePath = filePath;
		}

		public DataSheet AddDataSheet(string name, List<DataColumn> dataColumns, List<SpecialCell> specialCells)
		{
			if (DataSheets == null)
				DataSheets = new List<DataSheet>();

			var ds = new DataSheet(name, dataColumns, specialCells);
			ds.File = this;
			DataSheets.Add(ds);
			return ds;
		}

		public void DropRows()
		{
			DataSheets.ForEach(s => s.DropRows());
			var ds = DataSheets.Where(d => d.Rows == null || d.Rows.Count() == 0);
			ds.ToList().ForEach(d => DataSheets.Remove(d));
		}

		public int RecCount()
		{
			return DataSheets.Sum(d => d.RecCount());
		}

		public StringBuilder GetDelimitedRows(StringBuilder sb, string fldDelimiter, string rowDelimiter)
		{

			DataSheets.ForEach(s => { s.GetDelimitedRows(sb, fldDelimiter, rowDelimiter); });
			return sb;
		}
	}

	class Spreadsheet
	{
		public Spreadsheet()
		{

		}

		public static DataFile ProcessFile(string filePath)
		{
			var dataFile = new DataFile(filePath);
			var fileName = System.IO.Path.GetFileName(filePath);

			var dst = new DataSourceTypes();

			try
			{
				using (SpreadsheetDocument ss = SpreadsheetDocument.Open(filePath, false))
				{
					var type = dst.DetermineType(ss);

					WorkbookPart wbp = ss.WorkbookPart;
					var stringTable = wbp.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

					var cellFormats = wbp.WorkbookStylesPart.Stylesheet.CellFormats;
					var numberingFormats = wbp.WorkbookStylesPart.Stylesheet.NumberingFormats;

					foreach (var dws in type.workSheets)
					// Iterate through the worksheets for this data type
					{
						var specialCells = dws.layout.CopySpecialCells();

						// Ignore data worksheets with no layout data.
						if (dws.layout == null) continue;

						// Locate file worksheet that cooresponds data layout.
						var sht = wbp.Workbook.Descendants<Sheet>().ElementAt(type.workSheets.IndexOf(dws));

						// Locate 
						var cols = dws.layout.columns.Select(dc => dc.col);
						var sref = specialCells.Where(s => s.CellReference != null).Select(s => s.CellReference);

						var dataSheet = dataFile.AddDataSheet(sht.Name, dws.layout.columns, specialCells);

						WorksheetPart wsp = wbp.GetPartById(sht.Id) as WorksheetPart;

						var tcs = wsp.Worksheet.Descendants<Cell>()
										.Where(c => c.InnerText.Length > 0)
										.Select(t => new { cell = t, row = GetRow(t.CellReference.InnerText), col = GetColumn(t.CellReference.InnerText) })
										.Where(k => sref.Contains(k.cell.CellReference.InnerText) || (k.row >= dws.layout.StartRow && cols.Contains(k.col)));

						string sval = null;

						var scMonth = specialCells.Where(cs => cs.CellName == "Month").FirstOrDefault();
						scMonth.Value = sht.Name;

						foreach (var tc in tcs)
						{
							var sc = specialCells.Where(cs => cs.CellReference == tc.cell.CellReference.InnerText).FirstOrDefault();
							int row = tc.row;
							int col = tc.col;
							DataColumn colmn = dws.layout.columns.Where(dc => row > 10 && dc.col == col).FirstOrDefault();

							sval = GetCellValue(tc.cell, stringTable.SharedStringTable, cellFormats, colmn);

							if (sc != null)
								sc.Value = sval;
							else {
								var dataCell = new DataCell { CellReference = tc.cell.CellReference.InnerText, Col = col, Value = sval };
								dataSheet.AddCell(dataCell);
							}
						}
					}

					dataFile.DropRows();
					Log.Msg($"processed {dataFile.RecCount()} records from file: {fileName}");
					return dataFile;
				}
			}
			catch (Exception ex)
			{
				Log.Msg(ex, $"loading file: {fileName}");
				return null;
			}
		}

		public static int GetRow(string address)
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

		public static string GetCellValue(Cell c, SharedStringTable stringTable, CellFormats formats, DataColumn colmn)
		{
			string sval = null;
			var styleIndex = (int)c.StyleIndex.Value;
			var cellFormat = (CellFormat)formats.ElementAt(styleIndex);

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
						if (cellFormat.NumberFormatId != null && c.DataType == null)
						{
							var d = DateTime.FromOADate(double.Parse(c.CellValue.InnerText));
							sval = d.ToString("MM/dd/yy");
						}
						else
							sval = null;
						break;
					case DataFormatType.DateTime:
						var dt = DateTime.FromOADate(double.Parse(c.CellValue.InnerText));
						sval = dt.ToString();
						break;
				}
			}

			return sval;
		}

		/// <summary>
		/// Checks the TitleCellsReference values against expected values. If everything matches then this is a pass.
		/// </summary>
		/// <param name="ws"></param>
		/// <param name="specialCells"></param>
		/// <param name="dataColumns"></param>
		/// <param name="stringTable"></param>
		/// <param name="formats"></param>
		/// <returns></returns>
		public static bool CheckSignature(Worksheet ws, List<SpecialCell> specialCells, List<DataColumn> dataColumns, SharedStringTablePart stringTable, CellFormats formats)
		{
			// Create a dictionary of cell reference expected value pairs for expected title cells
			Dictionary<string, string> pairs = new Dictionary<string, string>();
			specialCells.Where(sc => sc.TitleCellReference != null && sc.TitleString != null).ToList().ForEach(sc => pairs.Add(sc.TitleCellReference, sc.TitleString));
			dataColumns.Where(dc => dc.TitleCellReference != null && dc.TitleString != null).ToList().ForEach(dc => pairs.Add(dc.TitleCellReference, dc.TitleString));

			// All references to check
			var refs = pairs.Select(p => p.Key).ToList();

			// All cells in worksheet.
			var tcs = ws.Descendants<Cell>();

			// All referenced cells to check
			var tcs_c = tcs.Where(t => refs.Contains(t.CellReference.InnerText));

			// All referenced cells with computed and expected value.
			var tcs_d = tcs_c.Select(t => new { cell = t, val = GetCellValue(t, stringTable.SharedStringTable, formats, null), expected = pairs[t.CellReference.InnerText] });

			// All cells where expected does not match value
			var fail = tcs_d.Where(f => f.val != f.expected);

			// Should be zero if everything matched.
			return fail.Count() == 0;
		}
	}
}
