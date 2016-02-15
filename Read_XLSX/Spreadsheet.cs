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
		public DataSheet Sheet { get; set; }
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
				string val = null;

				switch (v.fldType)
				{
					case FieldType.column:
						val = Cells.ContainsKey(v.OutputOrder) ? Cells[v.OutputOrder].Value ?? "" : "";
						break;
					case FieldType.cell:
					case FieldType.fileName:
						var df = this.Sheet.CellData.FirstOrDefault(cd => cd.field.OutputOrder == v.OutputOrder);
						val = (df != null ? df.Value ?? "" : "");
						break;
				}

				// Don't want a delimiter after last column.
				sb.Append(val + (v.OutputOrder < lastCol ? delim : ""));
			}

			return sb;
		}
	}

	class DataSheet
	{
		public string Name { get; set; }
		public List<FieldCellMap> CellData { get; set; }
		public List<Field> Fields { get; set; }
		public Dictionary<int, DataRow> Rows { get; set; }
		private List<int> RequiredCols { get; set; }
		public DataFile File { get; set; }

		public DataSheet(string name, List<Field> fields, List<FieldCellMap> cellData)
		{
			Name = name;
			CellData = cellData;
			Fields = fields;
			RequiredCols = fields.Where(c => c.isRequired && c.fldType == FieldType.column).Select(c => c.OutputOrder).ToList();
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
			return Rows.Count();
		}

		public StringBuilder GetDelimitedRows(StringBuilder sb, string fldDelimiter, string rowDelimiter)
		{
			if (sb.Length == 0)
				GetColumnHeaders(sb, fldDelimiter, rowDelimiter);
			Rows.ToList().ForEach(r => { r.Value.DelimitedRow(sb, Fields, fldDelimiter); sb.Append(rowDelimiter); });
			return sb;
		}

		public StringBuilder GetColumnHeaders(StringBuilder sb, string fldDelimiter, string rowDelimieter)
		{
			int idx = 0;
			foreach (var c in Fields)
			{
				sb.Append(c.Name);
				if (++idx < Fields.Count())
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
		public SpreadSheetLayout dst { get; set; }

		public DataFile(string filePath)
		{
			FilePath = filePath;
		}

		public DataSheet AddDataSheet(string name, List<Field> colData, List<FieldCellMap> cellData)
		{
			if (DataSheets == null)
				DataSheets = new List<DataSheet>();

			var ds = new DataSheet(name, colData, cellData);
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
		public DataSourceTypes _dsts;
		private SharedStringTablePart stringTable;
		private CellFormats cellFormats;

		public Spreadsheet(DataSourceTypes dsts)
		{
			_dsts = dsts;
		}

		public DataFile ProcessFile(FileInfo file)
		{
			var dataFile = new DataFile(file.FullName);

			try
			{
				using (SpreadsheetDocument ss = SpreadsheetDocument.Open(file.FullName, false))
				{
					dataFile.dst = _dsts.DetermineLayout(ss, file);

					if (dataFile.dst == null)
					{
						Log.Msg($"FAILURE: {file.FullName}: Unable to determine format type of file");
						return dataFile;
					}

					WorkbookPart wbp = ss.WorkbookPart;
					stringTable = wbp.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

					cellFormats = wbp.WorkbookStylesPart.Stylesheet.CellFormats;
					var numberingFormats = wbp.WorkbookStylesPart.Stylesheet.NumberingFormats;

					foreach (var sheetLayout in dataFile.dst.ssLayout.Where(ssl => ssl.wsLayout != null && ssl.srcWorksheets != null && ssl.srcWorksheets.Count() > 0))
					// Iterate through the worksheets for this data type
					{
						foreach(var sht in sheetLayout.srcWorksheets)
						{
							var dataSheet = dataFile.AddDataSheet(sheetLayout.Name, sheetLayout.wsLayout.fields, sheetLayout.wsLayout.fieldCellMap.fldmaps);
							ProcessCells(sheetLayout, sht.WorksheetPart, dataSheet);
						}
					}
				}

				dataFile.DropRows();
				Log.Msg($"processed {dataFile.RecCount()} records from file: {file.Name}");
				return dataFile;

			}
			catch (Exception ex)
			{
				Log.Msg(ex, $"loading file: {file.Name}");
				return null;
			}
		}

		private void ProcessCells(SheetLayout sLayout, WorksheetPart wsp, DataSheet dataSheet)
		{
			Dictionary<int, Field> cols = new Dictionary<int, Field>();
			sLayout.wsLayout.colLayoutVersionMap.colmaps.ForEach(cm => cols.Add(cm.column, cm.field));
			
			var tcs = wsp.Worksheet.Descendants<Cell>()
										.Where(c => c.InnerText.Length > 0)
										.Select(t => new { cell = t, row = GetRowNum(t.CellReference.InnerText), col = GetColumn(t.CellReference.InnerText) })
										.Where(k => k.row >= sLayout.wsLayout.FirstRow && cols.ContainsKey(k.col));

			foreach (var tc in tcs)
			{
				string sval = GetCellValue(tc.cell, stringTable.SharedStringTable, cellFormats, cols[tc.col]);
				var dataCell = new DataCellValue { CellReference = tc.cell.CellReference.InnerText, rowNumber = tc.row, field = cols[tc.col], Value = sval };
				dataSheet.AddCell(dataCell);
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

			if (c == null || c.CellValue == null)
				return sval;

			if (c.StyleIndex != null)
			{
				styleIndex = (int)c.StyleIndex.Value;
				cellFormat = (CellFormat)formats.ElementAt(styleIndex);
			}


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
	}
}
