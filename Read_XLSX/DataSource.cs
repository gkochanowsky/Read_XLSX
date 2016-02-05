/*
	© 2015 Florida State University. All rights reserved.

	DESC:	classes used to define data source types and their data layout
				and use this information to determine the data source type
				for a given xlsx file.

	History
	===================================================================
	2016/02/05	G.K.	Created.

*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Read_XLSX
{
	class DataSourceTypes
	{
		List<DataSourceType> types { get; set; }

		public DataSourceTypes()
		{
			Init();
		}

		public DataSourceType DetermineType(SpreadsheetDocument ssd)
		{
			DataSourceType type = null;

			WorkbookPart wbp = ssd.WorkbookPart;
			var stringTable = wbp.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
			var cellFormats = wbp.WorkbookStylesPart.Stylesheet.CellFormats;

			var shts = wbp.Workbook.Descendants<Sheet>();

			// Only look at types with same number of worksheets
			var res = types.Where(r => r.workSheets.Count() == shts.Count()).ToList();

			int idx = 0;
			foreach (var sht in shts)
			// Get list of types with matching worksheet names in sequence.
			{
				res = res.Where(r => r.workSheets.ElementAt(idx).Name == sht.Name).ToList();
				idx++;
			}

			if (res.Count() == 0) return null;

			foreach (var dst in res)
			// Iterate through types
			{
				bool isPass = true;

				foreach (var dws in dst.workSheets)
				// Iterate through worksheets for type.
				{
					if (dws.layout == null) continue;

					// Locate corresponding file worksheet based on type worksheet index.
					var sht = wbp.Workbook.Descendants<Sheet>().ElementAt(dst.workSheets.IndexOf(dws));

					WorksheetPart wsp = wbp.GetPartById(sht.Id) as WorksheetPart;
					
					isPass &= CheckSignature(wsp.Worksheet, dws.layout.specialCells, dws.layout.columns, stringTable, cellFormats);
				}

				if (isPass)
				{
					type = dst;
					break;
				}
			}
			
			return type;
		}

		public bool CheckSignature(Worksheet ws, List<SpecialCell> specialCells, List<DataColumn> dataColumns, SharedStringTablePart stringTable, CellFormats formats)
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
			var tcs_d = tcs_c.Select(t => new { cell = t, val = Spreadsheet.GetCellValue(t, stringTable.SharedStringTable, formats, null), expected = pairs[t.CellReference.InnerText] });

			// All cells where expected does not match value
			var fail = tcs_d.Where(f => f.val != f.expected);

			// Should be zero if everything matched.
			return fail.Count() == 0;
		}

		private void Init()
		{
			// Create the datasheet layouts to be used by the data source types.
			var cga = new DataWorkSheetLayout
			{
				Name = "Complaint, Grievance and Appeal Information",
				specialCells = new List<SpecialCell>
				{
					new SpecialCell { CellReference = "E5", CellName = "MedicalProviderNbrs", TitleCellReference = "B5", TitleString = "Medicaid Provider #:" },
					new SpecialCell { CellReference = "D6", CellName = "CalendarYr", TitleCellReference = "B6", TitleString = "Calendar Year:" },
					new SpecialCell { CellReference = "D7", CellName = "PlanName", TitleCellReference = "B7", TitleString = "Plan Name:" },
					new SpecialCell { CellReference = "O3", CellName = "TotalMMA", TitleCellReference = "P3", TitleString = "Total MMA" },
					new SpecialCell { CellReference = "O4", CellName = "TotalLTC", TitleCellReference = "P4", TitleString = "Total LTC" },
					new SpecialCell { CellName = "Month" }
				},
				columns = new List<DataColumn>
				{
					new DataColumn { col = 2, Name = "Region", DataFormat = DataFormatType.String, TitleCellReference = "B9", TitleString = "Region # (1 - 11)" },
					new DataColumn { col = 3, Name = "County", DataFormat = DataFormatType.String, TitleCellReference = "C10", TitleString = "County Name Within Region:" },
					new DataColumn { col = 4, Name = "MedicaidID", DataFormat = DataFormatType.String, TitleCellReference = "D9", TitleString = "Recipient's Medicaid ", isRequired = true },
					new DataColumn { col = 5, Name = "LastName", DataFormat = DataFormatType.String, TitleCellReference = "E9", TitleString = "Recipient Last" },
					new DataColumn { col = 6, Name = "FirstName", DataFormat = DataFormatType.String, TitleCellReference = "F9", TitleString = "Recipient First" },
					new DataColumn { col = 7, Name = "MiddleInitial", DataFormat = DataFormatType.String, TitleCellReference = "G9", TitleString = "Mdl" },
					new DataColumn { col = 8, Name = "GrievanceDate", DataFormat = DataFormatType.Date, TitleCellReference = "H9", TitleString = "Date of  Grievance" },
					new DataColumn { col = 9, Name = "GrievanceType", DataFormat = DataFormatType.String, TitleCellReference = "I10", TitleString = "Grievance" },
					new DataColumn { col = 10, Name = "AppealDate", DataFormat = DataFormatType.Date, TitleCellReference = "J10", TitleString = "Appeal" },
					new DataColumn { col = 11, Name = "AppealAction", DataFormat = DataFormatType.String, TitleCellReference = "K10", TitleString = "Action " },
					new DataColumn { col = 12, Name = "DispositionDate", DataFormat = DataFormatType.Date, TitleCellReference = "L10", TitleString = "Disposition" },
					new DataColumn { col = 13, Name = "DispositionType", DataFormat = DataFormatType.String, TitleCellReference = "M10", TitleString = "Disposition" },
					new DataColumn { col = 14, Name = "DispositionStatus", DataFormat = DataFormatType.String, TitleCellReference = "N8", TitleString = "Disposition Status         R=Resolved  P=Pending" },
					new DataColumn { col = 15, Name = "ExpiditedRequest", DataFormat = DataFormatType.String, TitleCellReference = "O8", TitleString = "Expedited Request   Y=yes  N=No" },
					new DataColumn { col = 16, Name = "FileType", DataFormat = DataFormatType.String, TitleCellReference = "P8", TitleString = "File Type:     GM=Griev MMA                    AM=Appeal MMA      GL=Griev LTC   AL=Appeal LTC" },
					new DataColumn { col = 17, Name = "Originator", DataFormat = DataFormatType.String, TitleCellReference = "Q10", TitleString = "2 = Provider" },
				},
				StartRow = 11
			};

			// Create list of data source types.
			types = new List<DataSourceType>
			{
				new DataSourceType
				{
					Name = "Enrollee Complaints, Grievances and Appeals Report (0127)",
					workSheets = new List<DataWorkSheet>
					{
						new DataWorkSheet { Name = "Instructions" },
						new DataWorkSheet { Name = "Codes" },
						new DataWorkSheet { Name = "January", layout = cga },
						new DataWorkSheet { Name = "February", layout = cga },
						new DataWorkSheet { Name = "March", layout = cga },
						new DataWorkSheet { Name = "April", layout = cga },
						new DataWorkSheet { Name = "May", layout = cga },
						new DataWorkSheet { Name = "June", layout = cga },
						new DataWorkSheet { Name = "July", layout = cga },
						new DataWorkSheet { Name = "August", layout = cga },
						new DataWorkSheet { Name = "September", layout = cga },
						new DataWorkSheet { Name = "October", layout = cga },
						new DataWorkSheet { Name = "November", layout = cga },
						new DataWorkSheet { Name = "December", layout = cga },
						new DataWorkSheet { Name = "Summary" }
					}
				},
			};
		}
	}

	class DataSourceType
	{
		public string Name { get; set; }
		public List<DataWorkSheet> workSheets { get; set; }
	}

	class DataWorkSheet
	{
		public string Name { get; set; }
		public DataWorkSheetLayout layout { get; set; }
	}

	class DataWorkSheetLayout
	{
		public string Name { get; set; }
		public List<SpecialCell> specialCells { get; set; }
		public List<DataColumn> columns { get; set; }
		public int StartRow { get; set; }

		public List<SpecialCell> CopySpecialCells()
		{
			return specialCells.Select(s => new SpecialCell
			{
				CellReference = s.CellReference,
				CellName = s.CellName,
				Value = s.Value,
				TitleCellReference = s.TitleCellReference,
				TitleString = s.TitleString
			}).ToList();
		}
	}

	class SpecialCell
	{
		public string CellReference { get; set; }
		public string CellName { get; set; }
		public string Value { get; set; }
		public string TitleCellReference { get; set; }
		public string TitleString { get; set; }
	}

	public enum DataFormatType
	{
		String = 1,
		DateTime,
		Date
	}

	class DataColumn
	{
		public int col { get; set; }
		public string Name { get; set; }
		public string TitleCellReference { get; set; }
		public string TitleString { get; set; }
		public bool isRequired { get; set; }

		public DataFormatType DataFormat { get; set; }
	}
}
