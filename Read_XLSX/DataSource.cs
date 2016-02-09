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
		public List<DataSourceType> types { get; set; }

		public DataSourceTypes()
		{
			Init();
		}

		/// <remarks>
		/// Two ways to determin type based on value matchWorkSheetNames
		/// - when true then worksheet names must match DataSource type in order to be a match.
		/// - when false then if any worksheets match data source layout then is a match
		/// </remarks>
		public DataSourceType DetermineType(SpreadsheetDocument ssd)
		{
			DataSourceType type = null;

			WorkbookPart wbp = ssd.WorkbookPart;
			var stringTable = wbp.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
			var cellFormats = wbp.WorkbookStylesPart.Stylesheet.CellFormats;

			var shts = wbp.Workbook.Descendants<Sheet>();

			var procTypes = types.Where(r => !r.matchWorkSheetNames || (r.matchWorkSheetNames && r.workSheets.Count() == shts.Count())).ToList();

			int idx = 0;
			foreach (var sht in shts)
			// Get list of types with matching worksheet names in sequence.
			{
				procTypes = procTypes.Where(r => !r.matchWorkSheetNames || (r.matchWorkSheetNames && r.workSheets.ElementAt(idx).Name == sht.Name)).ToList();
				idx++;
			}

			if (procTypes.Count() == 0) return null;

			foreach (var dst in procTypes)
			// Iterate through types
			{
				bool isPass = true;

				if (dst.matchWorkSheetNames)
				{
					foreach (var dws in dst.workSheets)
					// Iterate through worksheets for type.
					{
						if (dws.layout == null) continue;

						// Locate corresponding file worksheet based on type worksheet index.
						var sht = wbp.Workbook.Descendants<Sheet>().ElementAt(dst.workSheets.IndexOf(dws));

						WorksheetPart wsp = wbp.GetPartById(sht.Id) as WorksheetPart;

						isPass &= CheckSignature(wsp.Worksheet, dws.layout.specialCells, dws.layout.columns, stringTable, cellFormats);
					}
				}
				else
				{
					isPass = false;

					foreach(var dws in dst.workSheets)
					{
						if (dws.layout == null) continue;

						foreach(var sht in wbp.Workbook.Descendants<Sheet>())
						{
							WorksheetPart wsp = wbp.GetPartById(sht.Id) as WorksheetPart;

							isPass |= CheckSignature(wsp.Worksheet, dws.layout.specialCells, dws.layout.columns, stringTable, cellFormats);

							if (isPass) break;
						}

						if (isPass) break;
					}
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

			fail.ToList().ForEach(a => Log.Msg($"Expected: '{a.expected}', Found: {a.val}"));

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
					new SpecialCell { CellReference = "B3", CellName = "Month" }
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

			var frer = new DataWorkSheetLayout
			{
				Name = "Enrollee Roster and Facility Residence Report",

				specialCells = new List<SpecialCell>
				{
					new SpecialCell { CellReference = "B3", CellName = "MC_PlanName", TitleCellReference = "A3", TitleString = "Managed Care Plan Name : " },
					new SpecialCell { CellReference = "B4", CellName = "MC_PlanID", TitleCellReference = "A4", TitleString = "Managed Care Plan ID :" },
					new SpecialCell { CellReference = "B5", CellName = "Month", TitleCellReference = "A5", TitleString = "Reporting Month (MM/DD/YYYY):" }
				},

				columns = new List<DataColumn>
				{
					new DataColumn { col = 1, Name = "LastName", DataFormat = DataFormatType.String, TitleCellReference = "A7", TitleString = "Enrolee Last Name" },
					new	DataColumn { col = 2, Name = "FirstName", DataFormat = DataFormatType.String, TitleCellReference = "B7", TitleString = "Enrolee First Name" },
					new DataColumn { col = 3, Name = "MedicaidID", DataFormat = DataFormatType.String, TitleCellReference = "C7", TitleString = "Medicaid ID", isRequired = true },
					new DataColumn { col = 4, Name = "SSN", DataFormat = DataFormatType.String, TitleCellReference = "D7", TitleString = "Social Security Number" },
					new DataColumn { col = 5, Name = "DOB", DataFormat = DataFormatType.Date, TitleCellReference = "E7", TitleString = "Date of Birth (mm/dd/yyyy)" },
					new DataColumn { col = 6, Name = "Addr", DataFormat = DataFormatType.String, TitleCellReference = "F7", TitleString = "Physical Address" },
					new DataColumn { col = 7, Name = "City", DataFormat = DataFormatType.String, TitleCellReference = "G7", TitleString = "City" },
					new DataColumn { col = 8, Name = "Zip", DataFormat = DataFormatType.String, TitleCellReference = "H7", TitleString = "Zip Code" },
					new DataColumn { col = 9, Name = "County", DataFormat = DataFormatType.String, TitleCellReference = "I7", TitleString = "County of Residence" },
					new DataColumn { col = 9, Name = "Setting", DataFormat = DataFormatType.String, TitleCellReference = "J7", TitleString = "Residential Setting Type (Home, ALF, SNF or AFCH)" },
					new DataColumn { col = 10, Name = "FacilityName", DataFormat = DataFormatType.String, TitleCellReference = "K7", TitleString = "Name of Facility" },
					new DataColumn { col = 11, Name = "FacilityLic", DataFormat = DataFormatType.String, TitleCellReference = "L7", TitleString = "Facility License Number" },
					new DataColumn { col = 12, Name = "Tansition", DataFormat = DataFormatType.String, TitleCellReference = "M7", TitleString = "Identify if transitioning into a SNF or back into Community (SNF, Community, or N/A)"},
					new DataColumn { col = 13, Name = "TransistionDate", DataFormat = DataFormatType.Date, TitleCellReference = "N7", TitleString = "Date of transition to SNF or Community (if applicable)" },
					new DataColumn { col = 14, Name = "Form2515Date", DataFormat = DataFormatType.Date, TitleCellReference = "O7", TitleString = "Date the 2515 form was sent to DCF if transitioning (if applicable)" },
					new DataColumn { col = 15, Name = "canLocate", DataFormat = DataFormatType.String, TitleCellReference = "P7", TitleString = "			Able to Locate?" + System.Environment.NewLine + "			Y/N" },
					new DataColumn { col = 16, Name = "canContact", DataFormat = DataFormatType.String, TitleCellReference = "Q7", TitleString = "			Able to Contact?" + System.Environment.NewLine + "			Y/N" },
					new DataColumn { col = 17, Name = "LastContaceDate", DataFormat = DataFormatType.Date, TitleCellReference = "R7", TitleString = "If unable to contact or locate enrolee, date of last contact? (N/A if not applicable)" },
					new DataColumn { col = 18, Name = "Comments", DataFormat = DataFormatType.String, TitleCellReference = "S7", TitleString = "Comments including demonstration of attempts to contact enrolee if applicable" }
				},

				StartRow = 8
			};

			var efrr = new DataWorkSheetLayout
			{
				Name = "Enrollee Facility Residence Report ",

				specialCells = new List<SpecialCell>
				{
					new SpecialCell { CellReference = "A3", CellName = "MC_PlanName" },
					new SpecialCell { CellReference = "A4", CellName = "MC_PlanID" },
					new SpecialCell { CellReference = "A5", CellName = "Month" }
				},

				columns = new List<DataColumn>
				{
					new DataColumn { col = 1, Name = "LastName", DataFormat = DataFormatType.String, TitleCellReference = "A7", TitleString = "Last Name" },
					new DataColumn { col = 2, Name = "FirstName", DataFormat = DataFormatType.String, TitleCellReference = "B7", TitleString = "First Name" },
					new DataColumn { col = 3, Name = "MedicaidID", DataFormat = DataFormatType.String, TitleCellReference = "C7", TitleString = "Medicaid ID", isRequired = true },
					new DataColumn { col = 4, Name = "SSN", DataFormat = DataFormatType.String, TitleCellReference = "D7", TitleString = "Social Security Number" },
					new DataColumn { col = 5, Name = "DOB", DataFormat = DataFormatType.Date, TitleCellReference = "E7", TitleString = "Date of Birth (mm/dd/yyyy)" },
					new DataColumn { col = 6, Name = "Addr", DataFormat = DataFormatType.String, TitleCellReference = "F7", TitleString = "Physical Address\n(full street address)" },
					new DataColumn { col = 7, Name = "City", DataFormat = DataFormatType.String, TitleCellReference = "G7", TitleString = "City" },
					new DataColumn { col = 8, Name = "Zip", DataFormat = DataFormatType.String, TitleCellReference = "H7", TitleString = "Zip Code" },
					new DataColumn { col = 9, Name = "County", DataFormat = DataFormatType.String, TitleCellReference = "I7", TitleString = "County of Residence" },
					new DataColumn { col = 9, Name = "Setting", DataFormat = DataFormatType.String, TitleCellReference = "J7", TitleString = "Type of Facility" },
					new DataColumn { col = 10, Name = "FacilityName", DataFormat = DataFormatType.String, TitleCellReference = "K7", TitleString = "Name of Facility" },
					new DataColumn { col = 11, Name = "FacilityLic", DataFormat = DataFormatType.String, TitleCellReference = "L7", TitleString = "Facility License Number" },
				},

				StartRow = 8
			};

			// Create list of data source types.
			types = new List<DataSourceType>
			{
				new DataSourceType
				{
					Name = "Enrollee Complaints, Grievances and Appeals Report (0127)",
					outputFileName = "Complaint_Greivance_Appeal_Info_0127",
					matchWorkSheetNames = true,
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

				new DataSourceType
				{
					Name = "Enrollee Facility Residence Report ",
					outputFileName = "Enrollee_Facility_Residence_0129",
					matchWorkSheetNames = false,
					workSheets = new List<DataWorkSheet>
					{
						new DataWorkSheet { layout = frer }
					}
				},

				new DataSourceType
				{
					Name = "Enrollee Facility Residence Report",
					outputFileName = "Enrollee_Facility_Residence_0129_v2",
					matchWorkSheetNames = false,
					workSheets = new List<DataWorkSheet>
					{
						new DataWorkSheet { layout = efrr }
					}
				}
			};
		}
	}

	class DataSourceType
	{
		public string Name { get; set; }
		public string outputFileName { get; set; }

		public bool matchWorkSheetNames { get; set; }
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
