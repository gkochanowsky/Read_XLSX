using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_XLSX
{
	class Config
	{
		public static readonly Config Data = new Config();

		public readonly List<string> Months;

		private Config()
		{
			Months = new List<string> { "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC" };
		}

		public static List<SpreadSheetLayout> Load(DataSourceTypes dst)
		{
			var wsLayout_cga = new WorkSheetLayout
			{
				Name = "Complaint, Grievance and Appeal Information",
				OutputFileName = "Data_Extract_Complaint_Grievance_Appeal_Info_0127",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.Both,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion { Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "B5", ValueRef = "E5" },
							new CellLocation { TitleRef = "B6", ValueRef = "D6" },
							new CellLocation { TitleRef = "B7", ValueRef = "D7" },
							new CellLocation { TitleRef = "B3", ValueRef = "B3" }
						}
					},
					new CellLayoutVersion { Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "B6", ValueRef = "C6" },
							new CellLocation { TitleRef = "B7", ValueRef = "C7" },
							new CellLocation { TitleRef = "B3", ValueRef = "B3" }
						}
					},
					new CellLayoutVersion { Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A5", ValueRef = "D5" },
							new CellLocation { TitleRef = "A6", ValueRef = "C6" },
							new CellLocation { TitleRef = "A7", ValueRef = "C7" },
							new CellLocation { TitleRef = "A3", ValueRef = "A3" }
						}
					},
					new CellLayoutVersion { Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A6", ValueRef = "C6" },
							new CellLocation { TitleRef = "A7", ValueRef = "C7" },
							new CellLocation { TitleRef = "A8", ValueRef = "C8" },
							new CellLocation { TitleRef = "A2", ValueRef = "A2" }
						}
					},
				},

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						titleLocations = new List<ColumnTitleLocation> {
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B9" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C10" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D9", "D10" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E9", "E10" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F9", "F10" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G9", "G10" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H9" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I9", "I10" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J9", "J10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K9", "K10" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L9", "L10" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M9", "M10" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N8" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O8" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P8" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "Q9","Q10" } },
						},
						FirstRow = 11
					},
					new ColumnLayoutVersion
					{
						Version = 2,
						titleLocations = new List<ColumnTitleLocation> {
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B9" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C9", "C10" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D9", "D10" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E9", "E10" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F9", "F10" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G9" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H9", "H10" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I9", "I10" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J9", "J10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K9", "K10" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L9", "L10" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M8" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N8" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O8" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P9", "P10" } },
						},
						FirstRow = 11
					},
					new ColumnLayoutVersion
					{
						Version = 3,
						titleLocations = new List<ColumnTitleLocation> {
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B10" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C9", "C10" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D9", "D10" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E9", "E10" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F9", "F10" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G9" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H9", "H10" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I9", "I10" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J9", "J10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K9", "K10" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L9", "L10" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M8" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N8" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O8" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P9", "P10" } },
						},
						FirstRow = 11
					},
					new ColumnLayoutVersion
					{
						Version = 4,
						titleLocations = new List<ColumnTitleLocation> {
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A10", "A11" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B10" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C10" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D10" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E10" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F10", "F11" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G10" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H10" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I10" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K10" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L10" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M10" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N10" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O10" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P10" } },
						},
						FirstRow = 11
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "Region", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Region # (1 - 11)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "County", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "County Name Within Region:" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "MedicaidID", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Recipient's Medicaid ID#:",
							"Recipient's Medicaid ID#",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "LastName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Recipient LastName:",
							"Recipient Last Name",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "FirstName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Recipient FirstName:",
							"Recipient First Name",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "MiddleInitial", DataFormat = DataFormatType.String,
						titles = new List<string> { "MdlInt." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "GrievanceDate", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date of Grievance" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "GrievanceType", DataFormat = DataFormatType.String,
						titles = new List<string> { "(1 - 11) Type of Grievance" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "AppealDate", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"Date ofAppeal",
							"Date of Appeal",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "AppealAction", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"(1 - 6) AppealAction",
							"(1 - 6) Appeal Action",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "DispositionDate", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"Date ofDisposition",
							"Date of Disposition",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "DispositionNoticeDate", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date Disposition Notice Sent" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "DispositionType", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"(1 - 12) Type ofDisposition",
							"(1 - 11) Type of Dispostion",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "DispositionStatus", DataFormat = DataFormatType.String,
						titles = new List<string> { "Disposition Status R=Resolved P=Pending" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "ExpiditedRequest", DataFormat = DataFormatType.String,
						titles = new List<string> { "Expedited Request Y=yes N=No" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "FileType", DataFormat = DataFormatType.String,
						titles = new List<string> { "File Type: GM=Griev MMA AM=Appeal MMA GL=Griev LTC AL=Appeal LTC" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "Originator", DataFormat = DataFormatType.String,
						titles = new List<string> { "Originator 1=Enrollee2 = Provider" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 18, Name = "ProviderNum", DataFormat = DataFormatType.String,
						titles = new List<string> { "Plan's 9 digit Medicaid Provider #:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 19, Name = "MedicalProviderNbrs", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Medicaid Provider #:",
							"Medicaid Provider ID#:"
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 20, Name = "CalendarYr", DataFormat = DataFormatType.String,
						titles = new List<string> { "Calendar Year:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 21, Name = "PlanName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Plan Name:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 22, Name = "Month", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "January", "February", "March", "April", "May", "June",
													"July", "August", "September", "October", "November", "December" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 23, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 24, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
				},
			};

			var wsLayout_erfr = new WorkSheetLayout
			{
				Name = "Enrollee Roster and Facility Residence Report",
				OutputFileName = "Data_Extract_Enrollee_Roster_Facility_Residence",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.Both,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion
					{
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A1", ValueRef = "C1" },
							new CellLocation { TitleRef = "A2", ValueRef = "C2" },
							new CellLocation { TitleRef = "A3", ValueRef = "C3" }
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "C2" },
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" }
						}
					},
					new CellLayoutVersion
					{
						Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "B3" },
							new CellLocation { TitleRef = "A4", ValueRef = "B4" },
							new CellLocation { TitleRef = "A5", ValueRef = "B5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 4,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "C5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 5,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "D5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 6,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "D3" },
							new CellLocation { TitleRef = "A4", ValueRef = "D4" },
							new CellLocation { TitleRef = "A5", ValueRef = "D5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 7,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined }
						}
					}
				},

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A7" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B7" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C7" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D7" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E7" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F7" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G7" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H7" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I7" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J7" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K7" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L7" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M7" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N7" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O7" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P7" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "Q7" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "R7" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "S7" } },
						},
						FirstRow = 8
					},
					new ColumnLayoutVersion
					{
						Version = 2,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A7" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B7" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C7" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D7" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E7" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F7" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G7" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H7" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I7" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J7" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K7" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L7" } },
						},
						FirstRow = 8
					},
					new ColumnLayoutVersion
					{
						Version = 3,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A6" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B6" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C6" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D6" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E6" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F6" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G6" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H6" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I6" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J6" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K6" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L6" } },
						},
						FirstRow = 7
					},
					new ColumnLayoutVersion
					{
						Version = 4,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A6" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B6" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D6" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F6" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G6" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I6" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J6" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L6" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N6" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O6" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "Q6" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "S6" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "T6" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "V6" } },
							new ColumnTitleLocation { col = 24, cellRefs = new List<string> { "X6" } },
							new ColumnTitleLocation { col = 25, cellRefs = new List<string> { "Y6" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "AA6" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "AB6" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "AD6" } },
						},
						FirstRow = 7
					},
					new ColumnLayoutVersion
					{
						Version = 5,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A5" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B5" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C5" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D5" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E5" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F5" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G5" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H5" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I5" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J5" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K5" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L5" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M5" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N5" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O5" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P5" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "Q5" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "R5" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "S5" } },
						},
						FirstRow = 6
					},
					new ColumnLayoutVersion
					{
						Version = 6,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A7" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B7" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C7" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D7" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E7" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F7" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G7" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H7" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I7" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J7" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K7" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L7" } },
						},
						FirstRow = 8
					},
					new ColumnLayoutVersion
					{
						Version = 7,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B8" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C8" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D8" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E8" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F8" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G8" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I8" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J8" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K8" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L8" } },
						},
						FirstRow = 9
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "LastName", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Enrollee Last Name",
							"Enrolee Last Name",
							"Last Name"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "FirstName", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Enrollee First Name",
							"Enrolee First Name",
							"First Name"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "MedicaidID", isRequired = true,
						DataFormat = DataFormatType.String, postProcRegex = new List<Tuple<string, string>> { new Tuple<string, string>(@"[^0-9]", "") },
						titles = new List<string> { "Medicaid ID" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "SSN",
						DataFormat = DataFormatType.String, postProcRegex = new List<Tuple<string, string>> { new Tuple<string, string>(@"[^0-9]", "") },
						titles = new List<string> { "Social Security Number" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "DOB", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"Date of Birth (mm/dd/yyyy)",
							"Date of Birth (MM/DD/YYYY)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "Addr", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Physical Address",
							"Physical Address (full street address)",
							"Address",
							"Physical Address (Full Street Address)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "City", DataFormat = DataFormatType.String,
						titles = new List<string> { "City" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "Zip", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Zip Code",
							"Zip code"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "County", DataFormat = DataFormatType.String,
						titles = new List<string> { "County of Residence" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "Region", DataFormat = DataFormatType.String,
						titles = new List<string> { "Region" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "FacilityType", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"County of ResidenceResidential Setting Type (Home, ALF, SNF or AFCH)",
							"County of ResidenceType of Facility",
							"Type of Facility",
							"Residential Setting Type (Home, ALF, SNF or AFCH)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "FacilityName", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Name of Facility",
							"Name of Facility (if applicable)",
							"Name of the Facility (if applicable)",
							"Name of the Facility",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "FacilityLic",
						DataFormat = DataFormatType.String, postProcRegex = new List<Tuple<string, string>> { new Tuple<string, string>(@"[^0-9]", "") },
						titles = new List<string>
						{
							"Facility License Number",
							"Facility License Number (if applicable)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "Tansition", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Identify if transitioning into a SNF or back into Community (SNF, Community, or N/A)",
							"Identify if transitioning into a SNF or back into Community (snf, Community, or N/A)",
							"Residential Setting Type (Home, ALF, SNF or AFCH",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "TransistionDate", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"Date of transition to SNF or Community (if applicable)",
							"Date of transition to SNF or Community",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Form2515Date", DataFormat = DataFormatType.DateMixed,
						titles = new List<string> { "Date the 2515 form was sent to DCF if transitioning (if applicable)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "canLocate", DataFormat = DataFormatType.String,
						titles = new List<string> { "Able to Locate? Y/N" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 18, Name = "canContact", DataFormat = DataFormatType.String,
						titles = new List<string> { "Able to Contact? Y/N" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 19, Name = "LastContactDate", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"If unable to contact or locate enrolee, date of last contact? (N/A if not applicable)",
							"If unable to contact or locate enrollee, date of last contact? (N/A if not applicable)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 20, Name = "Comments", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Comments including demonstration of attempts to contact enrolee if applicable",
							"Comments including demonstration of attempts to contact enrollee if applicable"
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 21, Name = "MC_PlanName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Managed Care Plan Name :",
							"Managed Care Plan Name:",
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 22, Name = "MC_PlanID", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Managed Care Plan ID:",
							"Managed Care Plan ID :"
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 23, Name = "Month", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string>
						{
							"Reporting Month (MM/DD/YYYY):",
							"Reporting Month:",
							"Reporting Month"
						}
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 24, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 25, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
				},
			};

			var wsLayout_mcms = new WorkSheetLayout
			{
				Name = "Missed Services Report (0131)",
				OutputFileName = "Data_Extract_Missed_Services",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.Both,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion
					{
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "A2", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined }
						}
					},
					new CellLayoutVersion
					{
						Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "A2" },
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined }
						}
					},
					new CellLayoutVersion
					{
						Version = 4,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "B3" },
							new CellLocation { TitleRef = "A4", ValueRef = "B4" }
						}
					},
					new CellLayoutVersion
					{
						Version = 5,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "C5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 6,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "B3" },
							new CellLocation { TitleRef = "A4", ValueRef = "B4" },
							new CellLocation { TitleRef = "A5", ValueRef = "B5" },
						}
					},
					new CellLayoutVersion
					{
						Version = 7,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined },
						}
					},
					new CellLayoutVersion
					{
						Version = 8,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "C5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 9,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A4", ValueRef = "B4" },
							new CellLocation { TitleRef = "A5", ValueRef = "B5" },
						}
					},
				},

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A5" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B5" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C5" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D5" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E5" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F5" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G5" } },
						},
						FirstRow = 6
					},
					new ColumnLayoutVersion
					{
						Version = 2,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A7" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B7" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C7" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D7" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E7" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F7" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G7" } },
						},
						FirstRow = 8
					},
					new ColumnLayoutVersion
					{
						Version = 3,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A7" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B7" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C7" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D7" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E7" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F7" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G7" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H7" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I7" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J7" } },
						},
						FirstRow = 8
					},
					new ColumnLayoutVersion
					{
						Version = 4,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A6" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B6" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C6" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D6" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E6" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F6" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G6" } },
						},
						FirstRow = 7
					},
					new ColumnLayoutVersion
					{
						Version = 5,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A4" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B4" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C4" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D4" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E4" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F4" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G4" } },
						},
						FirstRow = 5
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "LastName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Enrollee Last Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "FirstName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Enrollee First Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "MedicaidID", isRequired = true, DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Enrollee Medicaid ID",
							"Medicaid ID"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "ProviderName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Provider Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "Authorized_Service", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Authorized Service Type",
							"Authorization Service Type",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "Authorized_Units", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Authorized Service Units For The Reported Month",
							"Authorized Services Units for the Reported Month",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "Units_Missed", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Number of Missed Service Units In The Reported Month",
							"Number of Missed Services Units in the Reported Month",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "MissedServiceCode", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Reason for Missed Service (Enter Code)",
							"Reason for Missed Services (Enter Code)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "MissedServiceDate", DataFormat = DataFormatType.DateMixed,
						titles = new List<string>
						{
							"Date of Missed Service (XX/XX/XXXX)",
							"Date of Missed Service or Date Range if Multiple Dates Missed (XX/XX/XXXX)",
							"Date of Missed Services or Date Range if Multiple Dates Missed (XX/XX/XXXXX)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "Explanation", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Explanation and Resolution of Missed Services",
							"Resolution of Missed Service /Comments",
							"Resolution of Missed Services / Comments",
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 11, Name = "MC_PlanName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Managed Care Plan Name:",
							"Managed Care Plan Name :",
							"Coventry Health Plan Inc",
							"Humana Health Plan, Inc",
							"Sunshine State Health Plan",
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 12, Name = "MC_PlanID", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Managed Care Plan ID:",
							"Managed Care Plan ID :",
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 13, Name = "Month", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string>
						{
							"Reporting Month (mm/yyyy):",
							"Reporting Month (MM/DD/YYYY):",
							"Reporting Month:",
							"Reporting Month",
						}
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 14, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 15, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
				},
			};

			var wsLayout_nftr_to_comm = new WorkSheetLayout
			{
				Name = "Nursing Facility Transfer Report (0135) To Community",
				OutputFileName = "Data_Extract_Nursing_Facility_Transfer_To_Community",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.Both,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion
					{
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "B4" },
							new CellLocation { TitleRef = "A5", ValueRef = "B5" },
							new CellLocation { TitleRef = "A6", ValueRef = "B6" },
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "C5" },
							new CellLocation { TitleRef = "A6", ValueRef = "C6" },
						}
					},
					new CellLayoutVersion
					{
						Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A6", ValueRef = "A6", dataLayout = CellDataLayout.combined },
						}
					},
					new CellLayoutVersion
					{
						Version = 4,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A6", ValueRef = "C6" },
						}
					},
					new CellLayoutVersion
					{
						Version = 5,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A6", ValueRef = "B6" },
						}
					},
				},

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B8" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C8" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D8" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E8" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F8" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G8" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I8" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J8" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K8" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L8" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M8" } },
						},
						FirstRow = 9
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "LastName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Last Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "FirstName", DataFormat = DataFormatType.String,
						titles = new List<string> { "First Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "MedicaidID", isRequired = true, DataFormat = DataFormatType.String,
						titles = new List<string> { "Medicaid ID" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "SSN", DataFormat = DataFormatType.String,
						postProcRegex = new List<Tuple<string, string>> { new Tuple<string, string>(@"[^0-9]", "") },
						titles = new List<string> { "Social Security Number" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "DOB", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date of Birth (mm/dd/yyyy)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "County_Before", DataFormat = DataFormatType.String,
						titles = new List<string> { "County of Residence Pre-Community Placement" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "Enrollment_Date", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Effective Date of Enrollment with Managed Care Plan (mm/dd/yyyy)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "Admit_Date", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date Enrollee Admitted to Nursing Facility (mm/dd/yyyy)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "FacilityName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Name of Nursing Facility" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "ProviderNumber", DataFormat = DataFormatType.String,
						postProcRegex = new List<Tuple<string,string>> { new Tuple<string, string>(@"[^0-9]", "") },
						titles = new List<string> { "Nursing Facility Medicaid Provider Number" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "Transition_Date", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date Enrollee Transitioned to Community (mm/dd/yyyy)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "Residence_Addr", DataFormat = DataFormatType.String,
						titles = new List<string> { "Community Residence (ALF, AFCH, or Enrollee Residence) Include street address, name and license number of residence (if applicable)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "Post_Placement_County", DataFormat = DataFormatType.String,
						titles = new List<string> { "County of Residence Post-Community Placement" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 14, Name = "MC_PlanName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Managed Care Plan Name:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 15, Name = "MC_PlanID", DataFormat = DataFormatType.String,
						titles = new List<string> { "Managed Care Plan ID:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 16, Name = "Month", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string>
						{
							"Reporting Month (XX/XXXX):",
							"Reporting Month (MM/YYYY):",
							"Reporting Month:",
						}
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 17, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 18, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
				},
			};


			var wsLayout_nftr_to_nh = new WorkSheetLayout
			{
				Name = "Nursing Facility Transfer Report from Community (0135)",
				OutputFileName = "Data_Extract_Nursing_Facility_Transfer_To_NH",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.Both,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion
					{
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "B4" },
							new CellLocation { TitleRef = "A5", ValueRef = "B5" },
							new CellLocation { TitleRef = "A6", ValueRef = "B6" },
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A6", ValueRef = "A6", dataLayout = CellDataLayout.combined },
						}
					},
					new CellLayoutVersion
					{
						Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "C5" },
							new CellLocation { TitleRef = "A6", ValueRef = "C6" },
						}
					},
					new CellLayoutVersion
					{
						Version = 4,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A6", ValueRef = "A6", dataLayout = CellDataLayout.combined },
						}
					},
					new CellLayoutVersion
					{
						Version = 5,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A6", ValueRef = "C6" },
						}
					},
					new CellLayoutVersion
					{
						Version = 6,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A6", ValueRef = "B6" },
						}
					},
					new CellLayoutVersion
					{
						Version = 7,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "D4" },
							new CellLocation { TitleRef = "A5", ValueRef = "D5" },
							new CellLocation { TitleRef = "A6", ValueRef = "D6" },
						}
					},
					new CellLayoutVersion
					{
						Version = 8,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "B4", dataLayout = CellDataLayout.separate },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A6", ValueRef = "B6", dataLayout = CellDataLayout.separate }
						}
					}
				},

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B8" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C8" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D8" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E8" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F8" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G8" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I8" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J8" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K8" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L8" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M8" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N8" } },
						},
						FirstRow = 9
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "LastName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Last Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "FirstName", DataFormat = DataFormatType.String,
						titles = new List<string> { "First Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "MedicaidID", isRequired = true, DataFormat = DataFormatType.String,
						titles = new List<string> { "Medicaid ID" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "SSN", DataFormat = DataFormatType.String,
						postProcRegex = new List<Tuple<string,string>> { new Tuple<string, string>(@"[^0-9]", "") },
						titles = new List<string> { "Social Security Number" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "DOB", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date of Birth (mm/dd/yyyy)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "County_Before", DataFormat = DataFormatType.String,
						titles = new List<string> { "County of Residence Pre-Nursing Facility Placement" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "Enrollment_Date", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Effective Date of Enrollment with Managed Care Plan (mm/dd/yyyy)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "FacilityName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Name of Nursing Facility" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "County_After", DataFormat = DataFormatType.String,
						titles = new List<string> { "County of Residence Post-Nursing Facility Placement" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "ProviderNumber", DataFormat = DataFormatType.String,
						postProcRegex = new List<Tuple<string,string>> { new Tuple<string, string>(@"[^0-9]", "") },
						titles = new List<string> { "Nursing Facility Medicaid Provider Number" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "Admit_Date", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date Enrollee Admitted to Nursing Facility (mm/dd/yyyy)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "Residence_Addr_Prior", DataFormat = DataFormatType.String,
						titles = new List<string> { "Community Residence Prior to Nursing Facility (ALF, AFCH, or Enrollee Residence) Include street address, name and license number of residence (if applicable)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "isFrom_NH", DataFormat = DataFormatType.String,
						titles = new List<string> { "Was the Enrollee Previously Transitioned Into the Community From a Nursing Home? (Y/N)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "Prev_Transition_Date", DataFormat = DataFormatType.Date,
						titles = new List<string> { "If Yes, Date of Previous Transition (mm/dd/yyyy)" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 15, Name = "MC_PlanName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Managed Care Plan Name:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 16, Name = "MC_PlanID", DataFormat = DataFormatType.String,
						titles = new List<string> { "Managed Care Plan ID:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 17, Name = "Month", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string>
						{
							"Reporting Month (XX/XXXX):",
							"Reporting Month (XX/XXX):",
							"Reporting Month (MM/YYYY):",
							"Reporting Month:",
						}
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 18, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 19, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
				},
			};


			var wsLayout_co = new WorkSheetLayout
			{
				Name = "Community Outreach Report (0113)",
				OutputFileName = "Data_Extract_Community_Outreach",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.Both,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion
					{
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A1", ValueRef = "D1"  },
							new CellLocation { TitleRef = "A2", ValueRef = "A2", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "C2", ValueRef = "E2" }
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A1", ValueRef = "A1"  },
							new CellLocation { TitleRef = "A2", ValueRef = "A2", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "C2", ValueRef = "E2" }
						}
					},
				},

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A5" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B5" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C5" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D5" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E5" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F5" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G5" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H5" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I5" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J5" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K5" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L4" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M4" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N4" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O5" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P5" } },
						},
						FirstRow = 6
					}
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "Action_Taken", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "ACTION TAKEN (N=New, A=Amended)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "Event_Type", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "EVENT TYPE (P or H)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "Event_Name", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "EVENT NAME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "Start_Date", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string> { "START DATE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "End_Date", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string> { "END DATE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "Start_End_Time", DataFormat = DataFormatType.String,
						titles = new List<string> { "START/END TIME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "Sponsor_Name", DataFormat = DataFormatType.String,
						titles = new List<string> { "EVENT SPONSOR NAME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "Sponsor_Type", DataFormat = DataFormatType.String,
						titles = new List<string> { "EVENT SPONSOR TYPE OF ORGANIZATION" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "Event_Addr", DataFormat = DataFormatType.String,
						titles = new List<string> { "PHYSICAL LOCATION OF EVENT (STREET ADDRESS)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "Event_CSZ_County", DataFormat = DataFormatType.String,
						titles = new List<string> { "CITY AND COUNTY" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "Event_Contact", DataFormat = DataFormatType.String,
						titles = new List<string> { "EVENT CONTACT NAME:" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "Flier_Attached", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"INVITATION LETTER/FLYER ATTACHED: Yes/No",
							"INVITATION LETTER ATTACHED: Yes/No",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "Representatives", DataFormat = DataFormatType.String,
						titles = new List<string> { "NAMES OF PARTICIPATING OUTREACH REPRESENTATIVES" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "Service_Type", DataFormat = DataFormatType.String,
						titles = new List<string> { "TYPE OF HEALTH RELATED SERVICE(S) TO BE PROVIDED BY OUTREACH REPRESENTATIVE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "Promo_Items_LT_5dol", DataFormat = DataFormatType.String,
						titles = new List<string> { "PROMO ITEMS < $5.00 Yes/No" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Outreach_Material", DataFormat = DataFormatType.String,
						titles = new List<string> { "OUTREACH MATERIAL PROVIDED Yes/No" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 17, Name = "MC_PlanName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "SMMC-LTC PROGRAM" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 18, Name = "Month", DataFormat = DataFormatType.DateMixed,
						titles = new List<string> { "Reporting Month:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 19, Name = "Yr", DataFormat = DataFormatType.String,
						titles = new List<string> { "2014", "2015", "2016" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 20, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 21, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
				}
			};

			var wsLayout_co_info = new WorkSheetLayout
			{
				Name = "Community Outreach Report (0113) Info Sheet",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.CellOnly,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion
					{
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "B3", ValueRef = "C3" },
							new CellLocation { TitleRef = "B11", ValueRef = "C11" },
							new CellLocation { TitleRef = "C8", ValueRef = "C9" },
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "B4", ValueRef = "C4" },
							new CellLocation { TitleRef = "B12", ValueRef = "C12" },
							new CellLocation { TitleRef = "C9", ValueRef = "C10" },
						}
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.cell, OutputOrder = 1, Name = "MC_PlanName", isRequired = true, DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Plan Name:",
							"Managed Care Plan Name:",
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 2, Name = "Month", isRequired = true, DataFormat = DataFormatType.DateMixed,
						titles = new List<string> { "Reporting Month:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 3, Name = "Yr", isRequired = true, DataFormat = DataFormatType.String,
						titles = new List<string> { "Calendar Year:" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 4, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 5, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
				}
			};

			var wsLayout_co_Event = new WorkSheetLayout
			{
				Name = "Community Outreach Report (0113) Event_Info",
				OutputFileName = "Data_Extract_Community_Outreach",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.ColumnOnly,
				dst = dst,

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A5" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B5" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C5" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D5" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E5" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F5" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G5" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H5" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I5" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J5" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K5" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L4" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M4" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N4" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O5" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P5" } },
						},
						FirstRow = 6
					}
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "Action_Taken", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "ACTION TAKEN (N=New, A=Amended)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "Event_Type", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "EVENT TYPE (P or H)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "Event_Name", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "EVENT NAME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "Start_Date", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string> { "START DATE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "End_Date", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string> { "END DATE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "Start_End_Time", DataFormat = DataFormatType.String,
						titles = new List<string> { "START/END TIME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "Sponsor_Name", DataFormat = DataFormatType.String,
						titles = new List<string> { "EVENT SPONSOR NAME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "Sponsor_Type", DataFormat = DataFormatType.String,
						titles = new List<string> { "EVENT SPONSOR TYPE OF ORGANIZATION" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "Event_Addr", DataFormat = DataFormatType.String,
						titles = new List<string> { "PHYSICAL LOCATION OF EVENT (STREET ADDRESS)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "Event_CSZ_County", DataFormat = DataFormatType.String,
						titles = new List<string> { "CITY AND COUNTY" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "Event_Contact", DataFormat = DataFormatType.String,
						titles = new List<string> { "EVENT CONTACT NAME:" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "Flier_Attached", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"INVITATION LETTER/FLYER ATTACHED: Yes/No",
							"INVITATION LETTER ATTACHED: Yes/No",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "Representatives", DataFormat = DataFormatType.String,
						titles = new List<string> { "NAMES OF PARTICIPATING OUTREACH REPRESENTATIVES" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "Service_Type", DataFormat = DataFormatType.String,
						titles = new List<string> { "TYPE OF HEALTH RELATED SERVICE(S) TO BE PROVIDED BY OUTREACH REPRESENTATIVE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "Promo_Items_LT_5dol", DataFormat = DataFormatType.String,
						titles = new List<string> { "PROMO ITEMS < $5.00 Yes/No" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Outreach_Material", DataFormat = DataFormatType.String,
						titles = new List<string> { "OUTREACH MATERIAL PROVIDED Yes/No" }
					},
				}
			};

			var wsLayout_cor_jurat = new WorkSheetLayout
			{
				Name = "Community Outreach Representative Registration (0116) Jurat",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.CellOnly,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion
					{
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A5", ValueRef = "A6" },
							new CellLocation { TitleRef = "A23", ValueRef = "A23", dataLayout = CellDataLayout.combined },
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A5", ValueRef = "A6" },
							new CellLocation { TitleRef = "A23", ValueRef = "A24" },
						}
					},
					new CellLayoutVersion
					{
						Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A5" },
							new CellLocation { TitleRef = "A22", ValueRef = "A23" },
						}
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.cell, OutputOrder = 1, Name = "MC_PlanName", isRequired = true, DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Managed Care Plan Name",
							"Plan Name",
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 2, Name = "Month", DataFormat = DataFormatType.DateMixed,
						titles = new List<string> { "For the quarter ending:" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 3, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 4, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
				}
			};

			var wsLayout_cor_activity = new WorkSheetLayout
			{
				Name = "Community Outreach Representative Registration (0116) Activity",
				OutputFileName = "Data_Extract_Community_Outreach_Representative",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.ColumnOnly,
				dst = dst,

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A4" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B4" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C4" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D4" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E4" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F4" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G4" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H4" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I4" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J4" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K4" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L4" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M4" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N4" } },
						},
						FirstRow = 5
					},
					new ColumnLayoutVersion
					{
						Version = 2,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A3" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B3" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C3" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D3" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E3" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F3" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G3" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H3" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I3" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J3" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K3" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L3" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M3" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N3" } },
						},
						FirstRow = 4
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "Action", DataFormat = DataFormatType.String,
						titles = new List<string> { "ACTION R = Lisc/Certificate Renewed C = Info Updated N = New Representative T = Terminated" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "Action_Date", DataFormat = DataFormatType.Date,
						titles = new List<string> { "DATE OF ACTION TAKEN" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "Last_Name", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "LAST NAME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "First_Name", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "FIRST NAME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "Addr", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "ADDRESS" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "City", DataFormat = DataFormatType.String,
						titles = new List<string> { "CITY" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "Licence_Cert_Num", DataFormat = DataFormatType.String,
						titles = new List<string> { "LIC / CERT NUMBER (use NA if not applicable)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "Licence_Cert_Issue_Date", DataFormat = DataFormatType.DateMixed,
						titles = new List<string> { "LIC / CERT ISSUE DATE (use NA if not applicable)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "Licence_Cert_End_Date", DataFormat = DataFormatType.DateMixed,
						titles = new List<string> { "LIC / CERT END DATE (use NA if not applicable)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "Licence_Cert_Issuer", DataFormat = DataFormatType.String,
						titles = new List<string> { "LIC / CERT ISSUED BY: (DOH, DPR, DFS, ect.,use NA if not applicable)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "License_Cert_Type", DataFormat = DataFormatType.String,
						titles = new List<string> { "LIC /CERT TYPE OR DESIGNATION" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "Phone_Office", DataFormat = DataFormatType.String,
						titles = new List<string> { "COMM OUTR REP OFFICE TELEPHONE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "Phone_Cell", DataFormat = DataFormatType.String,
						titles = new List<string> { "COMM OUTR REP CELLULAR TELEPHONE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "Prev_Employer", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"PREVIOUS EMPLOYER",
							"PREVIOUS HEALTH PLAN EMPLOYER",
						}
					},
				}
			};

			var wsLayout_me_info = new WorkSheetLayout
			{
				Name = "Marketing/Public/Educational Events Report (0159) Info",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.CellOnly,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion
					{
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "B3", ValueRef = "C3" },
							new CellLocation { TitleRef = "B9", ValueRef = "C9" },
							new CellLocation { TitleRef = "B11", ValueRef = "C11" }
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "B3", ValueRef = "C3" },
							new CellLocation { TitleRef = "C8", ValueRef = "C9" },
							new CellLocation { TitleRef = "B11", ValueRef = "C11" }
						}
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.cell, OutputOrder = 1, Name = "MC_PlanName", isRequired = true, DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Managed Care Plan Name:",
							"Plan Name:"
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 2, Name = "Month", DataFormat = DataFormatType.String,
						titles = new List<string> { "Reporting Month:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 3, Name = "Yr", DataFormat = DataFormatType.String,
						titles = new List<string> { "Calendar Year:" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 4, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 5, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
				}
			};

			var wsLayout_me_events = new WorkSheetLayout
			{
				Name = "Marketing/Public/Educational Events Report(0159) Events",
				OutputFileName = "Data_Extract_Marketing_Event",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.ColumnOnly,
				dst = dst,

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A6" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B6" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C6" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D6" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E6" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F6" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G6" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H6" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I6" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J6" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K6" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L6" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M6" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N6" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O6" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P6" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "Q6" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "R6" } },
						},
						FirstRow = 7
					},
					new ColumnLayoutVersion
					{
						Version = 2,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A5" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B5" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C5" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D5" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E5" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F5" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G5" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H5" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I5" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J5" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K5" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L5" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M5" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N5" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O4" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P4" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "Q5" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "R5" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "S5" } }
						},
						FirstRow = 6
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "Action", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"ACTION TAKEN N=New; A=Amended; or C=Canceled",
							"ACTION TAKEN (N=New, A=Amended)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "Event_Type", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"EVENT TYPE M=Marketing; P=Public; or E=Educational (Please see Tab 1's Definitions)",
							"EVENT TYPE:",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "Event_Name", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"EVENT NAME (Please see Tab 1's Instructions)",
							"EVENT NAME"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "Plan_Type", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"PLAN TYPE (MMA, LTC)",
							"PLAN TYPE (MMA, LTC, COMP)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "Event_Formality", DataFormat = DataFormatType.String,
						titles = new List<string> { "TYPE OF MARKETING EVENT (if applicable) FE=Formal Event or IE=Informal Event (Please see Tab 1's Definitions)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "Start_Date", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string> { "START DATE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "End_Date", DataFormat = DataFormatType.Date,
						titles = new List<string> { "END DATE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "Start_Time", DataFormat = DataFormatType.String,
						titles = new List<string> { "START TIME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "End_Time", DataFormat = DataFormatType.String,
						titles = new List<string> { "END TIME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "Sponsor", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"EVENT SPONSOR NAME",
							"EVENT SPONSOR",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "Event_Addr", DataFormat = DataFormatType.String,
						titles = new List<string> { "PHYSICAL LOCATION OF EVENT (Street Address)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "Event_City_County", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"EVENT CITY AND COUNTY",
							"CITY AND COUNTY",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "Event_Region", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"EVENT REGION",
							"REGION",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "Contact", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"PLAN'S LEAD CONTACT NAME AND PHONE NUMBER:",
							"EVENT CONTACT NAME AND PHONE NUMBER:",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "Contact_Phone", DataFormat = DataFormatType.String,
						locType = LocateType.byRelated, RelatedCol = 14
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Notice_Submitted", DataFormat = DataFormatType.String,
						titles = new List<string> { "INVITATION NOTICE SUBMITTED: Yes/No" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "Event_Representative", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"NAME(S) OF PLAN'S PARTICIPATING MARKETING AGENT(S) OR EVENT REPRESENATIVE(S)",
							"NAMES OF PARTICIPATING PLAN MARKETING REPRESENTATIVES/AGENTS",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 18, Name = "Representative_DFS_LicNum", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"DFS LICENSE NUMBER(S) OF PARTICIPATING PLAN MARKETING AGENT(S) (if applicable)",
							"PARTICIPATING MARKETING REPRESENTATIVE/AGENT DFS LICENSE NUMBER",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 19, Name = "Nominal_Items", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"FILE NAME(S) OF AGENCY APPROVED NOMINAL VALUE ITEM(S) < $15.00 TO BE DISTRIBUTED (if applicable)",
							"PROMO ITEMS < $15.00 RETAIL  Yes/No",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 20, Name = "Handout_Provided", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"FILE NAME(S) OF AGENCY APPROVED EVENT WRITTEN MATERIAL(S) TO BE DISTRIBUTED (if applicable)",
							"HANDOUT MATERIAL PROVIDED Yes/No",
						}
					}
				}
			};

			var wsLayout_pdo = new WorkSheetLayout
			{
				Name = "Monthly PDO Report(0137)",
				OutputFileName = "Data_Extract_PDO",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.Both,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion
					{
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "B2" },
							new CellLocation { TitleRef = "A3", ValueRef = "B3" },
							new CellLocation { TitleRef = "A4", ValueRef = "B4" }
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "C2" },
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" }
						}
					},
					new CellLayoutVersion
					{
						Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A1",  ValueRef = "A1", dataLayout = CellDataLayout.aggregate, aggregateCellCnt = 3, aggregateCellSeparator = @"\n",
								cellMaps = new List<AggregateFieldCellMap>
								{
									new AggregateFieldCellMap { aggregateIdx = 0, dataLayout = CellDataLayout.combined },
									new AggregateFieldCellMap { aggregateIdx = 2, dataLayout = CellDataLayout.lookup, lookupString = "month" },
								}
							},
						}
					},
					new CellLayoutVersion
					{
						Version = 4,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "B4", ValueRef = "B2" },
							new CellLocation { TitleRef = "B6", ValueRef = "B6", dataLayout = CellDataLayout.combined }
						}
					},
					new CellLayoutVersion
					{
						Version = 5,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "B3" },
							new CellLocation { TitleRef = "A4", ValueRef = "B4" },
							new CellLocation { TitleRef = "A5", ValueRef = "B5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 6,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A1",  ValueRef = "A1", dataLayout = CellDataLayout.aggregate, aggregateCellCnt = 3, aggregateCellSeparator = @"\n",
								cellMaps = new List<AggregateFieldCellMap>
								{
									new AggregateFieldCellMap { aggregateIdx = 0, dataLayout = CellDataLayout.lookup, lookupString = "Monthly PDO Report" },
									new AggregateFieldCellMap { aggregateIdx = 2, dataLayout = CellDataLayout.lookup, lookupString = "month" },
								}
							},
						}
					},
					new CellLayoutVersion
					{
						Version = 7,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "A1" },
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.aggregate, aggregateCellCnt = 2, aggregateCellSeparator = "TO",
								cellMaps = new List<AggregateFieldCellMap>
								{
									new AggregateFieldCellMap { aggregateIdx = 0, dataLayout = CellDataLayout.lookup, lookupString = "month" }
								}
							}
						}
					},
				},

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A2" }, isGroupData = true, dataLayout = CellDataLayout.combined },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B2" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C2" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D3" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E3" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F3" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G3" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H3" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I2" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J2" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K2" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L2" } },
						},
						FirstRow = 4
					},
					new ColumnLayoutVersion
					{
						Version = 2,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A5" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B5" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C5" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D6" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E6" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F6" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G6" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H6" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I5" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J5" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K5" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L5" } },
						},
						FirstRow = 7
					},
					new ColumnLayoutVersion
					{
						Version = 3,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B10" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C10" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D10" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F10" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G10" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I10" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K10" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L10" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N10" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O10" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "P10" } },
						},
						FirstRow = 11
					},
					new ColumnLayoutVersion
					{
						Version = 4,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A7" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B7" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C7" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D7" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E7" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F7" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G7" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H7" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I7" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J7" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K7" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L7" } },
						},
						FirstRow = 8
					},
					new ColumnLayoutVersion
					{
						Version = 5,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A7" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B7" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D7" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E7" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F7" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G7" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H7" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I7" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J7" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K7" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L7" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M7" } },
						},
						FirstRow = 8
					},
					new ColumnLayoutVersion
					{
						Version = 6,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A4" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B4" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C4" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D4" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E4" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F4" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G4" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H4" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I4" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J4" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K4" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L4" } },
						},
						FirstRow = 5
					},
					new ColumnLayoutVersion
					{
						Version = 7,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A5" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B5" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C5" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D5" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E5" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F5" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G5" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H5" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I5" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J5" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K5" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L5" } },
						},
						FirstRow = 6
					},
					new ColumnLayoutVersion
					{
						Version = 8,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A5" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B5" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C5" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D5" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E5" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F5" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G5" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H5" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I5" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J5" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K5" } },
						},
						FirstRow = 6
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "LastName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Participant Last Name" },
						ignore = new List<string>
						{
							"None",
							"PDO Services Enrollment Totals",
							"Adult Companion Care",
							"Attendant Care",
							"Homemaker Services",
							"Intermittent and Skilled Nursing",
							"Personal Care Services",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "FirstName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Participant First Name",
							"Participant First Name      Medicaid ID",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "MedicaidID", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Medicaid ID" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "hasAdultCompanion", DataFormat = DataFormatType.String,
						titles = new List<string> { "Adult Companion Care" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "hasAttendantCare", DataFormat = DataFormatType.String,
						titles = new List<string> { "Attendant Care" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "hasHomemakerServices", DataFormat = DataFormatType.String,
						titles = new List<string> { "Homemaker Services" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "hasNursing", DataFormat = DataFormatType.String,
						titles = new List<string> { "Intermittent and Skilled Nursing" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "hasPersonalCareServices", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Personal Care Services",
							"Personal Care",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "EnrollmentStatus", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Enrollment Status (Enrolled/Disenrolled)",
							"Enrollment Status(Enrolled/Disenrolled)",
							"Enrollment Status",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "EnrollmentDate", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"Enrollment Date (mm/dd/yyyy)",
							"Enrollment Date",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "DisenrollmentDate", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"Disenrollment Date (mm/dd/yyyy)",
							"Disenrollment Date",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "DisenrollmentReason", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Disenrollment Reason (Reasons can be found in the PDO Manual)",
							"Disenrollment Reason"
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 13, Name = "MC_PlanName", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"NAME OF MANAGED CARE PLAN:",
							"Monthly PDO Report",
							"Managed Care Plan Name:",
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 14, Name = "MC_PlanID", DataFormat = DataFormatType.String,
						titles = new List<string> { "Managed Care Plan ID:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 15, Name = "Date", DataFormat = DataFormatType.DateMixed,
						titles = new List<string> { "Reporting Month (MM/DD/YYYY):", "Month", "From" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Region", DataFormat = DataFormatType.String, rowType = RowType.GroupData,
						titles = new List<string> { "REGION" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 17, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 18, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
				}
			};

			var wsLayout_mccma = new WorkSheetLayout
			{
				Name = "Case Management File Audit Report (0102)",
				OutputFileName = "Data_Extract_Case_Management_Audit",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				layoutType = LayoutType.Both,
				dst = dst,

				cellLayouts = new List<CellLayoutVersion>
				{
					new CellLayoutVersion
					{
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "A2", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined }
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "F2", dataLayout = CellDataLayout.separate },
							new CellLocation { TitleRef = "A3", ValueRef = "C3", dataLayout = CellDataLayout.separate },
							new CellLocation { TitleRef = "A4", ValueRef = "C4", dataLayout = CellDataLayout.separate }
						}
					},
					new CellLayoutVersion
					{
						Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "E2", dataLayout = CellDataLayout.separate },
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined }
						}
					},
					new CellLayoutVersion
					{
						Version = 4,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "B2", dataLayout = CellDataLayout.separate },
							new CellLocation { TitleRef = "A3", ValueRef = "B3", dataLayout = CellDataLayout.separate },
							new CellLocation { TitleRef = "A4", ValueRef = "B4", dataLayout = CellDataLayout.separate }
						}
					}
				},

				verifyFirstRowData = true,

				colLayouts = new List<ColumnLayoutVersion>
				{
					new ColumnLayoutVersion
					{
						Version = 1,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "A6", "K6" } },

							// 8 -> 22, 25 -> 37, 40 -> 55
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },

							new ColumnTitleLocation { col = 25, cellRefs = new List<string> { "A25" } },
							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },

							new ColumnTitleLocation { col = 40, cellRefs = new List<string> { "A40" } },
							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "A41" } },
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
							new ColumnTitleLocation { col = 55, cellRefs = new List<string> { "A55" } },
							new ColumnTitleLocation { col = 56, cellRefs = new List<string> { "A56" } },
						},
						FirstRow = 11
					},
					new ColumnLayoutVersion
					{
						Version = 2,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "A6", "K6" } },

							// 8 -> 22, 25 -> 37, 40 -> 54
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },

							new ColumnTitleLocation { col = 25, cellRefs = new List<string> { "A25" } },
							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },

							new ColumnTitleLocation { col = 40, cellRefs = new List<string> { "A40" } },
							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "A41" } },
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
						},
						FirstRow = 11
					},
					new ColumnLayoutVersion
					{
						Version = 3,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "A6", "K6" } },

							// 8 -> 23, 26 -> 38, 41 -> 55
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },
							new ColumnTitleLocation { col = 23, cellRefs = new List<string> { "A23" } },

							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },
							new ColumnTitleLocation { col = 38, cellRefs = new List<string> { "A38" } },

							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "A41" } },
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
							new ColumnTitleLocation { col = 55, cellRefs = new List<string> { "A55" } },
						},
						FirstRow = 11
					},
					new ColumnLayoutVersion
					{
						Version = 4,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "A7" } },

							// 9 -> 24, 27 -> 39, 42 -> 56
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },
							new ColumnTitleLocation { col = 23, cellRefs = new List<string> { "A23" } },
							new ColumnTitleLocation { col = 24, cellRefs = new List<string> { "A24" } },

							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },
							new ColumnTitleLocation { col = 38, cellRefs = new List<string> { "A38" } },
							new ColumnTitleLocation { col = 39, cellRefs = new List<string> { "A39" } },

							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
							new ColumnTitleLocation { col = 55, cellRefs = new List<string> { "A55" } },
							new ColumnTitleLocation { col = 56, cellRefs = new List<string> { "A56" } }
						},
						FirstRow = 2
					},
					new ColumnLayoutVersion
					{
						Version = 5,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "A6", "L7" } },

							// 8 -> 25, 27 -> 39, 42 -> 52, 54 -> 56
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },
							new ColumnTitleLocation { col = 23, cellRefs = new List<string> { "A23" } },
							new ColumnTitleLocation { col = 24, cellRefs = new List<string> { "A24" } },
							new ColumnTitleLocation { col = 25, cellRefs = new List<string> { "A25" } },

							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },
							new ColumnTitleLocation { col = 38, cellRefs = new List<string> { "A38" } },
							new ColumnTitleLocation { col = 39, cellRefs = new List<string> { "A39" } },

							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },

							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
							new ColumnTitleLocation { col = 55, cellRefs = new List<string> { "A55" } },
							new ColumnTitleLocation { col = 56, cellRefs = new List<string> { "A56" } },
						},
						FirstRow = 13
					},
					new ColumnLayoutVersion
					{
						Version = 6,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "A6", "K5" } },

							// 8 -> 22, 25 -> 37, 40 -> 54 same as v2
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },

							new ColumnTitleLocation { col = 25, cellRefs = new List<string> { "A25" } },
							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },

							new ColumnTitleLocation { col = 40, cellRefs = new List<string> { "A40" } },
							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "A41" } },
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
						},
						FirstRow = 12
					},
					new ColumnLayoutVersion
					{
						Version = 7,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "A7", "B5" } },

							// 9 -> 23, 26 -> 38, 41 -> 55
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },
							new ColumnTitleLocation { col = 23, cellRefs = new List<string> { "A23" } },

							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },
							new ColumnTitleLocation { col = 38, cellRefs = new List<string> { "A38" } },

							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "A41" } },
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
							new ColumnTitleLocation { col = 55, cellRefs = new List<string> { "A55" } },
						},
						FirstRow = 2
					},
					new ColumnLayoutVersion
					{
						Version = 8,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "A6", "L5" } },

							// 8 -> 22, 25 -> 37, 40 -> 54 *
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },

							new ColumnTitleLocation { col = 25, cellRefs = new List<string> { "A25" } },
							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },

							new ColumnTitleLocation { col = 40, cellRefs = new List<string> { "A40" } },
							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "A41" } },
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
						},
						FirstRow = 12
					},
					new ColumnLayoutVersion
					{
						Version = 9, // copied from 4
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "A6" } },  // 7 -> 6

							// 8 -> 23, 26 -> 38, 41 -> 55
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },  // added
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },
							new ColumnTitleLocation { col = 23, cellRefs = new List<string> { "A23" } },

							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },	// 24 -> 26
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },
							new ColumnTitleLocation { col = 38, cellRefs = new List<string> { "A38" } },

							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "A41" } },	// 39 -> 41 
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
							new ColumnTitleLocation { col = 55, cellRefs = new List<string> { "A55" } },
//							new ColumnTitleLocation { col = 56, cellRefs = new List<string> { "A56" } }		// removed row
						},
						FirstRow = 2
					},
					new ColumnLayoutVersion
					{
						Version = 10,	// copied from v2
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "A6", "C6" } },	// 7 -> 6, K6 -> C6

							// 8 -> 22, 25 -> 37, 40 -> 54
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },

							new ColumnTitleLocation { col = 25, cellRefs = new List<string> { "A25" } },
							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },

							new ColumnTitleLocation { col = 40, cellRefs = new List<string> { "A40" } },
							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "A41" } },
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
						},
						FirstRow = 4	// 11 -> 4
					},
					new ColumnLayoutVersion
					{
						Version = 11,	// copied from v2
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "A6", "B6" } },	// col 7 -> 6, K6 -> B6

							// 8 -> 22, 25 -> 37, 40 -> 54
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },

							new ColumnTitleLocation { col = 25, cellRefs = new List<string> { "A25" } },
							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },

							new ColumnTitleLocation { col = 40, cellRefs = new List<string> { "A40" } },
							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "A41" } },
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
						},
						FirstRow = 3	// 11 -> 3
					},
					new ColumnLayoutVersion
					{
						Version = 12,	// copied from v2
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "A6", "H5" } },	// col 7 -> 6, K6 -> H5

							// 8 -> 22, 25 -> 37, 40 -> 54
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "A12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "A13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "A14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "A15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "A16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "A17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "A18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "A19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "A20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "A21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "A22" } },

							new ColumnTitleLocation { col = 25, cellRefs = new List<string> { "A25" } },
							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "A34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },

							new ColumnTitleLocation { col = 40, cellRefs = new List<string> { "A40" } },
							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "A41" } },
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "A42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "A43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "A44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "A45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "A46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "A47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "A48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "A49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "A50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "A51" } },
							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
						},
						FirstRow = 8	// 11 -> 8
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "EnrolleeID", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Enrollee ID NumberID #",
							"Enrollee ID NumberY or N",
							"Enrollee Medicaid ID NumberID #",
							"Enrollee Medicaid ID Number",
							"Enrollee Medicaid ID Number:Y/N or N/A",
							"Enrollee ID NumberY or N"
						},
						postProcRegex = new List<Tuple<string,string>>
						{
							new Tuple<string, string>("ID", ""),
							new Tuple<string, string>("Totals", ""),
							new Tuple<string, string>("Total", ""),
							new Tuple<string, string>("Compliance", ""),
							new Tuple<string, string>("Yes", ""),
							new Tuple<string, string>("No", ""),
							new Tuple<string, string>("N/A", ""),
							new Tuple<string, string>("[#%+&Y]", ""),
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "ContactWithin5Days", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Was the initial contact completed within 5 business days if in community? 7 days for nursing facility?",
							"Was the initial contact completed within 5 business days if in community? 7days for nursing facility?",
							"Was the initial contact completed within 5 business days if in community?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "NhContactWithin7Days", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Or within 7 business days if in a nursing facility?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "PhoneFollowupWithing7Days", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Did CM conduct a telephone follow-up call within 7 business days after initial assessment?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "ContactsDocumented", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Were all contacts to enrollee that were attempted or made, documented in the case notes?",
							"Were all contacts to enrollee documented in the case notes?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "PublicGuardianship", DataFormat = DataFormatType.String,
						titles = new List<string> { "Was the enrollee referred to the Public Guardianship Program if needed?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "ExplainedRights", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Did the CM explain the enrollee’s rights and responsibilities? Including grievance, appeal, and fair hearing info?",
							"Did the CM explain the enrollee's rights and responsibilities? Including grievance, appeal, and the fair hearing info?",
							"Did the CM explain the enrollee's rights and responsibilities?",
							"Did the CM explain the enrollee’s rights and responsibilities?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "ExplainMedicaidRights", DataFormat = DataFormatType.String,
						titles = new List<string> { "Did the CM explain the enrollee's Medicaid Fair Hearing rights?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "ExplainGrievance", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Did the CM explain the grievance, appeal and fair hearing information to the enrollee?",
							"Did the CM explain grievance, appeal, and fair hearing information to the enrollee?",
							"Did the CM explain grievance and appeal information to the enrollee?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "VisitsDocumented", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Has CM documented contacts and face to face visits in a timely fashion?",
							"Has CM documented contacts and face to face visits in a timely fashion? (per the timetable specified in the contract?)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "ServiceChangesDocumented", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Has CM documented changes in services in a timely fashion?",
							"Has CM documented changes in services in a timely fashion? (per the timetable specified in the contract?)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "EmergencyPlanOnFile", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Is the enrollee's personal emergency plan in the case file?",
							"Is the enrollee's persoN/Al emergency plan in the case file?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "RegisteredSpecialNeedsShelter", DataFormat = DataFormatType.String,
						titles = new List<string> { "Is the enrollee registered with a Special Needs Shelter?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "PCPIdentified", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Is the enrollee's PCP identified?",
							"Is the enrollee's primary care physician identified?"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "FileHasEligibilityDocs", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Were copies of eligibility documents included in case file? (ie-LOC determinations, etc)",
							"Were copies of eligibility documents included in case file? (ie-LOC determiN/Ations, etc)",
							"Were copies of eligibility documents included in case file? (i.e.-LOC determinations, etc.)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "FileHas701bAssessment", DataFormat = DataFormatType.String,
						titles = new List<string> { "Is the 701-B assessment present in the case file and completed properly?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "FileHasHighRiskScreening_Monitoring", DataFormat = DataFormatType.String,
						titles = new List<string> { "Was there evidence of special screening for and monitoring of high risk persons and conditions documented in the case file?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 18, Name = "FileHasProviderChoice", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Is there documentation of individual provider choice and Medicaid Fair Hearing information?",
							"Is there documentation of individual provider choice?",
						}
					},

					new Field { fldType = FieldType.column, OutputOrder = 19, Name = "EnrolleeHasCarePlanCopy", DataFormat = DataFormatType.String,
						titles = new List<string> { "Did the enrollee receive a copy of their current care plan?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 20, Name = "DocumentedRisks_Barriers", DataFormat = DataFormatType.String,
						titles = new List<string> { "Were risks and barriers documented in the risk assessment?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 21, Name = "DocumentedInterventions", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Are interventions documented in the care plan?",
							"Are interventions docmented in the care plan?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 22, Name = "DocumentedServiceSchedules", DataFormat = DataFormatType.String,
						titles = new List<string> { "Are service schedules documented in the care plan?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 23, Name = "DocumentedMedicationManagement", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Are medication management strategies documented in the care plan?",
							"Are medication maN/Agement strategies documented in the care plan?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 24, Name = "DocumentedProgress", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Is progress documented in the care plan or case file?",
							"Is progress of documented in the care plan or case file?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 25, Name = "CarePlanSentToPhysician", DataFormat = DataFormatType.String,
						titles = new List<string> { "Was the care plan sent to the enrollee's primary care physician?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 26, Name = "CarePlanSigned", DataFormat = DataFormatType.String,
						titles = new List<string> { "Is the care plan signed on the initial date of development, and every additional update?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 27, Name = "CarePlanSignedOnDevelopmentDate", DataFormat = DataFormatType.String,
						titles = new List<string> { "Is the care plan signed on the initial date of development" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 28, Name = "CarePlanSignedEveryUpdate", DataFormat = DataFormatType.String,
						titles = new List<string> { "Is the care plan signed for every care plan update?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 29, Name = "ConsistentServiceAuthorizations", DataFormat = DataFormatType.String,
						titles = new List<string> { "Are the service authorizations consistent with the plan of care?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 30, Name = "CarePlanUpdated", DataFormat = DataFormatType.String,
						titles = new List<string> { "If the enrollee's services have changed, does the care plan reflect these updates?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 31, Name = "FileHasCareSummary", DataFormat = DataFormatType.String,
						titles = new List<string> { "Is a plan of care summary included in the case file?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 32, Name = "EnrolleeHasCareSummary", DataFormat = DataFormatType.String,
						titles = new List<string> { "Did the enrollee or representative receive a copy of the plan of care summary?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 33, Name = "FileHasManagedDiagnosisDocumentation", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Is there documentation in the case file that the enrollee's diagnoses are being managed effectively?",
							"Is there documentation in the case file that the enrollee's diagnoses are being maN/Aged effectively?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 34, Name = "MonthlyPhoneVerification", DataFormat = DataFormatType.String,
						titles = new List<string> { "Were monthly telephone contacts made and documented to verify satisfaction and receipt of services?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 35, Name = "QtrlyVisitHome", DataFormat = DataFormatType.String,
						titles = new List<string> { "Were face to face visits made and documented every three months to evaluate and document the home characteristics?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 36, Name = "QtrlyVisitReview", DataFormat = DataFormatType.String,
						titles = new List<string> { "Was the care plan reviewed during face to face quarterly visit?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 37, Name = "YrlyVisitDocumented", DataFormat = DataFormatType.String,
						titles = new List<string> { "Was the annual face-to-face visit with enrollee documented and completed for annual reassessment?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 38, Name = "EnrolleeCareLevelCurrent", DataFormat = DataFormatType.String,
						titles = new List<string> { "Is the enrollee's level of care current?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 39, Name = "YrlyHandbookReview", DataFormat = DataFormatType.String,
						titles = new List<string> { "Has the CM documented reviewing the enrollee handbook with the enrollee/reps on a yearly basis?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 40, Name = "ContactProviderClNeeds", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Has the CM documented contacting the CL’s HCBS providers to discuss their assessment of the CL's needs?",
							"Has the CM documented contacting the CL's HCBS providers to discuss their assessment of the CL's needs?",
							"Has the CM documented contacting the enrollees HCBS providers to discuss their assessment of the enrollees' needs?",
							"Has the CM documented contacting the enrollee’s HCBS providers to discuss their assessment of the enrollee's needs?",
							"Has the CM documented contacting the enrollee’s HCBS providers to discuss their assessment of the enrollee's needs and status?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 41, Name = "EnrolleeHasOutsideReferrals", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Have referrals have been provided to the enrollee outside of the Managed Care Organization?",
							"Have referrals been provided to the enrollee outside of the Managed Care Organization?",
							"Have referrals been provided to the enrollee outside of the MaN/Aged Care Organization?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 42, Name = "EnrolleeNotifiedOnRejection", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"If the enrollee has a service request that is denied, reduced, terminated or suspended, were they notified in writing?",
							"If the enrollee has a service request that is denied, reduced, termiN/Ated or suspended, were they notified in writing?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 43, Name = "YrlyProviderContactEnrolleeNeeds", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Were the enrollee’s HCBS providers contacted at least annually to discuss their assessment of the CL’s needs and status?",
							"Were the enrollee's HCBS providers contacted at least annually to discuss their assessment of the CL's needs and status?",
							"Were the enrollee's HCBS providers contacted at least annually to discuss their assessment of the enrollee's needs and status?",
							"Were the enrollee’s HCBS providers contacted at least annually to discuss their assessment of the enrollee’s needs and status?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 44, Name = "AlfLikeHomeDocuments", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Was there documentation of home-like characteristic for enrollee's in ALFs.",
							"Was there documentation of home-like characteristic for enrollee's in ALF's.",
							"Was there documentation of home-like characteristic for enrollees in ALFs.",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 45, Name = "FileHasNeedsAssesmentsPhysicianReferrals", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Are needs assessments and physican referrals included in case file?",
							"Are needs assessments and physican referrals included in the case file?",
							"Are needs assessments and physician referrals included in the case file?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 46, Name = "FileHasCaseNarrativesOfContacts", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Are ongoing case narratives present in the case file, that document case management visits and other contacts?",
							"Are ongoing case N/Arratives present in the case file, that document case maN/Agement visits and other contacts?",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 47, Name = "FileHasSatisfactionSurveys", DataFormat = DataFormatType.String,
						titles = new List<string> { "Are satisfaction surveys present in the case file?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 48, Name = "QualityRemediationInCaseNotes", DataFormat = DataFormatType.String,
						titles = new List<string> { "Do the case notes document the review of complaints and the quality remediation to resolve and prevent problems?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 49, Name = "DocumentedQrtlyProgress", DataFormat = DataFormatType.String,
						titles = new List<string> { "Is progress documented at least quarterly on the members care plan for the person centered care plans?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 50, Name = "TimelyAnnualAssesment", DataFormat = DataFormatType.String,
						titles = new List<string> { "Was the annual assessment completed timely? (no more than 60 days before the LOC date and no less than 30 days before the LOC date)" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 51, Name = "MC_PlanName", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"LTC Managed Care Contractor:",
							"LTC Managed Care Organization:",
							"LTC MaN/Aged Care Organization:",
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 52, Name = "Reviewer", DataFormat = DataFormatType.String,
						titles = new List<string> { "Reviewer:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 53, Name = "Date", DataFormat = DataFormatType.String,
						titles = new List<string> { "Date:" }
					},
					new Field { fldType = FieldType.filePath, OutputOrder = 54, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.fileName, OutputOrder = 55, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true }
				}
			};

			// Create list of data source types.
			dst.types = new List<SpreadSheetLayout>
			{
				new SpreadSheetLayout
				{
					Name = "Enrollee Complaints, Grievances and Appeals Report (0127)",
					procType = ProcessType.MatchAllDataWorkSheets,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { Names = new List<string> { "Instructions" } },
						new SheetLayout { Names = new List<string> { "Codes" } },
						new SheetLayout { Names = new List<string> { "January" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "February" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "March" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "April" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "May" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "June" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "July" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "August" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "September" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "October" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "November" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "December" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "Summary" } },
						new SheetLayout { Names = new List<string> { "October 2014" }, sheetType = SheetType.SourceData, isOptional = true, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "November 2014" }, sheetType = SheetType.SourceData, isOptional = true, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "December 2014" }, sheetType = SheetType.SourceData, isOptional = true, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "Sheet1" }, isOptional = true }
					}
				},

				new SpreadSheetLayout
				{
					Name = "Enrollee Roster and Facility Residence Report (0129)",
					procType = ProcessType.MatchByClosestWorkSheetLayout,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { sheetType = SheetType.SourceData, wsLayout = wsLayout_erfr }
					}
				},

				new SpreadSheetLayout
				{
					Name = "Managed Care Missed Services Report",
					procType = ProcessType.MatchByClosestWorkSheetLayout,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { sheetType = SheetType.SourceData, wsLayout = wsLayout_mcms }
					}
				},

				new SpreadSheetLayout
				{
					Name = "Nursing Facility Transfer Report",
					procType = ProcessType.MatchAllDataWorkSheets,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout
						{
							Names = new List<string>
							{
								"Nursing Facility to Community",
								"Nursing Facility_to Community",
								"Nursing Faclity To Community",
							},
							sheetType = SheetType.SourceData,
							wsLayout = wsLayout_nftr_to_comm
						},
						new SheetLayout
						{
							Names = new List<string>
							{
								"Community to Nursing Facility",
								"Community_to_NursingFacility",
								"Community To Nursing Faclity",
							},
							sheetType = SheetType.SourceData,
							wsLayout = wsLayout_nftr_to_nh
						},
						new SheetLayout
						{
							Names = new List<string>
							{
								"Sheet1",
								"Sheet2",
								"Sheet3",
							},
							isOptional = true
						},
					}
				},

				new SpreadSheetLayout
				{
					Name = "Community Outreach v1",
					procType = ProcessType.MatchByClosestWorkSheetLayout,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { Names = new List<string> { "Event Info" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_co },
					}
				},

				new SpreadSheetLayout
				{
					Name = "Community Outreach v2",
					procType = ProcessType.MatchAllDataWorkSheets,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout
						{
							Names = new List<string>
							{
								"Info Sheet",
								"Plan Info Sheet"
							},
							sheetType = SheetType.CommonData,
							wsLayout = wsLayout_co_info
						},
						new SheetLayout { Names = new List<string> { "Event Info" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_co_Event },
					}
				},

				new SpreadSheetLayout
				{
					Name = "Community Outreach Representative",
					procType = ProcessType.MatchAllDataWorkSheets,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { Names = new List<string> { "Instructions" }, isOptional = true },
						new SheetLayout { Names = new List<string> { "Jurat" }, sheetType = SheetType.CommonData, wsLayout = wsLayout_cor_jurat },
						new SheetLayout { Names = new List<string> { "Representative Activity" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cor_activity },
						new SheetLayout { Names = new List<string> {  "Sheet1", "Sheet3" }, isOptional = true }
					}
				},

				new SpreadSheetLayout
				{
					Name = "Marketing/Public/Educational Events",
					procType = ProcessType.MatchAllDataWorkSheets,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { Names = new List<string> { "Instructions-Definitions" }, isOptional = true },
						new SheetLayout { Names = new List<string> { "Plan Info Sheet" }, sheetType = SheetType.CommonData, wsLayout = wsLayout_me_info },
						new SheetLayout { Names = new List<string> { "Monthly Events Report" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_me_events }
					}
				},

				new SpreadSheetLayout
				{
					Name = "Marketing/Public/Educational Events and Community Outreach",
					procType = ProcessType.MatchAllDataWorkSheets,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { Names = new List<string> { "Plan Info Sheet" }, sheetType = SheetType.CommonData, wsLayout = wsLayout_me_info },
						new SheetLayout { Names = new List<string> { "Marketing Events" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_me_events },
						new SheetLayout
						{
							Names = new List<string>
							{
								"Community Outreach Events",
								"Education Comm Outreach Events",
							},
							sheetType = SheetType.SourceData, wsLayout = wsLayout_co_Event
						},
						new SheetLayout { Names = new List<string> { "Sheet1" }, isOptional = true }
					}
				},

				new SpreadSheetLayout
				{
					Name = "PDO Report",
					procType = ProcessType.MatchByClosestWorkSheetLayout,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { sheetType = SheetType.SourceData, wsLayout = wsLayout_pdo }
					}
				},

				new SpreadSheetLayout
				{
					Name = "Managed Care Case Management File Audit",
					procType = ProcessType.MatchByClosestWorkSheetLayout,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { sheetType = SheetType.SourceData, wsLayout = wsLayout_mccma }
					}
				}

			};

			dst.types.ForEach(ssl => ssl.sLayouts.ForEach(sl => sl.ssLayout = ssl));

			return dst.types;
		}
	}
}
