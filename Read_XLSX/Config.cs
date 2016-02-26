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
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "Region", DataFormat = DataFormatType.String,
						titles = new List<string> { "Region # (1 - 11)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "County", DataFormat = DataFormatType.String,
						titles = new List<string> { "County Name Within Region:" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "MedicaidID", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Recipient's Medicaid ID#:",
							"Recipient's Medicaid ID#",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "LastName", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Recipient LastName:",
							"Recipient Last Name",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "FirstName", DataFormat = DataFormatType.String,
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
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "DispositionType", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"(1 - 12) Type ofDisposition",
							"(1 - 11) Type of Dispostion",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "DispositionStatus", DataFormat = DataFormatType.String,
						titles = new List<string> { "Disposition Status R=Resolved P=Pending" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "ExpiditedRequest", DataFormat = DataFormatType.String,
						titles = new List<string> { "Expedited Request Y=yes N=No" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "FileType", DataFormat = DataFormatType.String,
						titles = new List<string> { "File Type: GM=Griev MMA AM=Appeal MMA GL=Griev LTC AL=Appeal LTC" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Originator", DataFormat = DataFormatType.String,
						titles = new List<string> { "Originator 1=Enrollee2 = Provider" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "ProviderNum", DataFormat = DataFormatType.String,
						titles = new List<string> { "Plan's 9 digit Medicaid Provider #:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 18, Name = "MedicalProviderNbrs", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Medicaid Provider #:",
							"Medicaid Provider ID#:"
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 19, Name = "CalendarYr", DataFormat = DataFormatType.String,
						titles = new List<string> { "Calendar Year:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 20, Name = "PlanName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Plan Name:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 21, Name = "Month", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "January", "February", "March", "April", "May", "June",
													"July", "August", "September", "October", "November", "December" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 22, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
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
							new CellLocation { TitleRef = "A3", ValueRef = "A3", isCombined = true },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", isCombined = true },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", isCombined = true }
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
						DataFormat = DataFormatType.String, postProcRegex = new List<string> { @"[^0-9]", "" },
						titles = new List<string> { "Medicaid ID" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "SSN",
						DataFormat = DataFormatType.String, postProcRegex = new List<string> { @"[^0-9]", "" },
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
						DataFormat = DataFormatType.String, postProcRegex = new List<string> { @"[^0-9]", "" },
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
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Form2515Date", DataFormat = DataFormatType.String,
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
					new Field { fldType = FieldType.fileName, OutputOrder = 24, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
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
							new CellLocation { TitleRef = "A2", ValueRef = "A2", isCombined = true },
							new CellLocation { TitleRef = "A3", ValueRef = "A3", isCombined = true },
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "A3", isCombined = true },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", isCombined = true },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", isCombined = true }
						}
					},
					new CellLayoutVersion
					{
						Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "A2" },
							new CellLocation { TitleRef = "A3", ValueRef = "A3", isCombined = true }
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
							new CellLocation { TitleRef = "A4", ValueRef = "A4", isCombined = true },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", isCombined = true },
						}
					},
					new CellLayoutVersion
					{
						Version = 8,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "A3", isCombined = true },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "C5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 9,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "A3", isCombined = true },
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
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "MissedServiceDate", DataFormat = DataFormatType.String,
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
					new Field { fldType = FieldType.fileName, OutputOrder = 14, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
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
							new CellLocation { TitleRef = "A4", ValueRef = "A4", isCombined = true },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", isCombined = true },
							new CellLocation { TitleRef = "A6", ValueRef = "A6", isCombined = true },
						}
					},
					new CellLayoutVersion
					{
						Version = 4,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", isCombined = true },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", isCombined = true },
							new CellLocation { TitleRef = "A6", ValueRef = "C6" },
						}
					},
					new CellLayoutVersion
					{
						Version = 5,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", isCombined = true },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", isCombined = true },
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
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "SSN", DataFormat = DataFormatType.String, postProcRegex = new List<string> { @"[^0-9]", "" },
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
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "ProviderNumber", DataFormat = DataFormatType.String, postProcRegex = new List<string> { @"[^0-9]", "" },
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
					new Field { fldType = FieldType.fileName, OutputOrder = 17, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
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
							new CellLocation { TitleRef = "A4", ValueRef = "A4", isCombined = true },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", isCombined = true },
							new CellLocation { TitleRef = "A6", ValueRef = "A6", isCombined = true },
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
							new CellLocation { TitleRef = "A5", ValueRef = "A5", isCombined = true },
							new CellLocation { TitleRef = "A6", ValueRef = "A6", isCombined = true },
						}
					},
					new CellLayoutVersion
					{
						Version = 5,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", isCombined = true },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", isCombined = true },
							new CellLocation { TitleRef = "A6", ValueRef = "C6" },
						}
					},
					new CellLayoutVersion
					{
						Version = 6,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A4", ValueRef = "A4", isCombined = true },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", isCombined = true },
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
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "SSN", DataFormat = DataFormatType.String, postProcRegex = new List<string> { @"[^0-9]", "" },
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
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "ProviderNumber", DataFormat = DataFormatType.String, postProcRegex = new List<string> { @"[^0-9]", "" },
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
					new Field { fldType = FieldType.fileName, OutputOrder = 18, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
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
							new CellLocation { TitleRef = "A2", ValueRef = "A2", isCombined = true },
							new CellLocation { TitleRef = "C2", ValueRef = "E2" }
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A1", ValueRef = "A1"  },
							new CellLocation { TitleRef = "A2", ValueRef = "A2", isCombined = true },
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
					new Field { fldType = FieldType.cell, OutputOrder = 18, Name = "Yr", DataFormat = DataFormatType.String,
						titles = new List<string> { "2014", "2015", "2016" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 19, Name = "Month", DataFormat = DataFormatType.String,
						titles = new List<string> { "Reporting Month:" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 20, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
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
					new Field { fldType = FieldType.cell, OutputOrder = 2, Name = "Yr", isRequired = true, DataFormat = DataFormatType.String,
						titles = new List<string> { "Calendar Year:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 3, Name = "Month", isRequired = true, DataFormat = DataFormatType.String,
						titles = new List<string> { "Reporting Month:" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 5, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
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
								"Sheet2",
								"Sheet3"
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
				}
			};

			dst.types.ForEach(ssl => ssl.sLayouts.ForEach(sl => sl.ssLayout = ssl));

			return dst.types;
		}
	}
}
