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

		public SpreadSheetLayout lastSSL { get; set; }

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
							"Enrollee's Medicaid ID#"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "LastName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Recipient LastName:",
							"Recipient Last Name",
							"Enrollee Last Name"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "FirstName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"Recipient FirstName:",
							"Recipient First Name",
							"Enrollee First Name"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "MiddleInitial", DataFormat = DataFormatType.String,
						titles = new List<string> { "MdlInt." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "GrievanceDate", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date of Grievance" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "GrievanceType", DataFormat = DataFormatType.String,
						titles = new List<string> {
							"(1 - 11) Type of Grievance",
							"(1 - 13) Type of Grievance"
						}
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
							"(1 - 15) Type of Dispostion"
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
						titles = new List<string> {
							"Originator 1=Enrollee2 = Provider",
							"Originator 1=Enrollee 2 = Provider"
						}
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

			var wsLayout_cga_comp = new WorkSheetLayout
			{
				Name = "Enrollee Complaint Log",
				OutputFileName = "Data_Extract_Enrollee_Complaint_Log",
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
							new CellLocation { TitleRef = "A5", ValueRef = "B5" },
							new CellLocation { TitleRef = "A6", ValueRef = "B6" },
							new CellLocation { TitleRef = "A7", ValueRef = "B7" },
							new CellLocation { TitleRef = "A8", ValueRef = "B8" },
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
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A11" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B11" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C11" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D11" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E11" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F11" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G11" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H11" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I11" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J11" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K11" } },
						},
						FirstRow = 12
					}
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "ComplaintDate", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date Complaint Rcvd" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "ComplaintantLastName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Complainant Last Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "ComplaintantFirstName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Complainant First Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "EnrolleeLastName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Enrollee Last Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "EnrolleeFirstName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Enrollee First Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "EnrolleeMedicaidID", DataFormat = DataFormatType.String,
						titles = new List<string> { "Enrollee's Medicaid ID" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "ComplaintNature", DataFormat = DataFormatType.String,
						titles = new List<string> { "Nature of Complaint" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "ComplaintType", DataFormat = DataFormatType.String,
						titles = new List<string> { "Type of Complaint" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "ResolutionDescription", DataFormat = DataFormatType.String,
						titles = new List<string> { "Description of Resolution" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "DispositionDate", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date of Disposition" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "FinalDisposition", DataFormat = DataFormatType.String,
						titles = new List<string> { "Final Disposition" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 12, Name = "PlanName", DataFormat = DataFormatType.String, isRequired = false,
						titles = new List<string> { "Plan Name" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 13, Name = "PlanID", DataFormat = DataFormatType.String, isRequired = false,
						titles = new List<string> { "Plan's ID number" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 14, Name = "YR", DataFormat = DataFormatType.String, isRequired = false,
						titles = new List<string> { "Year" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 15, Name = "Month", DataFormat = DataFormatType.String, isRequired = false,
						titles = new List<string> { "Month" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 16, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 17, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
				}
			};

			var wsLayout_comp_log = new WorkSheetLayout
			{
				Name = "Humana_Enrollee Complaint Log",
				OutputFileName = "Data_Extract_Humana_Log_Complaint",
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
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A1" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B1" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C1" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D1" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E1" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F1" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G1" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H1" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I1" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J1" } },
						},
						FirstRow = 2
					}
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "ComplaintDate", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date Complaint Received" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "ComplaintantName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Complainant Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "EnrolleeName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Enrollee Name" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "EnrolleeMedicaidID", DataFormat = DataFormatType.String,
						titles = new List<string> { "Enrollee's Medicaid ID" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "ComplaintNature", DataFormat = DataFormatType.String,
						titles = new List<string> { "Nature of Complaint" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "ComplaintType", DataFormat = DataFormatType.String,
						titles = new List<string> { "Type of Complaint" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "ResolutionDescription", DataFormat = DataFormatType.String,
						titles = new List<string> { "Description of Resolution" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "DispositionDate", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date of Disposition" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "FinalDisposition", DataFormat = DataFormatType.String,
						titles = new List<string> { "Final Disposition" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "HumanaID", DataFormat = DataFormatType.String,
						titles = new List<string> { "Humana Member ID" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 11, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 12, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true }
				}
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
					},
					new CellLayoutVersion
					{
						Version = 8, // copied from v3
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "B3" },
							new CellLocation { TitleRef = "A4", ValueRef = "B4" },
						}
					},
					new CellLayoutVersion
					{
						Version = 9, // copied from v7
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
						}
					},
					new CellLayoutVersion
					{
						Version = 10, // copied from v8
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "D3" },
							new CellLocation { TitleRef = "A4", ValueRef = "D4" },
						}
					},
					new CellLayoutVersion
					{
						Version = 11, // copied from v4
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined }
						}
					},
					new CellLayoutVersion
					{
						Version = 12, // copied from v11
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined }
						}
					},
					new CellLayoutVersion
					{
						Version = 13, // copied from v8
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
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
					new ColumnLayoutVersion
					{
						Version = 8, // copied from v5,
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
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "S6" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "T6" } },
						},
						FirstRow = 6
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
							"Date of Birth (MM/DD/YYYY)",
							"DateofBirth(mm/dd/yyyy)",
							"DateofBirth (mm/dd/yyyy)"
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
							"Residential Setting Type (Home, ALF, SNF or AFCH)",
							"Residential Setting Type (Home, ALF, SNF, or AFCH)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "FacilityName", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Name of Facility",
							"Name of Facility (if applicable)",
							"Name of the Facility (if applicable)",
							"Name of the Facility",
							"Name of Facility (N/A if not applicable)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "FacilityLic",
						DataFormat = DataFormatType.String, postProcRegex = new List<Tuple<string, string>> { new Tuple<string, string>(@"[^0-9]", "") },
						titles = new List<string>
						{
							"Facility License Number",
							"Facility License Number (if applicable)",
							"Facility License Number (N/A if not applicable)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "Tansition", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Identify if transitioning into a SNF or back into Community (SNF, Community, or N/A)",
							"Identify if transitioning into a SNF or back into Community (snf, Community, or N/A)",
							"Residential Setting Type (Home, ALF, SNF or AFCH",
							"Identify if transitioning into a SNF or back into Community (SNF, Community, or N/A if not applicable)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "TransistionDate", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"Date of transition to SNF or Community (if applicable)",
							"Date of transition to SNF or Community",
							"Date of transition to SNF or Community (if applicable) (mm/dd/yyyy)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Form2515Date", DataFormat = DataFormatType.DateMixed,
						titles = new List<string>
						{
							"Date the 2515 form was sent to DCF if transitioning (if applicable)",
							"Date the 2515/2506A form was sent to DCF if transitioning (N/A if not applicable)(mm/dd/yyyy)",
							"Date the 2515 dorm was sent to DCF if transitioning (if applicable)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "canLocate", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Able to Locate? Y/N",
							"Able to Locate? Y/N or N/A"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 18, Name = "canContact", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Able to Contact? Y/N",
							"Able to Contact? Y/N or N/A"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 19, Name = "LastContactDate", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"If unable to contact or locate enrolee, date of last contact? (N/A if not applicable)",
							"If unable to contact or locate enrollee, date of last contact? (N/A if not applicable)",
							"If unable to contact or locate  enrollee,  date of last contact? (N/A if not applicable)(mm/dd/yyyy)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 20, Name = "Comments", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Comments including demonstration of attempts to contact enrolee if applicable",
							"Comments including demonstration of attempts to contact enrollee if applicable",
							"Comments (including demonstration of attempts to contact enrollee, N/A if not applicable)"
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
					new Field { fldType = FieldType.filePath, OutputOrder = 25, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.column, OutputOrder = 26, Name = "TotDaysTransTo2515_06Form", DataFormat = DataFormatType.String,
						titles = new List<string> { "Total number of days between Date of transition and Date 2515/2506 form was sent to DCF (N/A if not applicable)" } }
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
					new CellLayoutVersion
					{
						Version = 10, // copied from v2
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A5", ValueRef = "B5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 11, // copied from v4,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined }
						}
					},
					new CellLayoutVersion
					{
						Version = 12, // copied from v10
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A4", ValueRef = "A4", dataLayout = CellDataLayout.combined },
						}
					},
					new CellLayoutVersion
					{
						Version = 13, // copied from v5
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "A5", dataLayout = CellDataLayout.combined }
						}
					},
					new CellLayoutVersion
					{
						Version = 14, // copied from v5
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
						}
					},
					new CellLayoutVersion
					{
						Version = 15, // copied from v5
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "C2" },
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
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
					new ColumnLayoutVersion
					{
						Version = 6, // copied from v3
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
						},
						FirstRow = 8
					},
					new ColumnLayoutVersion
					{
						Version = 7, // copied from v6
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
						},
						FirstRow = 7 
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
							"Number of Authorized Service Units"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "Units_Missed", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Number of Missed Service Units In The Reported Month",
							"Number of Missed Services Units in the Reported Month",
							"Number of Missed Service Units per date of missed service",
							"Number of Missed Service Units per date missed",
							"number of missed service unit per date of missed service",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "MissedServiceCode", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Reason for Missed Service (Enter Code)",
							"Reason for Missed Services (Enter Code)",
							"reason for missed service ( enter code)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "MissedServiceDate", DataFormat = DataFormatType.DateMixed,
						titles = new List<string>
						{
							"Date of Missed Service (XX/XX/XXXX)",
							"Date of Missed Service or Date Range if Multiple Dates Missed (XX/XX/XXXX)",
							"Date of Missed Services or Date Range if Multiple Dates Missed (XX/XX/XXXXX)",
							"Date of Missed Services (MM/DD/YYYY)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "Explanation", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Explanation and Resolution of Missed Services",
							"Resolution of Missed Service /Comments",
							"Resolution of Missed Services / Comments",
							"Resolution of Missed Service & Comments",
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
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Region", DataFormat = DataFormatType.String,
						titles = new List<string> { "Region" },
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "County", DataFormat = DataFormatType.String,
						titles = new List<string> { "County of Residence" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 18, Name = "PercentMissed", DataFormat = DataFormatType.String,
						titles = new List<string> { "% of Authorized Service Units per date missed" }
					},
					new Field { fldType =  FieldType.column, OutputOrder = 19, Name = "MissedServicesNotification_DT" , DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date Managed Care Plan was Notified of Missed Service (MM/DD/YYYY)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 20, Name ="ServicesResumed_DT", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date Services Resumed (MM/DD/YYYY)" }
					}
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
					},
					new ColumnLayoutVersion
					{
						Version = 2, // copied from v1
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
						},
						FirstRow = 6
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "Action_Taken", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"ACTION TAKEN (N=New, A=Amended)",
							"ACTION TAKEN N=New or A=Amended"
						}
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
						titles = new List<string>
						{
							"CITY AND COUNTY",
							"EVENT CITY AND COUNTY"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "Event_Contact", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"EVENT CONTACT NAME:",
							"PLAN'S CONTACT NAME AND PHONE NUMBER:"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "Flier_Attached", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"INVITATION LETTER/FLYER ATTACHED: Yes/No",
							"INVITATION LETTER ATTACHED: Yes/No",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "Representatives", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"NAMES OF PARTICIPATING OUTREACH REPRESENTATIVES",
							"NAME(S) OF PLAN'S PARTICIPATING EVENT EDUCATIONAL REPRESENTATIVE(S)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "Service_Type", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"TYPE OF HEALTH RELATED SERVICE(S) TO BE PROVIDED BY OUTREACH REPRESENTATIVE",
							"TYPE(S) OF EDUCATIONAL FUNCTION TO BE PROVIDED BY REPRESENTATIVE(S) 1=Behavioral Health; 2=Disease Prevention; 3=Fitness & Exercise; 4=Food & Nutrition; 5=Gov’t Assistance Programs; 6=Preventive Techniques; 7=Stress Management; 8=Substance Abuse; 9=Wellness & Healthy Lifestyle; and/or 10=Other Clinical",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "Promo_Items_LT_5dol", DataFormat = DataFormatType.String,
						titles = new List<string> { "PROMO ITEMS < $5.00 Yes/No" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Outreach_Material", DataFormat = DataFormatType.String,
						titles = new List<string> { "OUTREACH MATERIAL PROVIDED Yes/No" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "StartTime", DataFormat = DataFormatType.String,
						titles = new List<string> { "START TIME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 18, Name = "EndTime", DataFormat = DataFormatType.String,
						titles = new List<string> { "END TIME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 19, Name = "Region", DataFormat = DataFormatType.String,
						titles = new List<string> { "EVENT REGION" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 20, Name = "Phone", DataFormat = DataFormatType.String,
						locType = LocateType.byRelated, RelatedOutputOrder = 11
					},
					new Field { fldType = FieldType.column, OutputOrder = 21, Name = "Promo_Items_LT_15dol", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"FILE NAME(S) OF AGENCY APPROVED PROMO/ NOMINAL GIFT ITEM(S) < $15.00 TO BE DISTRIBUTED",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 22, Name = "DistMaterials", DataFormat = DataFormatType.String,
						titles = new List<string> { "FILE NAME(S) OF AGENCY APPROVED MARKETING MATERIAL(S) TO BE DISTRIBUTED" }
					}
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
						titles = new List<string>
						{
							"ACTION R = Lisc/Certificate Renewed C = Info Updated N = New Representative T = Terminated",
							"ACTION R = Lic/Certificate Renewed C = Info Updated N = New Representative T = Terminated"
						}
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
						titles = new List<string>
						{
							"LIC / CERT ISSUED BY: (DOH, DPR, DFS, ect.,use NA if not applicable)",
							"LIC / CERT ISSUED BY: (DOH, DPR, DFS, etc.,use NA if not applicable)"
						}
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
					new Field { fldType = FieldType.cell, OutputOrder = 1, Name = "MC_PlanName", DataFormat = DataFormatType.String,
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

			var wsLayout_masr = new WorkSheetLayout
			{
				Name = "Marketing Agent Status Report",
				OutputFileName = "Data_Extract_Marketing_Agent_Status",
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
							new CellLocation { TitleRef = "A1", ValueRef = "B1" },
							new CellLocation { TitleRef = "C2", ValueRef = "E2" }
						}
					},

					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "D2", ValueRef = "E2" },
							new CellLocation { TitleRef = "F2", ValueRef = "G2" }
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
						},
						FirstRow = 4
					}
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "AgentStatus_ChangeAction", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"AGENT STATUS/ CHANGE ACTION If applicable enter C=Info updated; N=New agent; R=License renewed; or T=Agent terminated",
							"AGENT STATUS/ CHANGE ACTION C=Info updated; N=New agent; R=License renewed; T=Agent terminated; or NC=No change",
							"ACTION R = License Renewed C = Info Updated N = New Agent T = Agent Terminated",
							"ACTION R = Liscense Renewed C = Info Updated N = New Agent T = Agent Terminated"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "Status_Change_Date", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"DATE OF STATUS/ CHANGE ACTION (Date required if entry in first column)",
							"DATE OF STATUS/ CHANGE ACTION",
							"DaTE OF ACTION TAKEN"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "LastName", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"AGENT'S LAST NAME",
							"LAST NAME"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "FirstName", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"AGENT'S FIRST NAME",
							"FIRST NAME"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "Address", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"ADDRESS Street, City",
							"AGENT'S ADDRESS Street, City, State",
							"STREET ADDRESS"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "DFS_LicNum", DataFormat = DataFormatType.String,
						titles = new List<string> { "DFS LICENSE NUMBER" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "DFS_LicEffDate", DataFormat = DataFormatType.DateMixed,
						titles = new List<string>
						{
							"EFFECTIVE DATE OF CURRENT DFS LICENSE",
							"ORIGINAL VALID LICENSE ISSUE DATE",
							"DFS LICENSE ISSUE DATE"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "DFS_Lic_ExpDate", DataFormat = DataFormatType.DateMixed,
						titles = new List<string>
						{
							"EXPIRATION DATE OF CURRENT DFS LICENSE",
							"PLAN APPOINTMENT EXPIRATION DATE",
							"DFS LICENSE END DATE"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "Phone", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"AGENT'S TELEPHONE NUMBER",
							"MARKETING AGENT'S OFFICE PHONE NUMBER"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "Cell", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"AGENT'S CELLULAR TELEPHONE NUMBER",
							"MARKETING AGENT'S CELLULAR TELEPHONE"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "PrevEmpoyedPlans", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"NAMES OF ALL MANAGED CARE PLANS PREVIOULY EMPLOYED",
							"PREVIOUS HEALTH PLAN EMPLOYER"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "TerminationReasonfd", DataFormat = DataFormatType.String,
						titles = new List<string> { "REASON(S) FOR TERMINATION (Required field if entry of T in first column)" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 13, Name = "MC_PlanName", DataFormat = DataFormatType.String,
						titles = new List<string> { "Plan:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 14, Name = "Qtr", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Quarter Ending:" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 15, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 16, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "PlanIssueDate", DataFormat = DataFormatType.DateMixed,
						titles = new List<string> { "ORIGINAL PLAN APPOINTMENT ISSUE DATE" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 18, Name = "DateFiled", DataFormat = DataFormatType.Date,
						titles = new List<string> { "Date Filed:" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 19, Name = "City", DataFormat = DataFormatType.String,
						titles = new List<string> { "CITY" }}
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
					new ColumnLayoutVersion
					{
						Version = 3, // copied from v2
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
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "S6" } }
						},
						FirstRow = 7
					},
					new ColumnLayoutVersion
					{
						Version = 4, // copied from v1
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
							"ACTION TAKEN N=New or A=Amended"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "Event_Type", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"EVENT TYPE M=Marketing; P=Public; or E=Educational (Please see Tab 1's Definitions)",
							"EVENT TYPE:",
							"EVENT TYPE HC=Health Care Setting; PE=Public Event; or SO=State Office"
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
						titles = new List<string>
						{
							"TYPE OF MARKETING EVENT (if applicable) FE=Formal Event or IE=Informal Event (Please see Tab 1's Definitions)",
							"TYPE OF MARKETING EVENT FE=Formal Event or IE=Informal Event",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "Start_Date", DataFormat = DataFormatType.DateMixed, isRequired = true,
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
							"PLAN'S LEAD CONTACT NAME",
							"PLAN'S CONTACT NAME AND PHONE NUMBER:"
						}
					},
					// This match should only occur when there is a column with a field map at RelatedOuputOrder and next column location has no map.
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "Contact_Phone", DataFormat = DataFormatType.String,
						locType = LocateType.byRelated, RelatedOutputOrder = 14
					},
					//new Field { fldType = FieldType.column, OutputOrder = 15, Name = "Contact_Phone", DataFormat = DataFormatType.String,
					//	titles = new List<string> { "PLAN'S LEAD CONTACT PHONE NUMBER" }
					//},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Notice_Submitted", DataFormat = DataFormatType.String,
						titles = new List<string> { "INVITATION NOTICE SUBMITTED: Yes/No" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "Event_Representative", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"NAME(S) OF PLAN'S PARTICIPATING MARKETING AGENT(S) OR EVENT REPRESENATIVE(S)",
							"NAMES OF PARTICIPATING PLAN MARKETING REPRESENTATIVES/AGENTS",
							"NAME(S) OF PARTICIPATING PLAN MARKETING AGENT(S)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 18, Name = "Representative_DFS_LicNum", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"DFS LICENSE NUMBER(S) OF PARTICIPATING PLAN MARKETING AGENT(S) (if applicable)",
							"PARTICIPATING MARKETING REPRESENTATIVE/AGENT DFS LICENSE NUMBER",
							"DFS LICENSE NUMBER(S) OF PARTICIPATING PLAN MARKETING AGENT(S)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 19, Name = "Nominal_Items", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"FILE NAME(S) OF AGENCY APPROVED NOMINAL VALUE ITEM(S) < $15.00 TO BE DISTRIBUTED (if applicable)",
							"PROMO ITEMS < $15.00 RETAIL  Yes/No",
							"AGENCY APPROVED NOMINAL VALUE ITEM(S) < $15.00 TO BE DISTRIBUTED (Yes or No)",
							"FILE NAME(S) OF AGENCY APPROVED PROMO/ NOMINAL GIFT ITEM(S) < $15.00 TO BE DISTRIBUTED",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 20, Name = "Handout_Provided", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"FILE NAME(S) OF AGENCY APPROVED EVENT WRITTEN MATERIAL(S) TO BE DISTRIBUTED (if applicable)",
							"HANDOUT MATERIAL PROVIDED Yes/No",
							"FILE NAME(S) OF AGENCY APPROVED MARKETING MATERIAL(S) TO BE DISTRIBUTED",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 21, Name = "Comments", DataFormat = DataFormatType.String,
						titles = new List<string> { "Comments Entry required when Event amended (Please see Tab 1's Instructions)" }
					},
				}
			};

			var wsLayout_me_events_v2 = new WorkSheetLayout
			{
				Name = "Monthly Marketing/Public/Educational Events Report",
				OutputFileName = "Data_Extract_Marketing_Event_v2",
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
							new CellLocation { TitleRef = "A3", ValueRef = "E3" }
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
						},
						FirstRow = 7
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "Action", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"ACTION TAKEN N=New; A=Amended; or C=Canceled",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "Event_Type", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"EVENT TYPE M=Marketing; P=Public; or E=Educational (Please see Tab 1's Definitions)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "Event_Name", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string>
						{
							"EVENT NAME (Please see Tab 1's Instructions)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "Event_Formality", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"TYPE OF MARKETING EVENT (if applicable) FE=Formal Event or IE=Informal Event (Please see Tab 1's Definitions)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "Start_Date", DataFormat = DataFormatType.DateMixed, isRequired = true,
						titles = new List<string> { "START DATE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "End_Date", DataFormat = DataFormatType.Date,
						titles = new List<string> { "END DATE" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "Start_Time", DataFormat = DataFormatType.String,
						titles = new List<string> { "START TIME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "End_Time", DataFormat = DataFormatType.String,
						titles = new List<string> { "END TIME" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "Sponsor", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"EVENT SPONSOR NAME",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "Event_Addr", DataFormat = DataFormatType.String,
						titles = new List<string> { "PHYSICAL LOCATION OF EVENT (Street Address)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "Event_City_County", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"EVENT CITY AND COUNTY",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "Event_Region", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"EVENT REGION",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "Contact", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"PLAN'S LEAD CONTACT NAME AND PHONE NUMBER:",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "Event_Representative", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"NAME(S) OF PLAN'S PARTICIPATING MARKETING AGENT(S) OR EVENT REPRESENATIVE(S)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "Representative_DFS_LicNum", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"DFS LICENSE NUMBER(S) OF PARTICIPATING PLAN MARKETING AGENT(S) (if applicable)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Nominal_Items", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"FILE NAME(S) OF AGENCY APPROVED NOMINAL VALUE ITEM(S) < $15.00 TO BE DISTRIBUTED (if applicable)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "Handout_Provided", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"FILE NAME(S) OF AGENCY APPROVED EVENT WRITTEN MATERIAL(S) TO BE DISTRIBUTED (if applicable)",
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 18, Name = "MC_PlanName", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"PLAN:",
						}
					},
					new Field { fldType = FieldType.cell, OutputOrder = 19, Name = "Date", DataFormat = DataFormatType.DateMixed,
						titles = new List<string> { "Reporting Month" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 20, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 21, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },             }
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
					new CellLayoutVersion
					{
						Version = 8,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "A2", dataLayout = CellDataLayout.combined },
							new CellLocation { TitleRef = "A3", ValueRef = "A3", dataLayout = CellDataLayout.combined },
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
					new ColumnLayoutVersion
					{
						Version = 9,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A4" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B4" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C4" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D5" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E5" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F5" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G5" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H5" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I5" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J5" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K4" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L4" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M4" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N4" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O4" } },
						},
						FirstRow = 6
					},
					new ColumnLayoutVersion
					{
						Version = 10,
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
						Version = 11,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A6" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B6" } },
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
						},
						FirstRow = 7
					},
					new ColumnLayoutVersion
					{
						Version = 12,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 1, cellRefs = new List<string> { "A4" } },
							new ColumnTitleLocation { col = 2, cellRefs = new List<string> { "B4" } },
							new ColumnTitleLocation { col = 3, cellRefs = new List<string> { "C4" } },
							new ColumnTitleLocation { col = 4, cellRefs = new List<string> { "D5" } },
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "E4" } },
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "F5" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "G5" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "H5" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "I5" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "J5" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "K4" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "L4" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "M4" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "N4" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "O4" } },
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
						titles = new List<string>
						{
							"Medicaid ID",
							"`"
						}
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
							"PDO Enrollment Status (Enrolled/Disenrolled)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "EnrollmentDate", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"Enrollment Date (mm/dd/yyyy)",
							"Enrollment Date",
							"PDO Enrollment Date (mm/dd/yyyy)"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "DisenrollmentDate", DataFormat = DataFormatType.Date,
						titles = new List<string>
						{
							"Disenrollment Date (mm/dd/yyyy)",
							"Disenrollment Date",
							"PDO Disenrollment Date (mm/dd/yyyy)",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "DisenrollmentReason", DataFormat = DataFormatType.String,
						titles = new List<string>
						{
							"Disenrollment Reason (Reasons can be found in the PDO Manual)",
							"Disenrollment Reason",
							"PDO Disenrollment Reason Code"
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
						titles = new List<string>
						{
							"Reporting Month (MM/DD/YYYY):",
							"Month",
							"From",
							"Reporting Month"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "Region", DataFormat = DataFormatType.String, rowType = RowType.GroupData,
						titles = new List<string> { "REGION" }
					},
					new Field { fldType = FieldType.fileName, OutputOrder = 17, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.filePath, OutputOrder = 18, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.column, OutputOrder = 19, Name = "County", DataFormat = DataFormatType.String,
						titles = new List<string> { "County of Residence" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 20, Name = "Comments", DataFormat = DataFormatType.String,
						titles = new List<string> { "Comments" }
					},
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
					},
					new CellLayoutVersion
					{
						Version = 5,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "C2", dataLayout = CellDataLayout.separate },
							new CellLocation { TitleRef = "A3", ValueRef = "B3", dataLayout = CellDataLayout.separate },
							new CellLocation { TitleRef = "A4", ValueRef = "B4", dataLayout = CellDataLayout.separate },
						}
					},
					new CellLayoutVersion
					{
						Version = 6,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "B2", ValueRef = "B2", dataLayout = CellDataLayout.combined },
						}
					},
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
					new ColumnLayoutVersion
					{
						Version = 13, // copied from v5,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "A6" } },

							// 8 -> 25, 27 -> 39, 41 -> 51, 53 -> 55
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

							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
							new ColumnTitleLocation { col = 55, cellRefs = new List<string> { "A55" } },
						},
						FirstRow = 13
					},
					new ColumnLayoutVersion
					{
						Version = 14, // copied from v1,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "A6" } },

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
						Version = 15, // copied from v4,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 5, cellRefs = new List<string> { "B5" } },

							// 7 -> 24, 26 -> 38, 41 -> 51, 53 -> 55
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "B7" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "B8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "B9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "B10" } },
							new ColumnTitleLocation { col = 11, cellRefs = new List<string> { "B11" } },
							new ColumnTitleLocation { col = 12, cellRefs = new List<string> { "B12" } },
							new ColumnTitleLocation { col = 13, cellRefs = new List<string> { "B13" } },
							new ColumnTitleLocation { col = 14, cellRefs = new List<string> { "B14" } },
							new ColumnTitleLocation { col = 15, cellRefs = new List<string> { "B15" } },
							new ColumnTitleLocation { col = 16, cellRefs = new List<string> { "B16" } },
							new ColumnTitleLocation { col = 17, cellRefs = new List<string> { "B17" } },
							new ColumnTitleLocation { col = 18, cellRefs = new List<string> { "B18" } },
							new ColumnTitleLocation { col = 19, cellRefs = new List<string> { "B19" } },
							new ColumnTitleLocation { col = 20, cellRefs = new List<string> { "B20" } },
							new ColumnTitleLocation { col = 21, cellRefs = new List<string> { "B21" } },
							new ColumnTitleLocation { col = 22, cellRefs = new List<string> { "B22" } },
							new ColumnTitleLocation { col = 23, cellRefs = new List<string> { "B23" } },
							new ColumnTitleLocation { col = 24, cellRefs = new List<string> { "B24" } },

							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "B26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "B27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "B28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "B29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "B30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "B31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "B32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "B33" } },
							new ColumnTitleLocation { col = 34, cellRefs = new List<string> { "B34" } },
							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "B35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "B36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "B37" } },
							new ColumnTitleLocation { col = 38, cellRefs = new List<string> { "B38" } },

							new ColumnTitleLocation { col = 41, cellRefs = new List<string> { "B41" } },
							new ColumnTitleLocation { col = 42, cellRefs = new List<string> { "B42" } },
							new ColumnTitleLocation { col = 43, cellRefs = new List<string> { "B43" } },
							new ColumnTitleLocation { col = 44, cellRefs = new List<string> { "B44" } },
							new ColumnTitleLocation { col = 45, cellRefs = new List<string> { "B45" } },
							new ColumnTitleLocation { col = 46, cellRefs = new List<string> { "B46" } },
							new ColumnTitleLocation { col = 47, cellRefs = new List<string> { "B47" } },
							new ColumnTitleLocation { col = 48, cellRefs = new List<string> { "B48" } },
							new ColumnTitleLocation { col = 49, cellRefs = new List<string> { "B49" } },
							new ColumnTitleLocation { col = 50, cellRefs = new List<string> { "B50" } },
							new ColumnTitleLocation { col = 51, cellRefs = new List<string> { "B51" } },

							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "B53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "B54" } },
							new ColumnTitleLocation { col = 55, cellRefs = new List<string> { "B55" } },
						},
						FirstRow = 3
					},
					new ColumnLayoutVersion
					{
						Version = 15, // copied from v4,
						colLayoutType = ColLayoutType.Col_Row,
						titleLocations = new List<ColumnTitleLocation>
						{
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "A6" } },

							// 7 -> 24, 26 -> 38, 41 -> 51, 53 -> 55
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "A7" } },
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

							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
							new ColumnTitleLocation { col = 55, cellRefs = new List<string> { "A55" } },
						},
						FirstRow = 2
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
							"Enrollee Medicaid ID Number:",
							"Enrollee Medicaid ID Number:Y/N or N/A",
							"Enrollee ID NumberY or N",
							"Initial Contact",
							"ENROLLEE MEDICAID NUMBER"
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

			var wsLayout_mccma2 = new WorkSheetLayout
			{
				Name = "Managed Care Case Management File Audit Report v2",
				OutputFileName = "Data_Extract_Case_Management_Audit2",
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
							new CellLocation { TitleRef = "A2", ValueRef = "B2", dataLayout = CellDataLayout.separate },
							new CellLocation { TitleRef = "A3", ValueRef = "B3", dataLayout = CellDataLayout.separate },
							new CellLocation { TitleRef = "A4", ValueRef = "B4", dataLayout = CellDataLayout.separate }
						}
					},
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
							// 6 -> 10, 12 -> 21, 23 -> 33, 35 -> 50, 52 -> 64
							new ColumnTitleLocation { col = 6, cellRefs = new List<string> { "A6" } },
							new ColumnTitleLocation { col = 7, cellRefs = new List<string> { "A7" } },
							new ColumnTitleLocation { col = 8, cellRefs = new List<string> { "A8" } },
							new ColumnTitleLocation { col = 9, cellRefs = new List<string> { "A9" } },
							new ColumnTitleLocation { col = 10, cellRefs = new List<string> { "A10" } },

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

							new ColumnTitleLocation { col = 23, cellRefs = new List<string> { "A23" } },
							new ColumnTitleLocation { col = 24, cellRefs = new List<string> { "A24" } },
							new ColumnTitleLocation { col = 25, cellRefs = new List<string> { "A25" } },
							new ColumnTitleLocation { col = 26, cellRefs = new List<string> { "A26" } },
							new ColumnTitleLocation { col = 27, cellRefs = new List<string> { "A27" } },
							new ColumnTitleLocation { col = 28, cellRefs = new List<string> { "A28" } },
							new ColumnTitleLocation { col = 29, cellRefs = new List<string> { "A29" } },
							new ColumnTitleLocation { col = 30, cellRefs = new List<string> { "A30" } },
							new ColumnTitleLocation { col = 31, cellRefs = new List<string> { "A31" } },
							new ColumnTitleLocation { col = 32, cellRefs = new List<string> { "A32" } },
							new ColumnTitleLocation { col = 33, cellRefs = new List<string> { "A33" } },


							new ColumnTitleLocation { col = 35, cellRefs = new List<string> { "A35" } },
							new ColumnTitleLocation { col = 36, cellRefs = new List<string> { "A36" } },
							new ColumnTitleLocation { col = 37, cellRefs = new List<string> { "A37" } },
							new ColumnTitleLocation { col = 38, cellRefs = new List<string> { "A38" } },
							new ColumnTitleLocation { col = 39, cellRefs = new List<string> { "A39" } },
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

							new ColumnTitleLocation { col = 52, cellRefs = new List<string> { "A52" } },
							new ColumnTitleLocation { col = 53, cellRefs = new List<string> { "A53" } },
							new ColumnTitleLocation { col = 54, cellRefs = new List<string> { "A54" } },
							new ColumnTitleLocation { col = 55, cellRefs = new List<string> { "A55" } },
							new ColumnTitleLocation { col = 56, cellRefs = new List<string> { "A56" } },
							new ColumnTitleLocation { col = 57, cellRefs = new List<string> { "A57" } },
							new ColumnTitleLocation { col = 58, cellRefs = new List<string> { "A58" } },
							new ColumnTitleLocation { col = 59, cellRefs = new List<string> { "A59" } },
							new ColumnTitleLocation { col = 60, cellRefs = new List<string> { "A60" } },
							new ColumnTitleLocation { col = 61, cellRefs = new List<string> { "A61" } },
							new ColumnTitleLocation { col = 62, cellRefs = new List<string> { "A62" } },
							new ColumnTitleLocation { col = 63, cellRefs = new List<string> { "A63" } },
							new ColumnTitleLocation { col = 64, cellRefs = new List<string> { "A64" } },
						},
						FirstRow = 2
					},
				},

				fields = new List<Field>
				{
					new Field { fldType = FieldType.column, OutputOrder = 1, Name = "EnrolleeID", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> {
							"Enrollee Medicaid ID Number:",
							"Enrolle Medicaid ID Number:",
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 2, Name = "LastName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Enrollee Last Name:" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 3, Name = "FirstName", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Enrollee First Name:" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 4, Name = "DoB", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string> { "Enrollee Date of Birth:" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 5, Name = "EnrollmentDate", DataFormat = DataFormatType.Date, isRequired = true,
						titles = new List<string> { "Enrollment Date:" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 6, Name = "CompAssesIsCurrent", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Comprehensive assessment is current" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 7, Name = "PersEmrgPlanInFile", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Personal emergency plan is included in the case file?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 8, Name = "EnrolleeNeedsShelter", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Case manager determined whether the enrollee needed to register with Special Needs Shelter, and assisted with this registration" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 9, Name = "PhysicianIdentitied", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Enrollee's primary care physician is identified in the file" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 10, Name = "DocumentedProviderChoice", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Is there documentation of individual provider choice?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 11, Name = "SignedFreedomOfChoice", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> {
							"The enrollee completed and signed a Freedom of Choice form?",
							"Did the enrollee complete and sign a Freedom of Choice form?"
						}
					},
					new Field { fldType = FieldType.column, OutputOrder = 12, Name = "ChoiceDiscrepancyExplained", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "If the enrollee's choice is to live in a different placement type than where they currently reside, do the case notes explain the discrepancy and show all reasonable efforts made to accommodate the enrollees choice?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 13, Name = "DenialNoticeSent", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "If there was a denial, reduction, termination, or suspension of services, was a notice of action sent?"}
					},
					new Field { fldType = FieldType.column, OutputOrder = 14, Name = "NotifiedAboutPDO", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "The enrollee was informed of the option to participate in PDO (if receiving Companion, Attendant Care, Homemaker, Intermittent/Skilled Nursing Care, Personal Care)" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 15, Name = "ReferedToPublicGuardian", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "If the enrollee is not capable of making decisions and does not have a representative, did the case manager refer the enrollee to the Public Guardianship Program or other advocacy resource?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 16, Name = "OnSiteVisitIn5days", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Initial on-site visit to develop plan of care conducted within 5 business days of enrollment  if enrollee resides in community" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 17, Name = "OnSiteVisitIn7daysForNH", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Initial on-site visit to develop plan of care conducted within 7 business days if in nursing facility" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 18, Name = "RightsExplained", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "The case manager explained the enrollee's rights and responsibilities" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 19, Name = "FilingGreivanceExplained", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "The case manager explained the procedures for filing a grievance" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 20, Name = "FilingAppealExplained", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "The case manager explained the procedures for filing an appeal" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 21, Name = "FilingHearingExplained", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "The case manager explained the procedures for filing a fair hearing" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 22, Name = "PlanIdCardProvided", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan Identification card was provided to enrollee" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 23, Name = "HandbookProvided", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Enrollee Handbook was provided to enrollee" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 24, Name = "ProviderDirectoyProvided", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Provider Directory was provided to enrollee" }
					},
					new Field { fldType = FieldType.column, OutputOrder  = 25, Name = "DiscussedAdvancedDirectives", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Advance Directives were discussed with the enrollee" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 26, Name = "PhoneFollowupWithin7days", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Telephone follow up with enrollee or representative is completed within 7 business days." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 27, Name = "ServicesMatchAssessedNeeds", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan of care services are specific to assessed needs" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 28, Name = "PlanHasSupportsDocumented", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan of care contains documentation of services and supports regardless of the funding source" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 29, Name = "PlanDocServTypeScopeAmountDurationFreq", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Does plan of care document service type, scope, amount, duration and frequency of services" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 30, Name = "ConsistentServiceAuthorizations", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Service authorizations are consistent with plan of care" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 31, Name = "PlanEnrolleeSigned", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan of care was signed and dated by enrollee or representative" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 32, Name = "SummaryEnrolleeSigned", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan of care summary is signed and dated by enrollee or representative" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 33, Name = "EnrolleeHasPlanCopy", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "The case manager provided a copy of the plan of care to the enrollee or representative" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 34, Name = "EnrolleeHasSummaryCopy", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "The case manager provide a copy of the plan of care summary to the enrollee or representative." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 35, Name = "PlanUpdatedAnnually", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan of care is updated at least annually" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 36, Name = "ReviewedFaceEvery90Days", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan of care is reviewed in a face-to-face visit every 90 calendar days." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 37, Name = "ReviewedFaceLT90Days", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan of care is reviewed and updated in a face-to-face visit more frequently than once every 90 calendar days if the enrollee's condition changes or requires it" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 38, Name = "FaceReviewWithin5daysOfPlacementChange", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "The case manager conducted a face-to-face review within 5 business days following an enrollee's change of placement type (e.g., HCBS to an institutional setting, own home to assisted living facility, or institutional setting to HCBS)." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 39, Name = "DocumentBarriersInterventions", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Are barriers and interventions documented in the plan of care?" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 40, Name = "PlanHasPersonalGoals", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan of care contains personal goals" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 41, Name = "PlanSentPCPwithin10days", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan of care was forwarded to the enrollee's PCP within 10 business days of development" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 42, Name = "PlanSentFacWithin10days", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Plan of care was forwarded to the facility within 10 business days of development" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 43, Name = "MonthlyContactsComplete", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Monthly contacts are completed" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 44, Name = "HasEnrolleeCurrMedicalStatus", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Information on the enrollee's current medical status." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 45, Name = "HasEnrolleeCurrFuncStatus", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Information on the enrollee's current functional status." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 46, Name = "HasEnrolleeBehavStatus", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Information on the enrollee's behavioral health status." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 47, Name = "HasEnrolleeCurrStrengthNeeds", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Information on the enrollee's current strengths and needs." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 48, Name = "HasSpecialNeeds", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Information on special needs, if applicable" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 49, Name = "HasEnvironmentalConcerns", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Information of any environmental concerns" }
					},
					new Field { fldType = FieldType.column, OutputOrder = 50, Name = "IdentFamilySupportAvail", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Identification of family/informal support and their availability to assist the enrollee, including barriers to assistance." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 51, Name = "HasHomeLikeEnvInALFandAFCH", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Documentation of home-like environment characteristics for enrollees residing in ALF and AFCH facilities." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 52, Name = "HasHomeLikeInOwnHome", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Documentation of home characteristics for enrollees residing in their own home or non-facility based residence." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 53, Name = "ReceiptServicesSatisfactionDocumented", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Documentation of receipt of services, and enrollee satisfaction with services." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 54, Name = "DiscussGoalsBarriersInterventionsStatus", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Discussion of personal goals, barriers, interventions and the status of personal goals." }
					},
					new Field { fldType = FieldType.column, OutputOrder = 55, Name = "DocumentProviderProblemsAndActionPlan", DataFormat = DataFormatType.String, isRequired = true,
						titles = new List<string> { "Documentation of problems with service providers and the planned course of action, if applicable." }
					},



					new Field { fldType = FieldType.cell, OutputOrder = 56, Name = "MC_PlanName", DataFormat = DataFormatType.String,
						titles = new List<string> {
							"Humana American Eldercare",
							"LTC Managed Care Organization",
							"Molina Healthcare of Florida, Inc."
						}
					},

					new Field { fldType = FieldType.cell, OutputOrder = 57, Name = "Date", DataFormat = DataFormatType.String,
						titles = new List<string> { "Date:" }
					},
					new Field { fldType = FieldType.cell, OutputOrder = 58, Name = "QTR", DataFormat = DataFormatType.String,
						titles = new List<string> { "Quarter:" }
					},

					new Field { fldType = FieldType.filePath, OutputOrder = 59, Name = "FilePath", DataFormat = DataFormatType.String, isRequired = true },
					new Field { fldType = FieldType.fileName, OutputOrder = 60, Name = "FileName", DataFormat = DataFormatType.String, isRequired = true }
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
						new SheetLayout { Names = new List<string> { "January", "January G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "February", "February G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "March", "March G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "April", "April G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "May", "May G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "June", "June G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "July", "July G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "August", "August G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "September", "September G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "October", "October G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "November", "November G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "December", "December G&A" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "Summary", "Summary G&A", "Summary C" }, isOptional = true },
						new SheetLayout { Names = new List<string> { "October 2014" }, sheetType = SheetType.SourceData, isOptional = true, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "November 2014" }, sheetType = SheetType.SourceData, isOptional = true, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "December 2014" }, sheetType = SheetType.SourceData, isOptional = true, wsLayout = wsLayout_cga },
						new SheetLayout { Names = new List<string> { "Sheet1", "Sheet2", "Sheet3" }, isOptional = true },
						new SheetLayout { Names = new List<string> { "Jan C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "Feb C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "Mar C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "Apr C", "Apr MMA", "Apr HK" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "May C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "Jun C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "Jul C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "Aug C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "Sep C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "Oct C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "Nov C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
						new SheetLayout { Names = new List<string> { "Dec C" }, isOptional = true, sheetType = SheetType.SourceData, wsLayout = wsLayout_cga_comp },
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
						new SheetLayout { Names = new List<string> { "Monthly Events Report" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_me_events },
						new SheetLayout { Names = new List<string> { "sheet1", "sheet2", "sheet3" }, isOptional = true }
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
								"Educational Events"
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
						new SheetLayout { sheetType = SheetType.SourceData, wsLayout = wsLayout_mccma },
					}
				},

				new SpreadSheetLayout
				{
					Name = "Managed Care Case Management File Audit2",
					procType = ProcessType.MatchByClosestWorkSheetLayout,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { sheetType = SheetType.SourceData, wsLayout = wsLayout_mccma2 }
					}
				},

				new SpreadSheetLayout
				{
					Name = "Humana Log of Complaints",
					procType = ProcessType.MatchByClosestWorkSheetLayout,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { sheetType = SheetType.SourceData, wsLayout = wsLayout_comp_log }
					}
				},

				new SpreadSheetLayout
				{
					Name = "Marketing Agent Status Report",
					procType = ProcessType.MatchByClosestWorkSheetLayout,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { Names = new List<string> { "Marketing Agent Status" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_masr }
					}
				},

				new SpreadSheetLayout
				{
					Name = "Marketing Agent Status and Outreach Report",
					procType = ProcessType.MatchAllDataWorkSheets,
					types = dst,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { Names = new List<string> { "Instructions" }, isOptional = true },
						new SheetLayout { Names = new List<string> { "Jurat" }, sheetType = SheetType.CommonData, wsLayout = wsLayout_cor_jurat },
						new SheetLayout { Names = new List<string> { "Marketing Activity" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_masr },
						new SheetLayout { Names = new List<string> { "Community Outreach Activity" }, sheetType = SheetType.SourceData, wsLayout = wsLayout_cor_activity },
					}
				},

				new SpreadSheetLayout
				{
					Name = "Montly Marketing/Public/Educational Events Report",
					procType = ProcessType.MatchByClosestWorkSheetLayout,
					types = dst,
					sLayouts = new List<SheetLayout> { new SheetLayout {  sheetType = SheetType.SourceData, wsLayout = wsLayout_me_events_v2 } }
				}


			};

			dst.types.ForEach(ssl => ssl.sLayouts.ForEach(sl => sl.ssLayout = ssl));

			return dst.types;
		}
	}
}
