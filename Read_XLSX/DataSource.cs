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
	class MatchData
	{
		public bool isPass { get; set; }

		public List<FieldColumnVersionMap> fldColVersMaps { get; set; }

		public List<FieldCellVersionMap> fldCellVersMaps { get; set; }
	}

	class DataSourceTypes
	{
		public readonly DateTime timeStamp;

		public string RootFolder;

		public List<SpreadSheetLayout> types { get; set; }

		public DataSourceTypes(string rootFolder)
		{
			RootFolder = rootFolder;
			timeStamp = DateTime.Now;
			Init();
		}

		/// <remarks>
		/// Two ways to determin type based on value matchWorkSheetNames
		/// - when true then worksheet names must match DataSource type in order to be a match.
		/// - when false then if any worksheets match data source layout then is a match
		/// </remarks>
		public SpreadSheetLayout DetermineLayout(SpreadsheetDocument ssd, FileInfo file)
		{
			var mds = new List<MatchData>();

			// clear working references.
			types.ForEach(t => t.sLayouts.ForEach(s => 
			{
				if(s.srcWorksheets != null)
				{
					s.srcWorksheets.Clear();
					s.srcWorksheets = null;
				}
			
				if(s.dataSet != null)
				{
					if(s.dataSet.Rows != null)
					{
						s.dataSet.Rows.ToList().ForEach(r =>
						{
							r.Value.Cells.Clear();
							r.Value.Cells = null;
						});

						s.dataSet.Rows.Clear();
						s.dataSet.Rows = null;
					}
					s.dataSet.wsLayout = null;
					s.dataSet = null;
				}
				
				if (s.wsLayout != null)
				{
					if(s.wsLayout.fieldColMap != null)
					{
						if(s.wsLayout.fieldColMap.colmaps != null)
						{
							s.wsLayout.fieldColMap.colmaps.Clear();
							s.wsLayout.fieldColMap.colmaps = null;
						}

						s.wsLayout.fieldColMap = null;
					}
					
					if(s.wsLayout.fieldCellMap != null)
					{
						s.wsLayout.fieldCellMap.fldmaps.Clear();
						s.wsLayout.fieldCellMap.fldmaps = null;
						s.wsLayout.fieldCellMap = null;
					}
				}
			}));

			SpreadSheetLayout type = null;

			WorkbookPart wbp = ssd.WorkbookPart;
			var stringTable = wbp.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
			var cellFormats = wbp.WorkbookStylesPart.Stylesheet.CellFormats;

			var shts = wbp.Workbook.Descendants<Sheet>();

			var procTypes = types.Where(p => p.procType == ProcessType.MatchByClosestWorkSheetLayout || (shts.Count() >= p.sLayouts.Count(sl => !sl.isOptional) && shts.Count() <= p.sLayouts.Count()));

			procTypes = procTypes
				.Where(r => 
					r.procType == ProcessType.MatchByClosestWorkSheetLayout || 
					(
						r.procType == ProcessType.MatchAllDataWorkSheets && 
						r.sLayouts.Where(sl => !sl.isOptional && shts.Count() >= r.sLayouts.Count() && sl.Name != shts.ElementAt(r.sLayouts.IndexOf(sl)).Name).Count() == 0
					)
				)
				.ToList();

			if (procTypes.Count() == 0) return null;

			foreach (var dst in procTypes)
			// Iterate through types
			{
				bool isPass = true;

				switch (dst.procType)
				{
					case ProcessType.MatchAllDataWorkSheets:

						foreach(var sht in wbp.Workbook.Descendants<Sheet>())
						{
							var sheetLayout = dst.sLayouts.FirstOrDefault(sl => sl.Name == sht.Name);

							if (sheetLayout == null || sheetLayout.wsLayout == null) continue;

							WorksheetPart wsp = wbp.GetPartById(sht.Id) as WorksheetPart;

							var md = MatchLayouts(wsp.Worksheet, sheetLayout, stringTable, cellFormats, file);

							mds.Add(md);

							isPass &= md.isPass;
						}
						break;

					case ProcessType.MatchByClosestWorkSheetLayout:

						isPass = false;

						foreach (var sheetLayout in dst.sLayouts.Where(sl => sl.wsLayout != null))
						{
							if (sheetLayout.wsLayout == null) continue;

							foreach (var sht in wbp.Workbook.Descendants<Sheet>())
							{
								WorksheetPart wsp = wbp.GetPartById(sht.Id) as WorksheetPart;

								var md = MatchLayouts(wsp.Worksheet, sheetLayout, stringTable, cellFormats, file);

								isPass |= md.isPass;

								mds.Add(md);

								if (isPass) break;
							}

							if (isPass) break;
						}
						break;
				}

				if (isPass)
				{
					type = dst;
					break;
				}
			}
			
			if(type == null)
			{
				// Log best no match data.
				var cmds = mds.Select(m => new { md = m, noColMatchCntMin = m.fldColVersMaps.Min(cv => cv.noMatchCnt * 10000 + cv.noneNullTitleCnt * 100 + cv.disOrder), noCellMatchCntMin = m.fldCellVersMaps.Min(cev => cev.noMatchCnt * 10000 + cev.missingReqFldCnt * 100 + cev.noneNullTitleCnt) });
				var lmd = cmds.OrderBy(ds => ds.noColMatchCntMin).ThenBy(ds => ds.noCellMatchCntMin).FirstOrDefault();
				Log.New.Msg($"FAILURE: {file.FullName}: Unable to determine format type of file");
				lmd.md.fldColVersMaps.FirstOrDefault().colmaps.Where(cp => cp.field == null).OrderBy(cp => cp.column).ToList().ForEach(cpp =>
				{
					Log.New.Msg($"\t\tCol: {cpp.column}, Title: {cpp.title}");
				});
				Log.New.Msg("\t\t---------------");
				lmd.md.fldCellVersMaps.FirstOrDefault().fldmaps.Where(fp => fp.field == null).ToList().ForEach(fp =>
				{
					Log.New.Msg($"\t\tCell: {fp.cellLoc}, Title: {fp.Title}");
				});
			}

			return type;
		}

		/// <remarks>
		///		An Excel SpreadSheet contains one or more Worksheets, each may or maynot contain data of interest.
		///			- Excel is a terrible way to collect data from a large number of different sources in a consistent and reliable way.
		///			- Be that as it may, Excel is favored by organizations that prefer manpower over automation when performing data processing tasks.
		///				- Most state agencies are typical of this kind of organization.
		///					- To top it off most of these agencies give little thought to gathering data in a consistent way. So we are likely to
		///						recieve a dump of spreadsheets with a variety of inconsistencies.
		///		
		///		As varied as these spreadsheets may be, a spreadsheet is expected to contain only a single type of data set which is called a
		///			DataSourceType in this application.
		///		
		///		A DataSourceType describes how to process the worksheets in a spreadsheet. It indicates:
		///			- the name of the file to save extracted data.
		///			- a list of DataWorkSheets
		///			- an indicator as to how to process the spreadsheet against the list of DataWorkSheets:
		///			
		///				• MatchAllDataWorkSheets 
		///					- There must be a one to one correspondence between each DataWorkSheet and each SpreadSheet WorkSheet in order.
		///					- The DataWorkSheet name must match the SpreadSheet WorkSheet name.
		///					
		///				• MatchByClosestWorkSheetLayout 
		///					- Each SpreadSheet Worksheet will be matched against the closest DataWorkSheet/WorkSheetLayout
		///		
		///		A DataWorkSheet has
		///			- a Name to be used when processing the spreadsheet by MatchAllDataWorkSheets.
		///			- a WorkSheetLayout
		/// 
		///		A WorkSheetLayout is
		///			- a collection of field layout versions with additional information about how to determine where to look 
		///				for field cells on the spreadsheet.
		///			- a collection of data columns with addition information about how to determine where to look
		///				for the data column on the WorkSheet.
		///		
		///		- Each data column in the collection has associated with it a list of column titles that should 
		///			map from the WorkSheet to the data column.
		///			
		///		- The WorkSheetLayout also includes a collection of col layout versions. Each of these is a list of cells 
		///			that should be scraped for strings that are concatinated into a column title and the column to associate
		///			the title with.
		///			
		///		- For a given WorkSheet all layouts are processed 
		///		
		///		There is a layout of column title cells that will be scaped for column titles. Those titles 
		///		are then matched to a list of titles for a given data column. The assumption being that all titles to match
		///		are unique across all data columns for a given WorkSheetLayout
		/// </remarks>
		public MatchData MatchLayouts(Worksheet ws, SheetLayout sheetLayout, SharedStringTablePart stringTable, CellFormats formats, FileInfo file)
		{
			// All cells in worksheet.
			var tcs = ws.Descendants<Cell>();

			// Find and map columns to sheet layouts
			var field_ord = sheetLayout.wsLayout.fields
							.Where(c => c.fldType == FieldType.column)
							.OrderBy(c => c.OutputOrder)
							.Select(c => c.OutputOrder)
							.ToList();

			var md = new MatchData();

			// Obtain column titles for all signature versions.
			md.fldColVersMaps = new List<FieldColumnVersionMap>();

			foreach(var sig in sheetLayout.wsLayout.colLayouts)
			// for each column layout version scape the worksheet for column title values
			{
				var fldColMaps = new List<FieldColumnMap>();

				var col_ord = sig.titleLocations.OrderBy(so => so.col).Select(so => so.col).ToList();

				foreach(var colLayout in sig.titleLocations)
				{
					string title = "";
					foreach(var c in colLayout.cellRefs)
					// A column may have a number of title cells that must be scraped and concatinated to product the title used for matching to data columns.
					{
						var cl = tcs.FirstOrDefault(cll => cll.CellReference.InnerText == c);
						var tlt = Spreadsheet.GetCellValue(cl, stringTable.SharedStringTable, formats, null);
						title += tlt;
					}

					title = System.Text.RegularExpressions.Regex.Replace(title.Replace('\n', ' '), @"\s+", " ").Trim();
					fldColMaps.Add(new FieldColumnMap { column = colLayout.col, title = title, col_order = col_ord.IndexOf(colLayout.col) });
				}

				md.fldColVersMaps.Add(new FieldColumnVersionMap { colLayout = sig, colmaps = fldColMaps });
			}
					
			// Match the titles to the DataColumns
			foreach(var fcvm in md.fldColVersMaps)
			{
				foreach(var cm in fcvm.colmaps)
				{
					try
					{
						cm.field = sheetLayout.wsLayout.fields.Where(cc => cc.fldType == FieldType.column && cc.titles != null).FirstOrDefault(cc => cc.titles.Contains(cm.title));
						cm.field_order = cm.field != null ? field_ord.IndexOf(cm.field.OutputOrder) : -9999;
					}
					catch(Exception ex)
					{
						Log.New.Msg(ex);
					}
				}

				fcvm.noneNullTitleCnt = fcvm.colmaps.Where(cm => !string.IsNullOrWhiteSpace(cm.title)).Count();
				fcvm.noMatchCnt = fcvm.colmaps.Where(cm => cm.field == null).Count();
				fcvm.disOrder = (int)fcvm.colmaps.Where(dm => dm.field != null).Select(dm => Math.Pow((dm.field_order - dm.col_order), 2)).Sum();
				fcvm.colDups = fcvm.colmaps.Where(dm => dm.field != null).GroupBy(cd => cd.field).Where(d => d.Count() > 1).Count();
			}

			// Only match col layout versions with zero mismatch, favoring the version with the lowest disorder.
			var colLayout_v = md.fldColVersMaps.Where(sv => sv.noMatchCnt == 0 && sv.colDups == 0).OrderByDescending(sv => sv.disOrder).FirstOrDefault();

			sheetLayout.wsLayout.fieldColMap = colLayout_v;

			// Obtain titles for all field cell layouts
			md.fldCellVersMaps = new List<FieldCellVersionMap>();

			foreach (var fldLayout in sheetLayout.wsLayout.cellLayouts)
			{
				var fldLayoutVals = new List<FieldCellMap>();

				foreach(var cellLoc in fldLayout.cellLocations)
				{
					try
					{
						var cl = tcs.FirstOrDefault(cll => cll.CellReference.InnerText == cellLoc.TitleRef);
						var clVal = tcs.FirstOrDefault(clv => clv.CellReference.InnerText == cellLoc.ValueRef);

						var title = Spreadsheet.GetCellValue(cl, stringTable.SharedStringTable, formats, null);
						var val = Spreadsheet.GetCellValue(clVal, stringTable.SharedStringTable, formats, null);

						fldLayoutVals.Add(new FieldCellMap
						{
							cellLoc = cellLoc,
							Title = string.IsNullOrWhiteSpace(title) ? null : title.Trim(),
							Value = string.IsNullOrWhiteSpace(val) ? null : val.Trim()
						});
					}
					catch(Exception ex)
					{
						Log.New.Msg(ex);
					}
				}

				md.fldCellVersMaps.Add(new FieldCellVersionMap { fldmaps = fldLayoutVals, fldLayout = fldLayout });
			}

			var reqFlds = sheetLayout.wsLayout.fields.Where(sf => sf.fldType == FieldType.cell && sf.isRequired);

			// Match Titles to layout fields
			foreach (var flvv in md.fldCellVersMaps)
			{
				foreach(var fm in flvv.fldmaps.Where(m => m.Title != null))
				{
					foreach(var fld in sheetLayout.wsLayout.fields.Where(f => f.fldType == FieldType.cell))
					{
						try
						{
							if (fm.cellLoc.isCombined)
							{
								var titles = fld.titles.Where(t => fm.Title.StartsWith(t));

								if (titles.Count() > 0)
								{
									fm.field = fld;
									fm.Value = fm.Value != null ? fm.Value.Replace(titles.FirstOrDefault(), "").Trim() : null;
									if (fld.DataFormat == DataFormatType.Date || fld.DataFormat == DataFormatType.DateTime)
									{
										var val = fm.Value.Replace("(", "").Replace(")", "").Replace(":", "");
										DateTime outVal;
										if (DateTime.TryParse(val, out outVal))
										{
											if (fld.DataFormat == DataFormatType.Date)
												fm.Value = outVal.ToShortDateString();
											else
												fm.Value = outVal.ToString();
										}
										else
											fm.Value = null;
									}
									break;
								}
							}
							else
							{
								if (fm.Title != null && fld.titles.Contains(fm.Title))
								{
									fm.field = fld;

									if (fld.DataFormat == DataFormatType.Date || fld.DataFormat == DataFormatType.DateTime)
									{
										var cell = tcs.Where(c => c.CellReference == fm.cellLoc.ValueRef).FirstOrDefault();
										fm.Value = Spreadsheet.GetCellValue(cell, stringTable.SharedStringTable, formats, fld);
									}

									break;
								}
							}
						}
						catch (Exception ex)
						{
							Log.New.Msg(ex);
						}
					}
				}

				var fileFld = sheetLayout.wsLayout.fields.FirstOrDefault(fld => fld.fldType == FieldType.fileName);
				if ( fileFld != null)
				{
					flvv.fldmaps.Add(new FieldCellMap { field = fileFld, Value = file.FullName });
				}

				flvv.noneNullTitleCnt = flvv.fldmaps.Where(fm => !string.IsNullOrWhiteSpace(fm.Title)).Count();
				flvv.noMatchCnt = flvv.fldmaps.Where(fm => fm.field == null).Count();
				flvv.missingReqFldCnt = reqFlds.Where(rf => !flvv.fldmaps.Select(fm => fm.field).Contains(rf)).Count();
				flvv.noValCnt = flvv.fldmaps.Where(fm => fm.field != null && fm.field.isRequired && string.IsNullOrWhiteSpace(fm.Value)).Count();
			}

			var fldLayout_v = md.fldCellVersMaps.Where(fl => fl.noMatchCnt == 0 && fl.noValCnt == 0 && fl.missingReqFldCnt == 0).FirstOrDefault();

			sheetLayout.wsLayout.fieldCellMap = fldLayout_v;

			// TODO: this criterion of selection may need to be improved.
			if(sheetLayout.wsLayout.fieldCellMap != null && sheetLayout.wsLayout.fieldColMap != null)
			{
				if (sheetLayout.srcWorksheets == null)
					sheetLayout.srcWorksheets = new List<Worksheet>();

				sheetLayout.srcWorksheets.Add(ws);
				md.isPass = true;
			}
			else
				md.isPass = false;

			return md; 
		}

		private void Init()
		{
			var wsLayout_cga = new WorkSheetLayout
			{
				Name = "Complaint, Grievance and Appeal Information",
				OutputFileName = "Data_Extract_Complaint_Grievance_Appeal_Info_0127",
				fldDelim = "\t",
				recDelim = System.Environment.NewLine,
				dst = this,

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
				dst = this,

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
						Version = 1,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A2", ValueRef = "C2" },
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" }
						}
					},
					new CellLayoutVersion
					{
						Version = 2,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "B3" },
							new CellLocation { TitleRef = "A4", ValueRef = "B4" },
							new CellLocation { TitleRef = "A5", ValueRef = "B5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 3,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "C5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 4,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "C3" },
							new CellLocation { TitleRef = "A4", ValueRef = "C4" },
							new CellLocation { TitleRef = "A5", ValueRef = "D5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 5,
						cellLocations = new List<CellLocation>
						{
							new CellLocation { TitleRef = "A3", ValueRef = "D3" },
							new CellLocation { TitleRef = "A4", ValueRef = "D4" },
							new CellLocation { TitleRef = "A5", ValueRef = "D5" }
						}
					},
					new CellLayoutVersion
					{
						Version = 6,
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

			// Create list of data source types.
			types = new List<SpreadSheetLayout>
			{
				new SpreadSheetLayout
				{
					Name = "Enrollee Complaints, Grievances and Appeals Report (0127)",
					procType = ProcessType.MatchAllDataWorkSheets,
					types = this,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { Name = "Instructions" },
						new SheetLayout { Name = "Codes" },
						new SheetLayout { Name = "January", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "February", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "March", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "April", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "May", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "June", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "July", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "August", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "September", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "October", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "November", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "December", wsLayout = wsLayout_cga },
						new SheetLayout { Name = "Summary" },
						new SheetLayout { Name = "October 2014", isOptional = true, wsLayout = wsLayout_cga },
						new SheetLayout { Name = "November 2014", isOptional = true, wsLayout = wsLayout_cga },
						new SheetLayout { Name = "December 2014", isOptional = true, wsLayout = wsLayout_cga },
					}
				},

				new SpreadSheetLayout
				{
					Name = "Enrollee Roster and Facility Residence Report (0129)",
					procType = ProcessType.MatchByClosestWorkSheetLayout,
					types = this,
					sLayouts = new List<SheetLayout>
					{
						new SheetLayout { wsLayout = wsLayout_erfr }
					}
				}
			};

			this.types.ForEach(ssl => ssl.sLayouts.ForEach(sl => sl.ssLayout = ssl));
		}
	}

	#region Supporting Definitions

	public enum ProcessType
	{
		MatchAllDataWorkSheets,
		MatchByClosestWorkSheetLayout
	}

	class SpreadSheetLayout
	{
		public string Name { get; set; }

		public ProcessType procType;

		public List<SheetLayout> sLayouts { get; set; }

		public DataSourceTypes types { get; set; }

		public void Write()
		{
			sLayouts.Where(s => s.dataSet != null).ToList().ForEach(s => s.dataSet.Write());
		}
	}

	class SheetLayout
	{
		public string Name { get; set; }
		public bool isOptional { get; set; }
		public WorkSheetLayout wsLayout { get; set; }

		public DataSet dataSet { get; set; }

		/// <summary>
		/// Link to matched worksheet in source xlsx file.
		/// </summary>
		public List<Worksheet> srcWorksheets { get; set; }

		public SpreadSheetLayout ssLayout { get; set; }
	}

	public enum DataFormatType
	{
		String = 1,
		DateTime,
		Date
	}

	class WorkSheetLayout
	{
		/// <summary>
		/// May correspond to SpreadSheet WorkSheet name.
		/// </summary>
		public string Name { get; set; }

		public string OutputFileName { get; set; }

		public string fldDelim { get; set; }

		public string recDelim { get; set; }

		public DataSourceTypes dst { get; set; }

		/// <summary>
		/// Versions of sheet cell title and value locations to match for data cell extraction
		/// </summary>
		public List<CellLayoutVersion> cellLayouts { get; set; }

		/// <summary>
		/// Versions of sheet column title cells to match for data column cell extraction
		/// </summary>
		public List<ColumnLayoutVersion> colLayouts { get; set; }
		/// <summary>
		/// List of data columns present on worksheet
		/// </summary>
		public List<Field> fields { get; set; }


		public FieldCellVersionMap fieldCellMap { get; set; }
		public FieldColumnVersionMap fieldColMap { get; set; }
	}

	/// <summary>
	/// 
	/// </summary>
	class ColumnLayoutVersion
	{
		public int Version { get; set; }

		public List<ColumnTitleLocation> titleLocations { get; set; }

		public int FirstRow { get; set; }
	}

	/// <summary>
	/// Location of data column title cells
	/// </summary>
	class ColumnTitleLocation
	{
		public int col { get; set; }

		public List<string> cellRefs { get; set; }
	}

	public enum FieldType
	{
		cell,
		column,
		fileName,
	}

	class Field
	{
		public FieldType fldType { get; set; }

		public string Name { get; set; }

		public int OutputOrder { get; set; }

		public bool isRequired { get; set; }

		public DataFormatType DataFormat { get; set; }

		public List<string> postProcRegex { get; set; }

		public List<string> titles { get; set; }
	}

	/// <summary>
	/// The collection of data field cells locations for a version of the worksheet.
	/// </summary>
	class CellLayoutVersion
	{
		public int Version { get; set; }
		public List<CellLocation> cellLocations { get; set; }
	}

	/// <summary>
	/// Location of a data field cell title and value.
	/// </summary>
	class CellLocation
	{
		public string TitleRef { get; set; }
		public string ValueRef { get; set; }
		public bool isCombined { get; set; }
	}

	class FieldColumnMap
	{
		public int column { get; set; }

		public string title { get; set; }

		public Field field { get; set; }

		public int col_order { get; set; }

		public int field_order { get; set; }
	}

	class FieldColumnVersionMap
	{
		public ColumnLayoutVersion colLayout { get; set; }
		public List<FieldColumnMap> colmaps { get; set; }
		public int noMatchCnt { get; set; }
		public int disOrder { get; set; }
		public int colDups { get; set; }
		public int noneNullTitleCnt { get; set; }
	}

	class FieldCellVersionMap
	{
		public CellLayoutVersion fldLayout { get; set; }
		public List<FieldCellMap> fldmaps { get; set; }
		public int noMatchCnt { get; set; }
		public int noValCnt { get; set; }
		public int missingReqFldCnt { get; set; }
		public int noneNullTitleCnt { get; set; }
	}

	class FieldCellMap
	{
		public string Title { get; set; }

		public string Value { get; set; }

		public CellLocation cellLoc { get; set; }

		public Field field { get; set; }
	}

	#endregion
}
