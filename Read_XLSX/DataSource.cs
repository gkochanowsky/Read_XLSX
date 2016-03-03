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
using System.Text.RegularExpressions;
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

		public FieldColumnVersionMap fldColMap { get; set; }
		public FieldCellVersionMap fldCellMap { get; set; }

		public int matchCnt { get; set; }
	}

	class SpreadSheetLayoutMatch
	{

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
			Config.Load(this);
		}

		private void Clear()
		{
			// clear working references.
			types.ForEach(t => t.sLayouts.ForEach(s =>
			{
				if (s.srcWorksheets != null)
				{
					s.srcWorksheets.Clear();
					s.srcWorksheets = null;
				}

				if (s.dataSet != null)
				{
					if (s.dataSet.Rows != null)
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
					if (s.wsLayout.fieldColMap != null)
					{
						if (s.wsLayout.fieldColMap.colmaps != null)
						{
							s.wsLayout.fieldColMap.colmaps.Clear();
							s.wsLayout.fieldColMap.colmaps = null;
						}

						s.wsLayout.fieldColMap = null;
					}

					if (s.wsLayout.fieldCellMap != null)
					{
						s.wsLayout.fieldCellMap.fldmaps.Clear();
						s.wsLayout.fieldCellMap.fldmaps = null;
						s.wsLayout.fieldCellMap = null;
					}
				}
			}));
		}

		/// <remarks>
		/// Two ways to determin type based on value matchWorkSheetNames
		/// - when true then worksheet names must match DataSource type in order to be a match.
		/// - when false then if any worksheets match data source layout then is a match
		/// </remarks>
		public SpreadSheetLayout DetermineLayout(SpreadsheetDocument ssd, FileInfo file)
		{
			Clear();

			var mds = new List<MatchData>();

			SpreadSheetLayout type = null;

			WorkbookPart wbp = ssd.WorkbookPart;
			var stringTable = wbp.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
			var cellFormats = wbp.WorkbookStylesPart.Stylesheet.CellFormats;

			var shts = wbp.Workbook.Descendants<Sheet>();

			// Add all the match by closest worksheet layouts
			List<SpreadSheetLayout> procTypes = new List<SpreadSheetLayout>();

			foreach(var sl in types.Where(t => t.procType == ProcessType.MatchAllDataWorkSheets))
			{
				var matches = new List<SheetLayoutMap>();

				foreach(var s in shts)
				{
					var shtName = Regex.Replace(s.Name.Value.Replace('\n', ' '), @"\s+", " ").Trim().ToLower();

					var slm = new SheetLayoutMap { sht = s };

					foreach (var slo in sl.sLayouts)
					{
						var nms = slo.Names.Select(n => Regex.Replace(n.Replace('\n', ' '), @"\s+", " ").Trim().ToLower());

						if(nms.Contains(shtName))
						{
							slm.layout = slo;
							break;
						}
					}

					matches.Add(slm);
				}

				var req = sl.sLayouts.Where(slt => !slt.isOptional);
				var mapped = matches.Where(mp => mp.layout != null).Select(mp => mp.layout);
				var missed = req.Where(r => !mapped.Contains(r));

				if (matches.Where(m => m.layout == null).Count() == 0 && missed.Count() == 0)
					procTypes.Add(sl);
			}

			if(procTypes.Count() == 0)
				procTypes.AddRange(types.Where(p => p.procType == ProcessType.MatchByClosestWorkSheetLayout).ToList());

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
							var sheetLayout = dst.sLayouts.FirstOrDefault(sl => sl.Names.Select(n => n.Trim().ToLower()).Contains(sht.Name.Value.ToLower().Trim()));

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
				var cmds = mds.Select(m => new { md = m, noColMatchCntMin = m.fldColVersMaps.Min(cv => cv.noMatchCnt * 10000 + cv.notNullTitleCnt * 100 + cv.disOrder), noCellMatchCntMin = m.fldCellVersMaps.Min(cev => cev.noMatchCnt * 10000 + cev.missingReqFldCnt * 100 + cev.noneNullTitleCnt) });
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
		///		As varied as these spreadsheets may be, a spreadsheet is expected to contain only a single type of related data set which is called a
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
			var md = new MatchData();
			// All cells in worksheet.
			var tcs = ws.Descendants<Cell>();

			switch (sheetLayout.wsLayout.layoutType)
			{
				case LayoutType.Both:
					MatchColLayouts(md, tcs, sheetLayout, stringTable, formats, file);
					MatchCellLayouts(md, tcs, sheetLayout, stringTable, formats, file);
					md.isPass = sheetLayout.wsLayout.fieldCellMap != null && sheetLayout.wsLayout.fieldColMap != null;
					break;
				case LayoutType.CellOnly:
					MatchCellLayouts(md, tcs, sheetLayout, stringTable, formats, file);
					md.isPass = sheetLayout.wsLayout.fieldCellMap != null;
					break;
				case LayoutType.ColumnOnly:
					MatchColLayouts(md, tcs, sheetLayout, stringTable, formats, file);
					md.isPass = sheetLayout.wsLayout.fieldColMap != null;
					break;
			}

			if(md.isPass)
			{
				if (sheetLayout.srcWorksheets == null)
					sheetLayout.srcWorksheets = new List<Worksheet>();

				sheetLayout.srcWorksheets.Add(ws);
			}

			return md;
		}

		public MatchData MatchColLayouts(MatchData md, IEnumerable<Cell> tcs, SheetLayout sheetLayout, SharedStringTablePart stringTable, CellFormats formats, FileInfo file)
		{
			// Find and map columns to sheet layouts
			var field_ord = sheetLayout.wsLayout.fields
							.Where(c => c.fldType == FieldType.column)
							.OrderBy(c => c.OutputOrder)
							.Select(c => c.OutputOrder)
							.ToList();

			// Obtain column titles for all signature versions.
			md.fldColVersMaps = new List<FieldColumnVersionMap>();

			foreach (var sig in sheetLayout.wsLayout.colLayouts.OrderByDescending(scl => scl.titleLocations.Count()))
			// for each column layout version scape the worksheet for column title values
			{
				var fldColMaps = new List<FieldColumnMap>();

				var col_ord = sig.titleLocations.OrderBy(so => so.col).Select(so => so.col).ToList();

				foreach (var colLayout in sig.titleLocations)
				{
					string title = "";
					foreach (var c in colLayout.cellRefs)
					// A column may have a number of title cells that must be scraped and concatinated to product the title used for matching to data columns.
					{
						var cl = tcs.FirstOrDefault(cll => cll.CellReference.InnerText == c);
						var tlt = Spreadsheet.GetCellValue(cl, stringTable.SharedStringTable, formats, null);
						title += tlt;
					}

					title = System.Text.RegularExpressions.Regex.Replace(title.Replace('\n', ' '), @"\s+", " ").Trim().ToLower();
					fldColMaps.Add(new FieldColumnMap { column = colLayout.col, title = title, col_order = col_ord.IndexOf(colLayout.col) });
				}

				md.fldColVersMaps.Add(new FieldColumnVersionMap { colLayout = sig, colmaps = fldColMaps });
			}

			// Match the titles to the DataColumns
			foreach (var fcvm in md.fldColVersMaps)
			{
				foreach (var cm in fcvm.colmaps)
				{
					try
					{
						// TODO: Performance improvement if config strings are pre-processed for whitespace and case.
						cm.field = sheetLayout.wsLayout.fields
										.Where(cc => cc.fldType == FieldType.column && cc.titles != null)
										.FirstOrDefault(cc =>
										{
											var lct = cc.titles.Select(t =>
											{
												var tt = System.Text.RegularExpressions.Regex.Replace(t, @"\s+", " ");
												return tt.ToLower();
											});

											var hasTitle = lct.Contains(cm.title);

											return hasTitle;
										});

						// if required field and should be verified then check that first row has value
						if(sheetLayout.wsLayout.verifyFirstRowData && cm.field != null && cm.field.isRequired )
						{
							int col = fcvm.colLayout.colLayoutType == ColLayoutType.Row_Col ? cm.column : fcvm.colLayout.FirstRow;
							int row = fcvm.colLayout.colLayoutType == ColLayoutType.Row_Col ? fcvm.colLayout.FirstRow : cm.column;

							var valRef = Spreadsheet.GetCellRef(row, col);
							var clVal = tcs.FirstOrDefault(clv => clv.CellReference.InnerText == valRef);

							var val = Spreadsheet.GetCellValue(clVal, stringTable.SharedStringTable, formats, cm.field);

							if (!string.IsNullOrWhiteSpace(val))
							{
								cm.hasValue = true;
								cm.firstRowVal = val;
							}
						}

						cm.field_order = cm.field != null ? field_ord.IndexOf(cm.field.OutputOrder) : -9999;
					}
					catch (Exception ex)
					{
						Log.New.Msg(ex);
					}
				}

				// match by neighbor
				// - Locate flds located by related field
				var flds_byRelated = sheetLayout.wsLayout.fields.Where(f => f.locType == LocateType.byRelated);

				// - Locate column map for related field parent
				var related_pairs = flds_byRelated.Select(fr => new { fr = fr, cm = fcvm.colmaps.FirstOrDefault(rcm => rcm.field != null && rcm.field.OutputOrder == fr.RelatedCol) });

				// - Locate column map for related field
				var rf_cm = related_pairs.Where(rp => rp.cm != null).ToList().Select(rp => new { rp = rp, rc = fcvm.colmaps.FirstOrDefault(fcm => fcm.col_order == rp.cm.col_order + 1) });

				// - Update the field for located column map
				rf_cm.ToList().ForEach(rfcm =>
				{
					rfcm.rc.field = rfcm.rp.fr;
					rfcm.rc.field_order = rfcm.rc.field.OutputOrder;
				});

				fcvm.notNullTitleCnt = fcvm.colmaps.Where(cm => !string.IsNullOrWhiteSpace(cm.title)).Count();
				fcvm.noMatchCnt = fcvm.colmaps.Where(cm => cm.field == null).Count();
				fcvm.ReqNoValCnt = fcvm.colmaps.Where(cm => cm.field != null && cm.field.isRequired && !cm.hasValue).Count();
				fcvm.disOrder = (int)fcvm.colmaps.Where(dm => dm.field != null).Select(dm => Math.Pow((dm.field_order - dm.col_order), 2)).Sum();
				var dupCols = fcvm.colmaps.Where(dm => dm.field != null).GroupBy(cd => cd.field).Where(d => d.Count() > 1);
				fcvm.colDups = dupCols.Count();
			}

			// Only match col layout versions with zero mismatch, favoring the version with the lowest disorder.
			var colLayout_v = md.fldColVersMaps.Where(sv => sv.noMatchCnt == 0 && sv.colDups == 0 && sv.ReqNoValCnt == 0).OrderByDescending(sv => sv.notNullTitleCnt).ThenByDescending(sv => sv.disOrder).FirstOrDefault();

			md.fldColMap = sheetLayout.wsLayout.fieldColMap = colLayout_v;

			md.matchCnt += md.fldColMap != null ? 1 : 0;

			return md;
		}

		public MatchData MatchCellLayouts(MatchData md, IEnumerable<Cell> tcs, SheetLayout sheetLayout, SharedStringTablePart stringTable, CellFormats formats, FileInfo file)
		{
			// Obtain titles for all field cell layouts
			md.fldCellVersMaps = new List<FieldCellVersionMap>();

			foreach (var fldLayout in sheetLayout.wsLayout.cellLayouts)
			{
				var fldLayoutVals = new List<FieldCellMap>();

				foreach (var cellLoc in fldLayout.cellLocations)
				{
					try
					{
						var cl = tcs.FirstOrDefault(cll => cll.CellReference.InnerText == cellLoc.TitleRef);
						var clVal = tcs.FirstOrDefault(clv => clv.CellReference.InnerText == cellLoc.ValueRef);

						var title = Spreadsheet.GetCellValue(cl, stringTable.SharedStringTable, formats, null);
						if (title != null) title = System.Text.RegularExpressions.Regex.Replace(title, @"\s+", " ").Trim().ToLower();
						var val = Spreadsheet.GetCellValue(clVal, stringTable.SharedStringTable, formats, null);

						fldLayoutVals.Add(new FieldCellMap
						{
							cellLoc = cellLoc,
							Title = string.IsNullOrWhiteSpace(title) ? null : title.Trim().ToLower(),
							Value = string.IsNullOrWhiteSpace(val) ? null : val.Trim()
						});
					}
					catch (Exception ex)
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
				foreach (var fm in flvv.fldmaps.Where(m => m.Title != null))
				{
					foreach (var fld in sheetLayout.wsLayout.fields.Where(f => f.fldType == FieldType.cell))
					{
						try
						{
							if (fm.cellLoc.isCombined)
							{
								var titles = fld.titles.Select(t => t.ToLower()).Where(t => fm.Title.StartsWith(t));

								if (titles.Count() > 0)
								{
									fm.field = fld;
									var title = titles.FirstOrDefault();
									if (fm.Value != null && fm.Value.Length > titles.FirstOrDefault().Length)
										fm.Value = fm.Value != null ? fm.Value.Substring(title.Length, fm.Value.Length - title.Length).Trim() : null;

									if (fld.DataFormat == DataFormatType.Date || fld.DataFormat == DataFormatType.DateTime)
									{
										var val = fm.Value.Replace("(", "").Replace(")", "").Replace(":", "");
										if (val.Contains("Through"))
											val = val.Substring(0, val.IndexOf("Through", StringComparison.OrdinalIgnoreCase));
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
								if (fm.Title != null && fld.titles.Select(t => t.ToLower()).Contains(fm.Title))
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
				if (fileFld != null)
				{
					flvv.fldmaps.Add(new FieldCellMap { field = fileFld, Value = file.FullName });
				}

				flvv.noneNullTitleCnt = flvv.fldmaps.Where(fm => !string.IsNullOrWhiteSpace(fm.Title)).Count();
				flvv.noMatchCnt = flvv.fldmaps.Where(fm => fm.field == null).Count();
				flvv.missingReqFldCnt = reqFlds.Where(rf => !flvv.fldmaps.Select(fm => fm.field).Contains(rf)).Count();
				flvv.noValCnt = flvv.fldmaps.Where(fm => fm.field != null && fm.field.isRequired && string.IsNullOrWhiteSpace(fm.Value)).Count();
			}

			var fldLayout_v = md.fldCellVersMaps.Where(fl => fl.noMatchCnt == 0 && fl.noValCnt == 0 && fl.missingReqFldCnt == 0).FirstOrDefault();

			md.fldCellMap = sheetLayout.wsLayout.fieldCellMap = fldLayout_v;

			md.matchCnt += md.fldCellMap != null ? 1 : 0;

			return md;
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
			sLayouts.Where(s => s.dataSet != null).ToList().ForEach(s => s.dataSet.Write(this));
		}
	}

	public enum SheetType
	{
		Ignore,
		CommonData,
		SourceData
	}

	class SheetLayout
	{
		public List<string> Names { get; set; }
		public SheetType sheetType { get; set; }
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

	public enum LayoutType
	{
		ColumnOnly,
		CellOnly,
		Both
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

		public LayoutType layoutType { get; set; }

		public DataSourceTypes dst { get; set; }

		/// <summary>
		/// Versions of sheet cell title and value locations to match for data cell extraction
		/// </summary>
		public List<CellLayoutVersion> cellLayouts { get; set; }

		/// <summary>
		/// Versions of sheet column title cells to match for data column cell extraction
		/// </summary>
		public List<ColumnLayoutVersion> colLayouts { get; set; }

		public bool verifyFirstRowData { get; set; }

		/// <summary>
		/// List of data columns present on worksheet
		/// </summary>
		public List<Field> fields { get; set; }


		public FieldCellVersionMap fieldCellMap { get; set; }
		public FieldColumnVersionMap fieldColMap { get; set; }
	}

	public enum ColLayoutType
	{
		Row_Col,
		Col_Row
	}

	class ColumnLayoutVersion
	{
		public int Version { get; set; }

		public ColLayoutType colLayoutType { get; set; }

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

	public enum LocateType
	{
		byTitle,
		byRelated
	}

	class Field
	{
		public FieldType fldType { get; set; }
		public LocateType locType { get; set; }

		public string Name { get; set; }

		public int OutputOrder { get; set; }

		public int RelatedCol { get; set; }

		public bool isRequired { get; set; }

		public DataFormatType DataFormat { get; set; }

		public List<Tuple<string, string>> postProcRegex { get; set; }

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

		public bool hasValue { get; set; }

		public string firstRowVal { get; set; }
	}

	class FieldColumnVersionMap
	{
		public ColumnLayoutVersion colLayout { get; set; }
		public List<FieldColumnMap> colmaps { get; set; }
		public int noMatchCnt { get; set; }
		public int disOrder { get; set; }
		public int colDups { get; set; }
		public int notNullTitleCnt { get; set; }
		public int ReqNoValCnt { get; set; }
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

	class SheetLayoutMap
	{
		public Sheet sht { get; set; }
		public SheetLayout layout { get; set; }
	}

	#endregion
}
