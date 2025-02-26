﻿/*
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
					s.dataSet.sLayout.matchData = null;
					s.dataSet.sLayout = null;
					s.dataSet = null;
				}


			}));
		}

		/// <remarks>
		/// Two ways to determin type based on value matchWorkSheetNames
		/// - when true then worksheet names must match DataSource type in order to be a match.
		/// - when false then if any worksheets match data source layout then is a match
		/// </remarks>
		public SpreadSheetLayout DetermineLayout(SpreadsheetDocument ssd, FileInfo file, SpreadSheetLayout lastSSL)
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

			foreach (var sl in types.Where(t => t.procType == ProcessType.MatchAllDataWorkSheets))
			{
				var matches = new List<SheetLayoutMap>();

				foreach (var s in shts)
				{
					var shtName = Regex.Replace(s.Name.Value.Replace('\n', ' '), @"\s+", " ").Trim().ToLower();

					var slm = new SheetLayoutMap { sht = s };

					foreach (var slo in sl.sLayouts)
					{
						var nms = slo.Names.Select(n => Regex.Replace(n.Replace('\n', ' '), @"\s+", " ").Trim().ToLower());

						if (nms.Contains(shtName))
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
				var nullLayts = matches.Where(mp => mp.layout == null);

				var missedCnt = missed.Count();
				var nullLaytsCnt = nullLayts.Count();

				if (nullLaytsCnt == 0 && missedCnt == 0)
					procTypes.Add(sl);
			}

			if (procTypes.Count() == 0)
				procTypes.AddRange(types.Where(p => p.procType == ProcessType.MatchByClosestWorkSheetLayout).ToList());

			if (procTypes.Count() == 0) return null;

			if (lastSSL != null)
			// There was a last spread sheet layout, try that first.
			{
				var sortProcTypes = procTypes.Select(p => new { procType = p, ordIdx = (p == lastSSL ? 1 : 2) }).OrderBy(n => n.ordIdx);
				procTypes = sortProcTypes.Select(s => s.procType).ToList();
			}

			foreach (var dst in procTypes)
			// Iterate through types
			{
				bool isPass = true;

				switch (dst.procType)
				{
					case ProcessType.MatchAllDataWorkSheets:

						foreach (var sht in wbp.Workbook.Descendants<Sheet>())
						{
							var sheetLayout = dst.sLayouts.FirstOrDefault(sl => sl.Names.Select(n => n.Trim().ToLower()).Contains(sht.Name.Value.ToLower().Trim()));

							if (sheetLayout == null || sheetLayout.wsLayout == null) continue;

							WorksheetPart wsp = wbp.GetPartById(sht.Id) as WorksheetPart;

							var md = MatchLayouts(wsp.Worksheet, sheetLayout, stringTable, cellFormats, file);
							sheetLayout.matchData = md;

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
								sheetLayout.matchData = md;
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

			if (type == null)
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
			MatchData md = new MatchData();
			MatchData colMD = null;
			MatchData cellMD = null;

			// All cells in worksheet.
			var tcs = ws.Descendants<Cell>();

			switch (sheetLayout.wsLayout.layoutType)
			{
				case LayoutType.Both:
					colMD = MatchColLayouts(md, tcs, sheetLayout, stringTable, formats, file);
					cellMD = MatchCellLayouts(md, tcs, sheetLayout, stringTable, formats, file);
					md.fldCellMap = cellMD.fldCellMap;
					md.fldColMap = colMD.fldColMap;
					md.isPass = md.fldColMap != null && md.fldCellMap != null;
					break;
				case LayoutType.CellOnly:
					cellMD = MatchCellLayouts(md, tcs, sheetLayout, stringTable, formats, file);
					md.fldCellMap = cellMD.fldCellMap;
					md.isPass =  md.fldCellMap != null;
					break;
				case LayoutType.ColumnOnly:
					colMD = MatchColLayouts(md, tcs, sheetLayout, stringTable, formats, file);
					md.fldColMap = colMD.fldColMap;
					md.isPass = md.fldColMap != null;
					cellMD = FileFields(md, sheetLayout, file);
					md.fldCellMap = cellMD.fldCellMap;
					break;
			}

			if (md.isPass)
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
						if (sheetLayout.wsLayout.verifyFirstRowData && cm.field != null && cm.field.isRequired)
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
				// - Locate flds located by related adjacent field
				var flds_byRelated = sheetLayout.wsLayout.fields.Where(f => f.locType == LocateType.byRelated);

				if (flds_byRelated.Count() > 0 && fcvm.colmaps.Where(cm => cm.field != null).Count() > 0)
				// There are related fields in the wsLayout and maps to check by.
				{
					// Look for related fields where there is a map for the related field but no match for the column adjacent to the related field.
					var related_flds = from fbr in flds_byRelated
									   from m in fcvm.colmaps.Where(fcv => fcv.field != null)
									   from um in fcvm.colmaps.Where(fcv => fcv.field == null)
									   where fbr.RelatedOutputOrder == m.field.OutputOrder && m.col_order + 1 == um.col_order
									   select new { related_fld = fbr, cm = m, um = um };

					// Update the field map of the unmapped adjacent related field.
					related_flds.ToList().ForEach(rf => rf.um.field = rf.related_fld);
				}

				fcvm.notNullTitleCnt = fcvm.colmaps.Where(cm => !string.IsNullOrWhiteSpace(cm.title)).Count();
				fcvm.noMatchCnt = fcvm.colmaps.Where(cm => cm.field == null).Count();
				fcvm.ReqNoValCnt = fcvm.colmaps.Where(cm => cm.field != null && sheetLayout.wsLayout.verifyFirstRowData && cm.field.isRequired && !cm.hasValue).Count();
				fcvm.disOrder = (int)fcvm.colmaps.Where(dm => dm.field != null).Select(dm => Math.Pow((dm.field_order - dm.col_order), 2)).Sum();
				var dupCols = fcvm.colmaps.Where(dm => dm.field != null).GroupBy(cd => cd.field).Where(d => d.Count() > 1);
				fcvm.colDups = dupCols.Count();
			}

			// Only match col layout versions with zero mismatch, favoring the version with the lowest disorder.
			var colLayout_v = md.fldColVersMaps.Where(sv => sv.noMatchCnt == 0 && sv.colDups == 0 && (!sheetLayout.wsLayout.verifyFirstRowData || sv.ReqNoValCnt == 0)).OrderByDescending(sv => sv.notNullTitleCnt).ThenByDescending(sv => sv.disOrder).FirstOrDefault();

			md.fldColMap = colLayout_v;

			md.matchCnt += md.fldColMap != null ? 1 : 0;

			return md;
		}

		public MatchData FileFields(MatchData md, SheetLayout sheetLayout, FileInfo file)
		{
			// Obtain titles for all field cell layouts
			md.fldCellVersMaps = new List<FieldCellVersionMap>();

			if (sheetLayout.wsLayout.cellLayouts == null)
			{
				sheetLayout.wsLayout.cellLayouts = new List<CellLayoutVersion>();
				sheetLayout.wsLayout.cellLayouts.Add(new CellLayoutVersion { Version = 1 });
			}

			foreach(var fldLayout in sheetLayout.wsLayout.cellLayouts)
			{
				var fldLayoutVals = new List<FieldCellMap>();

				md.fldCellVersMaps.Add(new FieldCellVersionMap { fldmaps = fldLayoutVals, fldLayout = fldLayout });

				foreach (var flvv in md.fldCellVersMaps)
				{
					flvv.fldmaps = new List<FieldCellMap>();

					md.fldCellMap = flvv;

					// Add a filename layout if the field exists.
					var fileName = sheetLayout.wsLayout.fields.FirstOrDefault(fld => fld.fldType == FieldType.fileName);
					if (fileName != null)
					{
						flvv.fldmaps.Add(new FieldCellMap { Title = FieldType.fileName.ToString(), field = fileName, Value = file.Name });
					}

					// Add a filePath layout if the field exists.
					var filePath = sheetLayout.wsLayout.fields.FirstOrDefault(fld => fld.fldType == FieldType.filePath);
					if (filePath != null)
					{
						flvv.fldmaps.Add(new FieldCellMap { Title = FieldType.filePath.ToString(), field = filePath, Value = file.FullName });
					}
				}
			}

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

			md.fldCellVersMaps.ForEach(vm => vm.fldmaps.ForEach(vfm => vfm.versMap = vm));

			var reqFlds = sheetLayout.wsLayout.fields.Where(sf => sf.fldType == FieldType.cell && sf.isRequired);

			// Match Titles to layout fields
			foreach (var flvv in md.fldCellVersMaps)
			{
				var agMaps = new List<FieldCellMap>();

				foreach (var fm in flvv.fldmaps.Where(m => m.Title != null))
				{
					MatchField(tcs, fm, sheetLayout, stringTable, formats, agMaps);
				}

				flvv.fldmaps.AddRange(agMaps);

				// Only keep data containing CellDataLayouts
				var fmts = new List<CellDataLayout> { CellDataLayout.combined, CellDataLayout.separate };
				flvv.fldmaps = flvv.fldmaps.Where(fm => fm.cellLoc != null && fmts.Contains(fm.cellLoc.dataLayout)).ToList();

				// Add a filename layout if the field exists.
				var fileName = sheetLayout.wsLayout.fields.FirstOrDefault(fld => fld.fldType == FieldType.fileName);
				if (fileName != null)
				{
					flvv.fldmaps.Add(new FieldCellMap { Title = FieldType.fileName.ToString(), field = fileName, Value = file.Name });
				}

				// Add a filePath layout if the field exists.
				var filePath = sheetLayout.wsLayout.fields.FirstOrDefault(fld => fld.fldType == FieldType.filePath);
				if (filePath != null)
				{
					flvv.fldmaps.Add(new FieldCellMap { Title = FieldType.filePath.ToString(), field = filePath, Value = file.FullName });
				}

				// Compute how well the matching went.
				flvv.noneNullTitleCnt = flvv.fldmaps.Where(fm => !string.IsNullOrWhiteSpace(fm.Title)).Count();
				flvv.noMatchCnt = flvv.fldmaps.Where(fm => fm.field == null).Count();
				flvv.missingReqFldCnt = reqFlds.Where(rf => !flvv.fldmaps.Select(fm => fm.field).Contains(rf)).Count();
				flvv.noReqValCnt = flvv.fldmaps.Where(fm => fm.field != null && fm.field.isRequired && string.IsNullOrWhiteSpace(fm.Value)).Count();
				flvv.noValCnt = flvv.fldmaps.Where(fm => fm.field != null && string.IsNullOrWhiteSpace(fm.Value)).Count();
			}

			// Find the best acceptable layout match.
			var fldLayout_v = md.fldCellVersMaps
									.Where(fl => fl.noMatchCnt == 0 && fl.noReqValCnt == 0 && fl.missingReqFldCnt == 0)
									.OrderByDescending(fl => fl.fldmaps.Count())
									.OrderBy(fl => fl.noValCnt)
									.FirstOrDefault();

			var hasRequiredCellsFlds = sheetLayout.wsLayout.fields.Where(f => f.isRequired && f.fldType == FieldType.cell).Count() > 0;

			if (fldLayout_v == null && !hasRequiredCellsFlds)
			// No cell layout match but no required cell fields
			{
				// Create a layout for file fields if they are there.
				fldLayout_v = FileFields(md, sheetLayout, file).fldCellMap;
			}

			md.fldCellMap = fldLayout_v;

			md.matchCnt += md.fldCellMap != null ? 1 : 0;

			return md;
		}

		private bool MatchField(IEnumerable<Cell> tcs, FieldCellMap fm, SheetLayout sheetLayout, SharedStringTablePart stringTable, CellFormats formats, List<FieldCellMap> agMaps)
		{
			bool foundFld = false;

			foreach (var fld in sheetLayout.wsLayout.fields.Where(f => f.fldType == FieldType.cell))
			{
				if (foundFld)
					break;

				try
				{
					switch (fm.cellLoc.dataLayout)
					{
						case CellDataLayout.combined:
							foundFld = MatchCombinedField(fld, fm);
							break;
						case CellDataLayout.separate:
							foundFld = MatchSeparatedField(tcs, fld, fm, stringTable, formats);
							break;
						case CellDataLayout.aggregate:
							// Cell contains an aggregate of fields. Based on 
							{
								var aggCell = tcs.Where(c => c.CellReference == fm.cellLoc.TitleRef).FirstOrDefault();
								var val = Spreadsheet.GetCellValue(aggCell, stringTable.SharedStringTable, formats, null);
								if (val != null && fm.cellLoc.aggregateCellCnt > 0 && fm.cellLoc.aggregateCellSeparator != null && fm.cellLoc.aggregateCellSeparator.Count() > 0)
								{
									var cells = Regex.Split(val.Trim(), fm.cellLoc.aggregateCellSeparator).ToList();
									if (cells.Count() == fm.cellLoc.aggregateCellCnt)
									{
										var nfms = new List<FieldCellMap>();

										foreach(var agv in fm.cellLoc.cellMaps)
										{
											FieldCellMap nfm = null;

											switch (agv.dataLayout)
											{
												case CellDataLayout.combined:
													nfm = new FieldCellMap {
														Title = cells[agv.aggregateIdx].ToLower(),
														Value = cells[agv.aggregateIdx],
														versMap = fm.versMap,
														cellLoc = new CellLocation
														{
															dataLayout = CellDataLayout.combined,
															TitleRef = fm.cellLoc.TitleRef,
															ValueRef = fm.cellLoc.ValueRef
														} 
													};

													MatchField(tcs, nfm, sheetLayout, stringTable, formats, agMaps);
													if (nfm.field != null)
														nfms.Add(nfm);

													break;

												case CellDataLayout.lookup:
													var nfld1 = sheetLayout.wsLayout.fields.Where(f => f.fldType == FieldType.cell);
													var nfld2 = nfld1.Where(f => f.titles != null);

													var nfld = nfld2.FirstOrDefault(f => f.titles.Select(t => t.ToLower()).Contains(agv.lookupString.ToLower()));
													
													if (nfld != null)
													{
														nfm = new FieldCellMap
														{
															Title = agv.lookupString,
															Value = cells[agv.aggregateIdx],
															field = nfld,
															versMap = fm.versMap,
															cellLoc = new CellLocation
															{
																dataLayout = CellDataLayout.separate,
																TitleRef = fm.cellLoc.TitleRef,
																ValueRef = fm.cellLoc.ValueRef
															}
														};

														nfms.Add(nfm);
													}
													break;
											}
										}

										if(nfms.Count == fm.cellLoc.cellMaps.Count())
										{
											agMaps.AddRange(nfms);
											foundFld = true;
											break;
										}
									}
								}
							}
							break;
					}


				}
				catch (Exception ex)
				{
					Log.New.Msg(ex);
				}

			}
			return foundFld;
		}

		private bool MatchCombinedField(Field fld, FieldCellMap fm)
		{
			var titles = fld.titles.Select(t => t.ToLower()).Where(t =>
			{
				bool hasTitle = fm.Title.StartsWith(t);
				return hasTitle;
			});

			if (titles.Count() > 0)
			{
				fm.field = fld;
				var title = titles.FirstOrDefault();
				if (fm.Value != null && fm.Value.Length >= titles.FirstOrDefault().Length)
					fm.Value = fm.Value != null ? fm.Value.Substring(title.Length).Trim() : null;

				if (fld.DataFormat == DataFormatType.Date || fld.DataFormat == DataFormatType.DateTime || fld.DataFormat == DataFormatType.DateMixed)
				{
					var val = fm.Value.Replace("(", "").Replace(")", "").Replace(":", "");
					if (val.Contains("Through"))
						val = val.Substring(0, val.IndexOf("Through", StringComparison.OrdinalIgnoreCase));
					DateTime outVal;
					if (DateTime.TryParse(val, out outVal))
					{
						if (fld.DataFormat == DataFormatType.Date || fld.DataFormat == DataFormatType.DateMixed)
							fm.Value = outVal.ToString("MM/dd/yyyy");
						else
							fm.Value = outVal.ToString();
					}
					else
						fm.Value = fld.DataFormat == DataFormatType.DateMixed ? fm.Value : null;
				}

				if (string.IsNullOrWhiteSpace(fm.Value))
					fm.Value = null;

				return true;
			}

			return false;
		}

		private bool MatchSeparatedField(IEnumerable<Cell> tcs, Field fld, FieldCellMap fm, SharedStringTablePart stringTable, CellFormats formats)
		{
			if (fm.Title != null && fld.titles.Select(t => t.ToLower()).Contains(fm.Title))
			{
				fm.field = fld;

				if (fld.DataFormat == DataFormatType.Date || fld.DataFormat == DataFormatType.DateTime || fld.DataFormat == DataFormatType.DateMixed)
				{
					var cell = tcs.Where(c => c.CellReference == fm.cellLoc.ValueRef).FirstOrDefault();
					fm.Value = Spreadsheet.GetCellValue(cell, stringTable.SharedStringTable, formats, fld);
				}

				return true;
			}

			return false;
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

		public MatchData matchData { get; set; }
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
		Date,
		DateMixed,
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


//		public FieldCellVersionMap fieldCellMap { get; set; }
//		public FieldColumnVersionMap fieldColMap { get; set; }
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

		public bool isGroupData { get; set; }

		public CellDataLayout dataLayout { get; set; }
	}

	public enum FieldType
	{
		cell,
		column,
		fileName,
		filePath,
	}

	public enum LocateType
	{
		byTitle,
		byRelated
	}

	public enum RowType
	{
		RowData,
		GroupData // Match with fields that match col data for columns with cell title locations that have isGroupData = true
	}

	class Field
	{
		public FieldType fldType { get; set; }

		public LocateType locType { get; set; }

		public RowType rowType { get; set; }

		public string Name { get; set; }

		public int OutputOrder { get; set; }

		public int RelatedOutputOrder { get; set; }

		public bool isRequired { get; set; }

		public DataFormatType DataFormat { get; set; }

		public List<Tuple<string, string>> postProcRegex { get; set; }

		public List<string> titles { get; set; }

		public List<string> ignore { get; set; }
	}

	/// <summary>
	/// The collection of data field cells locations for a version of the worksheet.
	/// </summary>
	class CellLayoutVersion
	{
		public int Version { get; set; }
		public List<CellLocation> cellLocations { get; set; }
	}

	public enum CellDataLayout
	{
		separate,
		combined,
		aggregate,
		lookup,
	}

	class AggregateFieldCellMap
	{
		public int aggregateIdx { get; set; }

		public CellDataLayout dataLayout { get; set; }

		public string lookupString { get; set; }

	}

	/// <summary>
	/// Location of a data field cell title and value.
	/// </summary>
	class CellLocation
	{
		public string TitleRef { get; set; }
		public string ValueRef { get; set; }

		public CellDataLayout dataLayout { get; set; }
		public int aggregateCellCnt { get; set; }
		public string aggregateCellSeparator { get; set; }

		public List<AggregateFieldCellMap> cellMaps { get; set; }

//		public CellLayoutVersion cellLayoutVersion { get; set; }

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
		public int noReqValCnt { get; set; }
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

		public FieldCellVersionMap versMap { get; set; }
	}

	class SheetLayoutMap
	{
		public Sheet sht { get; set; }
		public SheetLayout layout { get; set; }
	}

	#endregion
}
