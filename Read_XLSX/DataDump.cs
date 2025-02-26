﻿/*
	© 2016 Florida State University. All rights reserved.

	DESC: Starts at a root folder and does an inventory of files with xls and xlsx extensions
			- Convert list of xls files that must be transformed to xlsx

	History
	======================================================================================================
	2016/02/05	G.K.	Created.
*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Read_XLSX
{
	class DataDump
	{
		public string RootFolder { get; set; }
		public List<FileInfo> xlsFiles { get; set; }
		public List<FileInfo> xlsxFiles { get; set; }
		public List<FileInfo> zipFiles { get; set; }

		public DataSourceTypes _dsts;

		public DataDump(string rootFolder)
		{
			if (!Directory.Exists(rootFolder))
			{
				Log.New.Msg($"Root directory: {rootFolder} does not exist");
				return;
			}

			RootFolder = rootFolder;
			_dsts = new DataSourceTypes(rootFolder);
		}

		public int Scan()
		{
			if (RootFolder == null)
				return 0;

			xlsFiles = new List<FileInfo>();
			xlsxFiles = new List<FileInfo>();
			zipFiles = new List<FileInfo>();

			ScanDirectoriesRecursive(RootFolder);

			return xlsFiles.Count() + xlsxFiles.Count() + zipFiles.Count(); ;
		}

		private void ScanDirectoriesRecursive(string folder)
		{
			var dir = new DirectoryInfo(folder);

			// Ignore files that appear to be opened with excel.
			var xlsFileInfo = dir.EnumerateFiles().Where(f => f.Extension == ".xls" && !f.Name.StartsWith("~"));
			var xlsxFileInfo = dir.EnumerateFiles().Where(f => f.Extension == ".xlsx" && !f.Name.StartsWith("~"));
			var zipFileInfo = dir.EnumerateFiles().Where(f => f.Extension == ".zip");

			xlsFiles.AddRange(xlsFileInfo);
			xlsxFiles.AddRange(xlsxFileInfo);
			zipFiles.AddRange(zipFileInfo);

			dir.GetDirectories().ToList().ForEach(d => ScanDirectoriesRecursive(d.FullName));
		}

		public void ProcessDataDump()
		{
			// scan dump for xls files and convert to xlsx
			ConvertFiles();

			// Determine DataSourceType and extract data from all xlsx files.
			ExtractData();
		}

		private void ConvertFiles()
		{
			Scan();

			// Convert all XLS file to XLSX.
			ConvertXLS();

			// Rescan
			Scan();
		}

		private void ExtractData()
		{
			var ss = new Spreadsheet(_dsts);
			int recCnt = 0;

			// Process XLSX files.
			foreach (var file in xlsxFiles)
			{
				recCnt += ss.ProcessFile(file);
			}

			Log.New.Msg($"Processed {recCnt} records");
		}

		public int ConvertXLS()
		{
			// Create list of xls files to be converted.
			var xlsNames = xlsFiles.Select(f => new { file = f, name = f.Name.Split('.')[0] });
			var xlsxNames = xlsxFiles.Select(f => f.Name.Split('.')[0]);
			var toConvert = xlsNames.Where(x => !xlsxNames.Contains(x.name)).Select(f => f.file.FullName).ToList();

			return ConvertXLStoXLSX(toConvert);
		}

		private int ConvertXLStoXLSX(List<string> files)
		{
			var app = new Excel.Application();

			int cnt = 0;
			foreach (string file in files)
			{
				if (!File.Exists(file))
				{
					Log.New.Msg($"XLS file: {file} could not be converted to XLSX because it doesn't exist.");
					continue;
				}

				var wb = app.Workbooks.Open(file);
				wb.SaveAs(Filename: file + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
				wb.Close();
				++cnt;
				Log.New.Msg($"Converted: {file} to .xlsx");
			}

			app.Quit();
			return cnt;
		}
	}
}
