/*
	© Florida State University. All rights reserved.

	TODO:
			- Add conversion of binary format files to XML format.
	
	History
	===========================================================================================
	2016/02/03	G.K.	Created.
*/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_XLSX
{
	class DataFolder
	{
		public string _folderPath { get; set; }
		public List<DataFile> dataFiles { get; set; }

		public DataFolder(string folderPath)
		{
			if (!Directory.Exists(folderPath))
				throw new Exception($"DataFolder: {folderPath} is invalid.");
			_folderPath = folderPath;
		}

		public int ProcessFolder()
		{
			var files = new DirectoryInfo(_folderPath).GetFiles().Where(f => f.Extension == ".xlsx" && !f.Name.StartsWith("~"));

			if (dataFiles == null) dataFiles = new List<DataFile>();

			foreach(var file in files)
			{
				var df = Spreadsheet.ProcessFile(file.FullName);
				if (df != null) dataFiles.Add(df);
			}

			return dataFiles.Count(); 
		}

		public int RecCount()
		{
			return dataFiles.Sum(d => d.RecCount());
		}

		private string GetDelimitedData()
		{
			StringBuilder sb = new StringBuilder();
			dataFiles.ForEach(f => f.GetDelimitedRows(sb, "\t", System.Environment.NewLine));
			return sb.ToString();
		}

		public void WriteData()
		{
			//string dir Path.GetDirectoryName(_folderPath).Split('\\');
			var dir = Directory.GetParent(_folderPath).Name;
			var outFileName = $"Enrollee_Complaints_Grievances_Appeals_Report_0127_{dir}_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.txt";
			var outPath = Path.Combine(_folderPath, outFileName);

			if (File.Exists(outPath))
				File.Delete(outPath);

			var data = GetDelimitedData();
			File.WriteAllText(outPath, data, Encoding.ASCII);
		}
	}
}
