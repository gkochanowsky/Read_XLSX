/*
	© 2016 Florida State University. All rights reserved.

	History
	===================================================================================================
	2016/02/03	G.K.	Created.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_XLSX
{
	class Program
	{
		static void Main(string[] args)
		{
//			var dts = new DataSourceTypes();

			string folder = @"D:\local\CPDC\Projects\Read_XLSX\test\";

			var dd = new DataDump(folder);
			int fileCnt = dd.Scan();
			fileCnt = dd.ProcessDataDump();

			//			var df = new DataFolder(folder);
			//			var files = df.ProcessFolder();
			//			Log.Msg($"Total Records: {df.RecCount()} over {files} files");
			//			df.WriteData();
			System.Console.WriteLine("Press any key to exit...");
			System.Console.ReadLine();
		}
	}
}
