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
			//			var tst = Spreadsheet.GetCellRef(7, 28);

			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_me";

						string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_mccma";

			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_erfr\";
			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_ntf";
			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_co";
			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_cga\";
			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_all";
			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_cor";
			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_pdo";




			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_roster\";
			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\FILES TO IMPORT\From Houser - Copy\Community Outreach Health FairsPublic Events Notification (0113)";
			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\FILES TO IMPORT\From Houser - Copy\Enrollee Roster and Facility Residence Report (0129)";
			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\FILES TO IMPORT\From Houser - Copy\Missed Services Report (0131)";

			var startDT = DateTime.Now;
			Log.SetDir(folder, startDT);

			Log.New.Msg($"Started {startDT.ToString()}");

			var dd = new DataDump(folder);
			dd.ProcessDataDump();

			var endDT = DateTime.Now;

			Log.New.Msg($"Ended {endDT.ToString()}");

			var elapsed = endDT - startDT;

			Log.New.Msg($"Elapsed time: {elapsed.Hours.ToString("00")}:{elapsed.Minutes.ToString("00")}:{elapsed.Seconds.ToString("00")} hh:mm:ss");

			System.Console.WriteLine("Press any key to exit...");
			System.Console.ReadLine();
		}
	}
}
