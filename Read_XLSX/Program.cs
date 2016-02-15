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
						string folder = @"D:\local\CPDC\Projects\Read_XLSX\test_one\";
			// string folder = @"D:\local\CPDC\Projects\Read_XLSX\test\";
			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\FILES TO IMPORT\From Houser - Copy\Community Outreach Health FairsPublic Events Notification (0113)";

			//			string folder = @"D:\local\CPDC\Projects\Read_XLSX\FILES TO IMPORT\From Houser - Copy\Enrollee Roster and Facility Residence Report (0129)";

			var dd = new DataDump(folder);
			dd.ProcessDataDump();

			System.Console.WriteLine("Press any key to exit...");
			System.Console.ReadLine();
		}
	}
}
