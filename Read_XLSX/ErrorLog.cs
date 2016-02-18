/*
	© 2016 Florida State University. All rights reserved.

	History
	=========================================================================================================
	2016/02/17	G.K.	Created.

*/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_XLSX
{
	class Log
	{
		private static Log instance = new Log();
		private DateTime timestamp;
		private string filePath;

		private Log()
		{
			timestamp = DateTime.Now;
			filePath = Path.Combine(Directory.GetCurrentDirectory(), $"Extract_XLXS_Data_Log_{timestamp.ToString("yyyyMMdd_HHmmss")}.txt" );
		}

		public static void SetDir(string dir, DateTime stamp)
		{	
			instance.timestamp = stamp;
			instance.filePath = Path.Combine(dir, $"Extract_XLXS_Data_Log_{instance.timestamp.ToString("yyyyMMdd_HHmmss")}.txt");
		}

		public static Log New {  get { return instance; } }

		public void Msg(string msg, string loc = null)
		{
			string m = msg + (!string.IsNullOrWhiteSpace(loc) ? ", loc: " + loc.Trim() : "");
			Console.WriteLine(m);
			var seps = new List<string> { System.Environment.NewLine };
			var strgs = msg.Split(seps.ToArray(), StringSplitOptions.None).ToList();
			strgs = strgs.Select(s => $"\t{s}").ToList();
			if (!string.IsNullOrWhiteSpace(loc))
				strgs.Insert(0, $"******** {loc.Trim()} : {DateTime.Now.ToString()}");
			else
				strgs.Insert(0, $"********* : {DateTime.Now.ToString()}");
			File.AppendAllLines(filePath, strgs);
		}

		public void Msg(Exception ex, string loc = null)
		{
			Msg(ex.Message, loc);
		}
	}
}
