using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_XLSX
{
	class Log
	{
		public static void Msg(string msg, string loc = null)
		{
			Console.WriteLine(msg + (!string.IsNullOrWhiteSpace(loc) ? ", loc: " + loc.Trim() : ""));
		}

		public static void Msg(Exception ex, string loc = null)
		{
			Log.Msg(ex.Message, loc);
		}
	}
}
