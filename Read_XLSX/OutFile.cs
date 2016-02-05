using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_XLSX
{
	class OutFile
	{
		private string _filePath;

		public OutFile(string filePath)
		{
			_filePath = filePath;
		}

		public int WriteDataFolder(DataFolder dfolder)
		{
			StringBuilder outData = new StringBuilder();
			foreach(var dfile in dfolder.dataFiles)
			{

			}
			return 0;
		}
	}
}
