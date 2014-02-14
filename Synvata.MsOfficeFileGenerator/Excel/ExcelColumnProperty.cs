using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace Synvata.MsOfficeFileGenerator.Excel
{
	public class ExcelColumnProperty
	{
		public PropertyInfo PropertyInfo { get; set; }
		public ExcelColumnAttribute ExcelColumnAttr { get; set; }
		public int CellFormatIndex { get; set; }

		public ExcelColumnProperty()
		{
			CellFormatIndex = -1;
		}
	}
}
