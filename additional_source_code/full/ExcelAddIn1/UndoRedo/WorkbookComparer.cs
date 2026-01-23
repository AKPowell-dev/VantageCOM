using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.UndoRedo;

public sealed class WorkbookComparer : IEqualityComparer<Microsoft.Office.Interop.Excel.Workbook>
{
	public bool Equals(Microsoft.Office.Interop.Excel.Workbook x, Microsoft.Office.Interop.Excel.Workbook y)
	{
		return Operators.CompareString(x.FullName, y.FullName, TextCompare: false) == 0;
	}

	bool IEqualityComparer<Microsoft.Office.Interop.Excel.Workbook>.Equals(Microsoft.Office.Interop.Excel.Workbook x, Microsoft.Office.Interop.Excel.Workbook y)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Equals
		return this.Equals(x, y);
	}

	public int GetHashCode(Microsoft.Office.Interop.Excel.Workbook obj)
	{
		return obj.FullName.GetHashCode();
	}

	int IEqualityComparer<Microsoft.Office.Interop.Excel.Workbook>.GetHashCode(Microsoft.Office.Interop.Excel.Workbook obj)
	{
		//ILSpy generated this explicit interface implementation from .override directive in GetHashCode
		return this.GetHashCode(obj);
	}
}
