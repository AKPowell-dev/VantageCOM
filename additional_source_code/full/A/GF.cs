using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class GF : JF
{
	[CompilerGenerated]
	private Action<WorksheetItem, Chart> A;

	internal Action<WorksheetItem, Chart> Query
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	internal GF(string A, string B, Action<WorksheetItem, Chart> C)
		: base(A, B, 0)
	{
		Query = C;
	}
}
