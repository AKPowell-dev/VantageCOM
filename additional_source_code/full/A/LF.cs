using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class LF : JF
{
	[CompilerGenerated]
	private Action<WorksheetItem, Worksheet> A;

	internal Action<WorksheetItem, Worksheet> Query
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

	internal LF(string A, string B, Action<WorksheetItem, Worksheet> C)
		: base(A, B, 0)
	{
		Query = C;
	}
}
