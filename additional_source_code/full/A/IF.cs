using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class IF : JF
{
	[CompilerGenerated]
	private Action<WorksheetItem, Range> A;

	internal Action<WorksheetItem, Range> Query
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

	internal IF(string A, string B, Action<WorksheetItem, Range> C)
		: base(A, B, 0)
	{
		Query = C;
	}

	internal IF(string A, string B, Action<WorksheetItem, Range> C, int D, string E)
		: base(A, B, D)
	{
		Query = C;
		base.PlaceholderText = E;
	}
}
