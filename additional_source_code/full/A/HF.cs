using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.SuperFind2.Results;

namespace A;

internal sealed class HF : JF
{
	[CompilerGenerated]
	private Action<WorksheetItem, object> A;

	internal Action<WorksheetItem, object> Query
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

	internal HF(string A, string B, Action<WorksheetItem, object> C)
		: base(A, B, 0)
	{
		Query = C;
	}

	internal HF(string A, string B, Action<WorksheetItem, object> C, int D)
		: base(A, B, D)
	{
		Query = C;
	}

	internal HF(string A, string B, Action<WorksheetItem, object> C, int D, string E)
		: base(A, B, D)
	{
		Query = C;
		base.PlaceholderText = E;
	}
}
