using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.SuperFind2.Queries;
using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class TF : BaseQuery
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

	internal TF(GF A)
		: base(A.UniqueId)
	{
		Query = A.Query;
	}

	internal TF(string A, Action<WorksheetItem, Chart> B)
		: base(A)
	{
		Query = B;
	}
}
