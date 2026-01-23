using System;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.SuperFind2.Queries;

public sealed class WorksheetQuery : BaseQuery
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

	internal WorksheetQuery(LF A)
		: base(A.UniqueId)
	{
		Query = A.Query;
	}

	internal WorksheetQuery(string A, Action<WorksheetItem, Worksheet> B)
		: base(A)
	{
		Query = B;
	}
}
