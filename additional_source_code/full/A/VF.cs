using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.SuperFind2.Queries;
using ExcelAddIn1.SuperFind2.Results;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class VF : BaseQuery
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

	internal VF(IF A)
		: base(A.UniqueId)
	{
		Query = A.Query;
	}

	internal VF(string A, Action<WorksheetItem, Range> B)
		: base(A)
	{
		Query = B;
	}

	internal VF(IF A, string B, string C, bool D, bool E, bool F, bool G, bool H, bool I, bool J)
		: base(A.UniqueId)
	{
		Query = A.Query;
		base.LookInComments = E;
		base.LookInCharts = F;
		base.LookInEmptyCells = G;
		base.LookInFormulas = H;
		base.LookInHyperlinks = I;
		base.LookInValues = J;
		base.Input1 = B;
		base.Input2 = C;
		base.MatchCase = D;
	}
}
