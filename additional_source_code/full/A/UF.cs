using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.SuperFind2.Queries;
using ExcelAddIn1.SuperFind2.Results;

namespace A;

internal sealed class UF : BaseQuery
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

	internal UF(HF A)
		: base(A.UniqueId)
	{
		Query = A.Query;
	}

	internal UF(string A, Action<WorksheetItem, object> B)
		: base(A)
	{
		Query = B;
	}

	internal UF(HF A, string B, string C, bool D, bool E, bool F, bool G, bool H, bool I, bool J)
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
