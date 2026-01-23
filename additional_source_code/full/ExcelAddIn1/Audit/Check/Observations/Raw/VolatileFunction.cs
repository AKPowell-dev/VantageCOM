using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class VolatileFunction : RawObservation
{
	[CompilerGenerated]
	private string A;

	internal string FunctionName
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

	public VolatileFunction(Range rng, string strFunction)
		: base(rng)
	{
		FunctionName = strFunction;
	}
}
