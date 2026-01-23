using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class ApproximateMatch : RawObservation
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

	public ApproximateMatch(Range rng, string strFunctionName)
		: base(rng)
	{
		FunctionName = strFunctionName;
	}
}
