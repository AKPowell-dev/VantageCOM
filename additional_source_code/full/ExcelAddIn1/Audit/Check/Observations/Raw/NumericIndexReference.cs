using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class NumericIndexReference : RawObservation
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

	public NumericIndexReference(Range rng, string strFunctionName)
		: base(rng)
	{
		FunctionName = strFunctionName;
	}
}
