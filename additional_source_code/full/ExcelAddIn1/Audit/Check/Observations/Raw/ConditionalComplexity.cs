using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class ConditionalComplexity : RawObservation
{
	[CompilerGenerated]
	private int A;

	internal int NumberOfConditions
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

	public ConditionalComplexity(Range rng, int intIfs)
		: base(rng)
	{
		NumberOfConditions = intIfs;
	}
}
