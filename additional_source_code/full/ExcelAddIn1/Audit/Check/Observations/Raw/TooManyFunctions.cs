using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class TooManyFunctions : RawObservation
{
	[CompilerGenerated]
	private int A;

	internal int NumberOfFunctions
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

	public TooManyFunctions(Range rng, int intFunctions)
		: base(rng)
	{
		NumberOfFunctions = intFunctions;
	}
}
