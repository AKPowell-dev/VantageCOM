using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class TooManyOperators : RawObservation
{
	[CompilerGenerated]
	private int A;

	internal int NumberOfOperators
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

	public TooManyOperators(Range rng, int intOperators)
		: base(rng)
	{
		NumberOfOperators = intOperators;
	}
}
