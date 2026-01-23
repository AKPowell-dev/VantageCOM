using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class TooManyGroupings : RawObservation
{
	[CompilerGenerated]
	private int A;

	internal int NumberOfGroupings
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

	public TooManyGroupings(Range rng, int intGroupings)
		: base(rng)
	{
		NumberOfGroupings = intGroupings;
	}
}
