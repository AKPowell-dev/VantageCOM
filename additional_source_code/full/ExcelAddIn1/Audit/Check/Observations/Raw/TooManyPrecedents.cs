using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class TooManyPrecedents : RawObservation
{
	[CompilerGenerated]
	private int A;

	internal int NumberOfPrecedents
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

	public TooManyPrecedents(Range rng, int intPrecedents)
		: base(rng)
	{
		NumberOfPrecedents = intPrecedents;
	}
}
