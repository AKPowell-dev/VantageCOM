using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public class RawObservation
{
	[CompilerGenerated]
	private Range A;

	internal Range Range
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

	public RawObservation(Range rng)
	{
		Range = rng;
	}
}
