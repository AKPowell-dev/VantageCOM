using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class DeepNesting : RawObservation
{
	[CompilerGenerated]
	private int A;

	internal int NestingLevel
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

	public DeepNesting(Range rng, int intLevels)
		: base(rng)
	{
		NestingLevel = intLevels;
	}
}
