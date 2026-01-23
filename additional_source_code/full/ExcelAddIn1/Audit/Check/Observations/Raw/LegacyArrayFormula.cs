using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class LegacyArrayFormula : RawObservation
{
	[CompilerGenerated]
	private string A;

	internal string Formula
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

	public LegacyArrayFormula(Range rng)
		: base(rng)
	{
		Formula = Conversions.ToString(NewLateBinding.LateGet(rng, null, VH.A(8714), new object[0], null, null, null));
	}
}
