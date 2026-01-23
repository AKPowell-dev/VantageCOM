using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.Check.Observations.Raw;

public sealed class CellFillColor : RawObservation
{
	[CompilerGenerated]
	private int A;

	internal int OleColor
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

	public CellFillColor(Range rng, int intOleColor)
		: base(rng)
	{
		OleColor = intOleColor;
	}
}
