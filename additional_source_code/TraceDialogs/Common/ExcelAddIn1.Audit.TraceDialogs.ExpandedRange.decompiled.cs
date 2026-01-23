using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.TraceDialogs;

public sealed class ExpandedRange
{
	[CompilerGenerated]
	private Range A;

	[CompilerGenerated]
	private bool A;

	[CompilerGenerated]
	private bool B;

	public Range Range
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	public bool RowsExpanded
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

	public bool ColumnsExpanded
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public ExpandedRange(Range rng, bool blnRowsExpanded, bool blnColumnsExpanded)
	{
		Range = rng;
		RowsExpanded = blnRowsExpanded;
		ColumnsExpanded = blnColumnsExpanded;
	}
}
