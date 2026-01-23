using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.UndoRedo;

public sealed class SavedCell
{
	[CompilerGenerated]
	private CellProp A;

	[CompilerGenerated]
	private CellProp B;

	[CompilerGenerated]
	private Range A;

	public CellProp NewProp
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

	public CellProp OldProp
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

	public Range Range
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

	internal SavedCell(Range A)
	{
		Range = A;
		NewProp = new CellProp();
		OldProp = new CellProp();
		NewProp.A(ref A);
	}
}
