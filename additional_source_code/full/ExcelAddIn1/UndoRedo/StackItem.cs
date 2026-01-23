using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.UndoRedo;

public sealed class StackItem
{
	[CompilerGenerated]
	private SavedCell[] A;

	[CompilerGenerated]
	private Range A;

	[CompilerGenerated]
	private string A;

	public SavedCell[] Cells
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

	public string Name
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

	internal StackItem(SavedCell[] A, Range B, string C)
	{
		Cells = A;
		Range = B;
		Name = C;
	}
}
