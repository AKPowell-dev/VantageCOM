using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.UndoRedo;

public sealed class WorkbookItem
{
	[CompilerGenerated]
	private Stack A;

	[CompilerGenerated]
	private Stack B;

	[CompilerGenerated]
	private Dictionary<Worksheet, Dictionary<Range, CellProp>> A;

	public Stack UndoStack
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

	public Stack RedoStack
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

	public Dictionary<Worksheet, Dictionary<Range, CellProp>> BaseSheets
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

	internal WorkbookItem()
	{
		UndoStack = new Stack();
		RedoStack = new Stack();
		BaseSheets = new Dictionary<Worksheet, Dictionary<Range, CellProp>>(new WorksheetComparer());
	}
}
