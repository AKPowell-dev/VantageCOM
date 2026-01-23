using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.TextOps.Redaction.Values;

public sealed class SelectionValue
{
	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private int B;

	[CompilerGenerated]
	private WdViewType A;

	[CompilerGenerated]
	private WdSelectionType A;

	[CompilerGenerated]
	private ShapeRange A;

	[CompilerGenerated]
	private ShapeRange B;

	[CompilerGenerated]
	private Application A;

	[CompilerGenerated]
	private Document A;

	[CompilerGenerated]
	private bool A;

	[CompilerGenerated]
	private bool B;

	[CompilerGenerated]
	private bool C;

	[CompilerGenerated]
	private bool D;

	[CompilerGenerated]
	private bool E;

	public int SelStart
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public int SelEnd
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
	}

	public WdViewType ViewType
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public WdSelectionType SelType
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public ShapeRange ShapeRange
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public ShapeRange ChildShapeRange
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
	}

	public Application WdApp
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public Document ActiveDocument
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
	}

	public bool IsInHeaderFooter
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
	}

	public bool IsTextInsideShape
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
	}

	public bool IsTextInsideTable
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
	}

	public bool IsNonFloatingShape
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
	}

	public bool IsFloatingShape
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
	}

	public SelectionValue(Application wdApp)
	{
		this.A = wdApp;
		Selection selection = wdApp.Selection;
		this.A = selection.ShapeRange;
		this.B = selection.ChildShapeRange;
		this.A = selection.Range.Start;
		this.B = selection.Range.End;
		this.A = wdApp.ActiveWindow.View.Type;
		this.A = selection.Type;
		B = clsUtilities.IsSelectionTextInsideShape(selection);
		C = clsUtilities.IsSelectionTextInsideTable(selection);
		E = clsUtilities.IsSelectionTypeFloatingShape(selection);
		D = clsUtilities.IsSelectionTypeNonFloatingShape(selection);
		A = clsUtilities.IsRangeInHeaderFooter(selection.Range);
	}
}
