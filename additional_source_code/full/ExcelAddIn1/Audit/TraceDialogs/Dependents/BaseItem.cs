using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.TraceDialogs.Dependents;

public class BaseItem : TraceItem
{
	private string A;

	private string B;

	[CompilerGenerated]
	private int A;

	[CompilerGenerated]
	private int B;

	public string Value
	{
		get
		{
			return this.A;
		}
		set
		{
			this.A = value;
			NotifyPropertyChanged(VH.A(41636));
		}
	}

	public string Info
	{
		get
		{
			return this.B;
		}
		set
		{
			this.B = value;
			NotifyPropertyChanged(VH.A(41647));
		}
	}

	public int SelectionIndex
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

	public int SelectionLength
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

	public BaseItem(BaseItem p, Range rng, int intLevel, string strIcon)
		: base(p, rng, intLevel, strIcon)
	{
		base.Index = 0;
		base.IsExpanded = false;
	}
}
