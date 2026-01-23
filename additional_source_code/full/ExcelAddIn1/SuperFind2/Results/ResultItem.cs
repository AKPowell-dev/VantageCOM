using System.Runtime.CompilerServices;
using System.Windows.Media;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.SuperFind2.Results;

public abstract class ResultItem : BaseItem
{
	[CompilerGenerated]
	private SheetItem A;

	[CompilerGenerated]
	private Range A;

	[CompilerGenerated]
	private Worksheet A;

	[CompilerGenerated]
	private int A;

	internal SheetItem Parent
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

	internal Range Range
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

	internal Worksheet Worksheet
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

	internal int UiIndex
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

	public ResultItem(SheetItem si, Worksheet ws, string strLabel, Geometry geo, int index)
		: base(strLabel, geo)
	{
		Parent = si;
		Worksheet = ws;
		base.Workbook = (Microsoft.Office.Interop.Excel.Workbook)ws.Parent;
		int indentLevel;
		if (si.Parent != null)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			indentLevel = 2;
		}
		else
		{
			indentLevel = 1;
		}
		((BaseItem)this).IndentLevel = indentLevel;
		UiIndex = index;
	}
}
