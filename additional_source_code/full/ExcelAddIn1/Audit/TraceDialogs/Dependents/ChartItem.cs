using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.TraceDialogs.Dependents;

public sealed class ChartItem : BaseItem
{
	[CompilerGenerated]
	private Chart A;

	public Chart Chart
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

	public ChartItem(BaseItem parent, Chart cht, string strLabel, string strValue, int intIndex)
		: base(parent, null, checked(parent.Level + 1), VH.A(41656))
	{
		base.Label = strLabel;
		base.Value = strValue;
		Chart = cht;
		base.Index = intIndex;
	}
}
