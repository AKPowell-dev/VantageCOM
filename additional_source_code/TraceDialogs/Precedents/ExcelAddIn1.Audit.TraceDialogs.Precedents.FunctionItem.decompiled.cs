using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.TraceDialogs.Precedents;

public sealed class FunctionItem : BaseItem
{
	public FunctionItem(BaseItem parent, Range rng, string strLabel, string strValue)
		: base(parent, rng, checked(parent.Level + 1), VH.A(44247))
	{
		base.Label = strLabel;
		base.Value = strValue;
	}
}
