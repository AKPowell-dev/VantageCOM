using A;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Audit.TraceDialogs.Dependents;

public sealed class MultiCellItem : BaseItem
{
	public MultiCellItem(BaseItem parent, Range rng, string strLabel, int intIndex)
		: base(parent, rng, checked(parent.Level + 1), VH.A(41778))
	{
		base.Label = strLabel;
		base.Value = VH.A(41885);
		base.Index = intIndex;
		base.Items.Add(new DummyItem(this));
	}
}
