using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs.Dependents;

public sealed class SingleCellItem : BaseItem
{
	public SingleCellItem(BaseItem parent, Range rng, string strLabel, int intIndex)
		: base(parent, rng, checked(parent.Level + 1), VH.A(41778))
	{
		base.Label = strLabel;
		base.Value = Base.CleanValueText(Conversions.ToString(rng.Text));
		base.Index = intIndex;
	}
}
