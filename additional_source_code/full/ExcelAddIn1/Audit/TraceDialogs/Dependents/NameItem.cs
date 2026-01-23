using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs.Dependents;

public sealed class NameItem : BaseItem
{
	public NameItem(BaseItem parent, Range rng, int intIndex)
		: base(parent, rng, checked(parent.Level + 1), VH.A(41900))
	{
		base.Label = ((Name)rng.Name).Name;
		base.Value = Base.CleanValueText(Conversions.ToString(rng.Text));
		base.Index = intIndex;
	}
}
