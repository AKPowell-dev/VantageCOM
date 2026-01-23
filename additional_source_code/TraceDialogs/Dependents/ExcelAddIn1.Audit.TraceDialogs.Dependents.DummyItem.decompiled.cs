using A;

namespace ExcelAddIn1.Audit.TraceDialogs.Dependents;

public sealed class DummyItem : BaseItem
{
	public DummyItem(BaseItem parent)
		: base(parent, null, checked(parent.Level + 1), "")
	{
		base.Label = VH.A(41757);
	}
}
