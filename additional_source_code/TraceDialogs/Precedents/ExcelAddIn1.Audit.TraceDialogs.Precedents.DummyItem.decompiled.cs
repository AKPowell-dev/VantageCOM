namespace ExcelAddIn1.Audit.TraceDialogs.Precedents;

public sealed class DummyItem : BaseItem
{
	public DummyItem(BaseItem parent)
		: base(parent, null, checked(parent.Level + 1), "")
	{
	}
}
