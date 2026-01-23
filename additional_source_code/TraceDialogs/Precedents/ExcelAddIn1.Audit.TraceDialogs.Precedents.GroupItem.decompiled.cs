using A;

namespace ExcelAddIn1.Audit.TraceDialogs.Precedents;

public sealed class GroupItem : BaseItem
{
	public GroupItem(BaseItem parent, string strLabel, string strValue)
		: base(parent, null, checked(parent.Level + 1), VH.A(45205))
	{
		base.Label = strLabel;
		base.Value = strValue;
	}
}
