using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs.Precedents;

public sealed class RootItem : BaseItem
{
	public RootItem(Range rng)
		: base(null, rng, 0, VH.A(42564))
	{
		base.Items.Add(new DummyItem(this));
		base.Label = rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.Value = Base.CleanValueText(Conversions.ToString(rng.Text));
	}
}
