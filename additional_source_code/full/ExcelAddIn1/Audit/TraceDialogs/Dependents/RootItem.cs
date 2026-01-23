using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs.Dependents;

public sealed class RootItem : BaseItem
{
	public RootItem(Range rng)
		: base(null, rng, 0, VH.A(42564))
	{
		base.Items.Add(new DummyItem(this));
		base.Label = rng.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		if (Operators.ConditionalCompareObjectEqual(rng.Cells.CountLarge, 1, TextCompare: false))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					base.Value = Base.CleanValueText(Conversions.ToString(rng.Text));
					return;
				}
			}
		}
		base.Value = VH.A(41885);
	}
}
