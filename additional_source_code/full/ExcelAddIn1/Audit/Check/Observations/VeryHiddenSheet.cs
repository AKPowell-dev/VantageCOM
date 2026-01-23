using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.Observations;

public sealed class VeryHiddenSheet : Observation
{
	internal VeryHiddenSheet(Severity A, object B)
		: base(Category.HiddenData, A, VH.A(26556))
	{
		base.Subtitle = Conversions.ToString(NewLateBinding.LateGet(B, null, VH.A(19019), new object[0], null, null, null));
		base.Explanation = VH.A(26591);
		base.Sheet = RuntimeHelpers.GetObjectValue(B);
		if (B is Worksheet)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			base.Worksheet = (Worksheet)B;
		}
		else
		{
			base.Chart = (Chart)B;
		}
		base.HasFix = true;
		base.CanFixMultiple = true;
	}

	public override void FixAction()
	{
		base.FixAction();
		if (base.Sheet is Worksheet)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					((Worksheet)base.Sheet).Visible = XlSheetVisibility.xlSheetVisible;
					return;
				}
			}
		}
		((Chart)base.Sheet).Visible = XlSheetVisibility.xlSheetVisible;
	}
}
