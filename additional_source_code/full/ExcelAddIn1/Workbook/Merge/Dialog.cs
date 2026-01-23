using MacabacusMacros.Auth;

namespace ExcelAddIn1.Workbook.Merge;

public sealed class Dialog
{
	public static void Show()
	{
		if (!Access.AllowExcelOperation((PlanType)5, (Restriction)2, false))
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			new wpfMerge().ShowDialog();
			_ = null;
			return;
		}
	}
}
