using A;
using MacabacusMacros.Auth;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.RowsColumns;

public sealed class BatchModify
{
	public static void Rows()
	{
		if (!A())
		{
			return;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!(MH.A.Application.Selection is Range))
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				new wpfModifyRows().ShowDialog();
				_ = null;
				Core.LogActivity(VH.A(170102));
				return;
			}
		}
	}

	public static void Columns()
	{
		if (!A() || !(MH.A.Application.Selection is Range))
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			new wpfModifyColumns().ShowDialog();
			_ = null;
			Core.LogActivity(VH.A(170125));
			return;
		}
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}
}
