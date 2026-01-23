using System;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Miscellaneous
{
	public static void LeaderDots()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Application application = MH.A.Application;
		application.ScreenUpdating = false;
		Range range;
		try
		{
			range = JH.A((Range)null);
			bool flag = JH.A(range);
			string text;
			if (!Operators.ConditionalCompareObjectEqual(application.ActiveCell.NumberFormat, VH.A(150564), TextCompare: false))
			{
				while (true)
				{
					switch (3)
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
				text = VH.A(150564);
			}
			else
			{
				text = VH.A(48146);
			}
			string numberFormat = text;
			foreach (Range item in range)
			{
				if (!Base.IsString(item))
				{
					continue;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				item.NumberFormat = numberFormat;
			}
			if (flag)
			{
				JH.A(range, VH.A(150571));
			}
			Base.LogActivity(VH.A(150571));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.HandleFormattingException(ex2);
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application = null;
		range = null;
	}

	public static void WrapText()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			bool flag = false;
			Application application = MH.A.Application;
			if (application.Selection is Range)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				Range range = JH.A((Range)null);
				if (!Base.IsWorksheetProtected(range.Worksheet))
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					application.ScreenUpdating = false;
					try
					{
						if (KH.A.UndoAlignment)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								break;
							}
							flag = JH.A(range);
						}
						application.CommandBars.ExecuteMso(VH.A(150594));
						if (flag)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								break;
							}
							JH.A(range, VH.A(148068));
						}
						Base.LogActivity(VH.A(150611));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Base.HandleFormattingException(ex2);
						ProjectData.ClearProjectError();
					}
					application.ScreenUpdating = true;
				}
				range = null;
			}
			application = null;
			return;
		}
	}
}
