using System;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class FontStyle
{
	[CompilerGenerated]
	private static int m_A;

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return FontStyle.m_A;
		}
		[CompilerGenerated]
		set
		{
			FontStyle.m_A = value;
		}
	}

	public static void Cycle()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
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
				try
				{
					if (CycleIndex > KH.A.FontStyleCycle.Count - 1)
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
						CycleIndex = 0;
					}
					A(KH.A.FontStyleCycle[CycleIndex]);
					CycleIndex++;
					if (CycleIndex != 1)
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
						Base.LogActivity(VH.A(149929));
						return;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}

	public static void DoFontStyle(IRibbonControl control)
	{
		A(control.Tag);
	}

	private static void A(string A)
	{
		Application application = MH.A.Application;
		bool flag = false;
		try
		{
			if (application.Selection is Range)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Range range = JH.A((Range)null);
					if (!Base.IsWorksheetProtected(range.Worksheet))
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
						application.ScreenUpdating = false;
						try
						{
							if (KH.A.UndoFont)
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
								flag = JH.A(range);
							}
							range.Font.Name = A;
							if (flag)
							{
								JH.A(range, VH.A(60635));
							}
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
					break;
				}
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Base.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		application = null;
	}
}
