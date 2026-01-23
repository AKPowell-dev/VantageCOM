using System;
using System.Collections;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class FontSize
{
	[CompilerGenerated]
	private static int m_A;

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return FontSize.m_A;
		}
		[CompilerGenerated]
		set
		{
			FontSize.m_A = value;
		}
	}

	public static void Increase()
	{
		A(VH.A(149813));
	}

	public static void Decrease()
	{
		A(VH.A(149846));
	}

	private static void A(string A)
	{
		Application application = MH.A.Application;
		bool flag = false;
		Range range2;
		Range range;
		Range activeCell;
		try
		{
			range = (Range)application.Selection;
			if (!Base.IsWorksheetProtected(range.Worksheet))
			{
				activeCell = application.ActiveCell;
				range2 = JH.A(range);
				if (KH.A.UndoFont)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					flag = JH.A(range);
				}
				Application application2 = application;
				application2.ScreenUpdating = false;
				if (string.IsNullOrEmpty(range2.Font.Size.ToString()))
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
					application2.EnableEvents = false;
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = range2.Cells.GetEnumerator();
						while (enumerator.MoveNext())
						{
							((Range)enumerator.Current).Select();
							application2.CommandBars.ExecuteMso(A);
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_00eb;
							}
							continue;
							end_IL_00eb:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					range.Select();
					activeCell.Activate();
					application2.EnableEvents = true;
				}
				else
				{
					application2.CommandBars.ExecuteMso(A);
				}
				application2.ScreenUpdating = true;
				application2 = null;
				if (flag)
				{
					JH.A(range, VH.A(149879));
				}
				Base.LogActivity(VH.A(149879));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		range2 = null;
		range = null;
		activeCell = null;
		application = null;
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
				switch (7)
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
					if (CycleIndex > KH.A.FontSizeCycle.Count - 1)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							break;
						}
						CycleIndex = 0;
					}
					A(KH.A.FontSizeCycle[CycleIndex]);
					CycleIndex++;
					if (CycleIndex == 1)
					{
						Base.LogActivity(VH.A(149898));
					}
					return;
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

	public static void DoFontSize(IRibbonControl control)
	{
		A(Conversions.ToSingle(control.Tag));
	}

	private static void A(float A)
	{
		Application application = MH.A.Application;
		bool flag = false;
		try
		{
			if (application.Selection is Range)
			{
				Range range = JH.A((Range)null);
				if (!Base.IsWorksheetProtected(range.Worksheet))
				{
					while (true)
					{
						switch (2)
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
					application.ScreenUpdating = false;
					try
					{
						range = JH.A((Range)null);
						if (KH.A.UndoFont)
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
							flag = JH.A(range);
						}
						range.Font.Size = A;
						if (flag)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								JH.A(range, VH.A(60635));
								break;
							}
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
