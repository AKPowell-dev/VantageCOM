using System;
using System.Collections;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Indent
{
	public static void Left()
	{
		A(XlHAlign.xlHAlignLeft, KH.A.IndentMaxLeft, VH.A(149962));
	}

	public static void Right()
	{
		A(XlHAlign.xlHAlignRight, KH.A.IndentMaxRight, VH.A(149985));
	}

	private static void A(XlHAlign A, int B, string C)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Range range = JH.A((Range)null);
			bool flag = false;
			if (!Base.IsWorksheetProtected(range.Worksheet))
			{
				if (KH.A.UndoAlignment)
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
				Application application = MH.A.Application;
				application.ScreenUpdating = false;
				try
				{
					int num = Conversions.ToInteger(application.ActiveCell.IndentLevel);
					bool flag2 = num > checked(B - 1);
					try
					{
						range.HorizontalAlignment = A;
						if (string.IsNullOrEmpty(range.IndentLevel.ToString()))
						{
							try
							{
								enumerator = range.GetEnumerator();
								while (enumerator.MoveNext())
								{
									Range range2 = (Range)enumerator.Current;
									Range range3 = range2;
									object obj;
									if (!flag2)
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
										obj = Operators.AddObject(range2.IndentLevel, 1);
									}
									else
									{
										obj = NewLateBinding.LateGet(null, typeof(Math), VH.A(69118), new object[2]
										{
											Operators.SubtractObject(range2.IndentLevel, num),
											0
										}, null, null, null);
									}
									range3.IndentLevel = RuntimeHelpers.GetObjectValue(obj);
									range2 = null;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_0161;
									}
									continue;
									end_IL_0161:
									break;
								}
							}
							finally
							{
								if (enumerator is IDisposable)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										(enumerator as IDisposable).Dispose();
										break;
									}
								}
							}
						}
						else
						{
							Range range4 = range;
							object obj2;
							if (!flag2)
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
								obj2 = Operators.AddObject(range.IndentLevel, 1);
							}
							else
							{
								obj2 = NewLateBinding.LateGet(null, typeof(Math), VH.A(69118), new object[2]
								{
									Operators.SubtractObject(range.IndentLevel, num),
									0
								}, null, null, null);
							}
							range4.IndentLevel = RuntimeHelpers.GetObjectValue(obj2);
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						Base.HandleFormattingException(ex2);
						ProjectData.ClearProjectError();
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				application.ScreenUpdating = true;
				application = null;
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
				Base.LogActivity(C);
			}
			range = null;
			return;
		}
	}
}
