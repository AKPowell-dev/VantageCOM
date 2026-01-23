using System;
using System.Collections;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Decimals
{
	public static void Increase()
	{
		A(VH.A(149418));
	}

	public static void Decrease()
	{
		A(VH.A(149451));
	}

	private static void A(string A)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Range range = null;
			Application application = MH.A.Application;
			Range range2;
			try
			{
				if (application.Selection is Range)
				{
					range2 = (Range)application.Selection;
					if (!Base.IsWorksheetProtected(range2.Worksheet))
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							bool flag = JH.A(range2);
							application.ScreenUpdating = false;
							application.EnableEvents = false;
							bool useSystemSeparators = application.UseSystemSeparators;
							string thousandsSeparator = application.ThousandsSeparator;
							string decimalSeparator = application.DecimalSeparator;
							application.ThousandsSeparator = VH.A(2378);
							application.DecimalSeparator = VH.A(64021);
							application.UseSystemSeparators = false;
							try
							{
								if (string.IsNullOrEmpty(range2.NumberFormat.ToString()))
								{
									if (Operators.ConditionalCompareObjectLessEqual(range2.Cells.CountLarge, 1000, TextCompare: false))
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											range = application.ActiveCell;
											enumerator = range2.GetEnumerator();
											try
											{
												while (enumerator.MoveNext())
												{
													((Range)enumerator.Current).Select();
													Decimals.A(application.CommandBars, A);
												}
												while (true)
												{
													switch (6)
													{
													case 0:
														break;
													default:
														goto end_IL_0152;
													}
													continue;
													end_IL_0152:
													break;
												}
											}
											finally
											{
												IDisposable disposable = enumerator as IDisposable;
												if (disposable != null)
												{
													disposable.Dispose();
												}
											}
											break;
										}
									}
									else
									{
										Decimals.A(application.CommandBars, A);
									}
								}
								else
								{
									Decimals.A(application.CommandBars, A);
								}
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								Base.HandleFormattingException(ex2);
								ProjectData.ClearProjectError();
							}
							if (range != null)
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
								range2.Select();
								range.Activate();
								range = null;
							}
							application.ThousandsSeparator = thousandsSeparator;
							application.DecimalSeparator = decimalSeparator;
							application.UseSystemSeparators = useSystemSeparators;
							if (flag)
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
								JH.A(range2, VH.A(148068));
							}
							Base.LogActivity(VH.A(94482));
							break;
						}
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
			finally
			{
				application.ScreenUpdating = true;
				application.EnableEvents = true;
			}
			application = null;
			range2 = null;
			return;
		}
	}

	private static void A(CommandBars A, string B)
	{
		try
		{
			A.ExecuteMso(B);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			try
			{
				if (!A.GetEnabledMso(B))
				{
					goto IL_0043;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (!A.GetVisibleMso(B))
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
					goto IL_0043;
				}
				goto end_IL_0010;
				IL_0043:
				Forms.ErrorMessage(VH.A(149484));
				end_IL_0010:;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			throw;
		}
	}
}
