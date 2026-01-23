using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.RowsColumns;

public sealed class AutoFit
{
	public enum HiddenBehavior
	{
		AlwaysAutoFit,
		NeverAutoFit,
		ShowPrompt
	}

	[CompilerGenerated]
	private static HiddenBehavior m_A;

	public static HiddenBehavior Behavior
	{
		[CompilerGenerated]
		get
		{
			return AutoFit.m_A;
		}
		[CompilerGenerated]
		set
		{
			AutoFit.m_A = value;
		}
	} = HiddenBehavior.AlwaysAutoFit;

	public static void Height()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Range range = null;
			if (Behavior == HiddenBehavior.AlwaysAutoFit)
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
				A();
			}
			else
			{
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				try
				{
					if (application.Selection is Range)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							Range range2 = (Range)((Range)application.Selection).Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
							long num = Conversions.ToLong(range2.Cells.CountLarge);
							long num2;
							if (num > 1)
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
								try
								{
									range = range2.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
									num2 = Conversions.ToLong(range.Cells.CountLarge);
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									num2 = 0L;
									ProjectData.ClearProjectError();
								}
							}
							else if (Conversions.ToBoolean(((Range)range2.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).Hidden))
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
								num2 = 0L;
							}
							else
							{
								num2 = 1L;
							}
							if (num == num2)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									A();
									break;
								}
								break;
							}
							if (Behavior == HiddenBehavior.ShowPrompt)
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
								if (A(VH.A(169712)))
								{
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										A();
										break;
									}
									break;
								}
							}
							if (Operators.ConditionalCompareObjectLess(num, range2.Worksheet.Rows.CountLarge, TextCompare: false))
							{
								application.ScreenUpdating = false;
								try
								{
									if (range != null)
									{
										Range range3 = JH.A(range, application);
										if (range3 != null)
										{
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												try
												{
													enumerator = range3.Rows.GetEnumerator();
													while (enumerator.MoveNext())
													{
														((Range)enumerator.Current).AutoFit();
													}
													while (true)
													{
														switch (5)
														{
														case 0:
															break;
														default:
															goto end_IL_0214;
														}
														continue;
														end_IL_0214:
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
												range3 = null;
												break;
											}
										}
									}
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									Forms.ErrorMessage(ex4.Message);
									ProjectData.ClearProjectError();
								}
								application.ScreenUpdating = true;
							}
							else
							{
								Forms.WarningMessage(VH.A(169753));
							}
							break;
						}
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					A();
					ProjectData.ClearProjectError();
				}
				finally
				{
					Range range2 = null;
					range = null;
				}
				application = null;
			}
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(169838));
			return;
		}
	}

	private static void A()
	{
		A(VH.A(169867));
	}

	public static void Width()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Range range = null;
			if (Behavior == HiddenBehavior.AlwaysAutoFit)
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
				B();
			}
			else
			{
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				try
				{
					if (application.Selection is Range)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							Range range2 = (Range)((Range)application.Selection).Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
							long num = Conversions.ToLong(range2.Cells.CountLarge);
							long num2;
							if (num > 1)
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
								try
								{
									range = range2.SpecialCells(XlCellType.xlCellTypeVisible, RuntimeHelpers.GetObjectValue(Missing.Value));
									num2 = Conversions.ToLong(range.Cells.CountLarge);
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									num2 = 0L;
									ProjectData.ClearProjectError();
								}
							}
							else if (Conversions.ToBoolean(((Range)range2.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]).Hidden))
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
								num2 = 0L;
							}
							else
							{
								num2 = 1L;
							}
							if (num == num2)
							{
								B();
							}
							else if (Behavior == HiddenBehavior.ShowPrompt && A(VH.A(169900)))
							{
								B();
							}
							else if (Operators.ConditionalCompareObjectLess(num, range2.Worksheet.Columns.CountLarge, TextCompare: false))
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									application.ScreenUpdating = false;
									try
									{
										if (range != null)
										{
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												Range range3 = JH.A(range, application);
												if (range3 == null)
												{
													break;
												}
												while (true)
												{
													switch (3)
													{
													case 0:
														continue;
													}
													try
													{
														enumerator = range3.Columns.GetEnumerator();
														while (enumerator.MoveNext())
														{
															((Range)enumerator.Current).AutoFit();
														}
														while (true)
														{
															switch (2)
															{
															case 0:
																break;
															default:
																goto end_IL_0206;
															}
															continue;
															end_IL_0206:
															break;
														}
													}
													finally
													{
														if (enumerator is IDisposable)
														{
															while (true)
															{
																switch (5)
																{
																case 0:
																	continue;
																}
																(enumerator as IDisposable).Dispose();
																break;
															}
														}
													}
													range3 = null;
													break;
												}
												break;
											}
										}
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										Forms.ErrorMessage(ex4.Message);
										ProjectData.ClearProjectError();
									}
									application.ScreenUpdating = true;
									break;
								}
							}
							else
							{
								Forms.WarningMessage(VH.A(169947));
							}
							break;
						}
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					B();
					ProjectData.ClearProjectError();
				}
				finally
				{
					Range range2 = null;
					range = null;
				}
				application = null;
			}
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(170038));
			return;
		}
	}

	private static void B()
	{
		A(VH.A(170065));
	}

	private static void A(string A)
	{
		try
		{
			MH.A.Application.CommandBars.ExecuteMso(A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static bool A(string A)
	{
		return MessageBox.Show(A, VH.A(40448), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1) == DialogResult.Yes;
	}
}
