using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit;

public sealed class Watches
{
	private static readonly string m_A = VH.A(49303) + ((_Application)MH.A.Application).get_International((object)XlApplicationInternational.xlThousandsSeparator).ToString() + VH.A(52500);

	[CompilerGenerated]
	private static int m_A;

	private static int WatchIndex
	{
		[CompilerGenerated]
		get
		{
			return Watches.m_A;
		}
		[CompilerGenerated]
		set
		{
			Watches.m_A = value;
		}
	} = 0;

	public static void Add()
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			bool flag = false;
			Range range;
			try
			{
				if (application.Selection is Range)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						range = (Range)application.Selection;
						int num = Conversions.ToInteger(range.Cells.CountLarge);
						if (num > 100)
						{
							if (Operators.ConditionalCompareObjectGreaterEqual(num, range.Worksheet.Columns.CountLarge, TextCompare: false))
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
								Forms.WarningMessage(VH.A(51741));
								flag = true;
							}
							else if (MessageBox.Show(VH.A(51818) + num.ToString(Watches.m_A) + VH.A(51907), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
							{
								flag = true;
							}
						}
						if (flag)
						{
							break;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							try
							{
								enumerator = range.GetEnumerator();
								while (enumerator.MoveNext())
								{
									Range source = (Range)enumerator.Current;
									application.Watches.Add(source);
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_0143;
									}
									continue;
									end_IL_0143:
									break;
								}
							}
							finally
							{
								if (enumerator is IDisposable)
								{
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										(enumerator as IDisposable).Dispose();
										break;
									}
								}
							}
							if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
								ExcelAddIn1.ExcelApp.StatusBar.SetText(VH.A(51940));
							}
							else
							{
								ExcelAddIn1.ExcelApp.StatusBar.SetText(VH.A(51963));
							}
							A(VH.A(51990));
							break;
						}
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(VH.A(52009));
				ProjectData.ClearProjectError();
			}
			application = null;
			range = null;
			return;
		}
	}

	public static void Remove()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
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
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				int num = 0;
				Range range;
				try
				{
					range = (Range)application.Selection;
					try
					{
						enumerator = range.GetEnumerator();
						while (enumerator.MoveNext())
						{
							string right = ((Range)enumerator.Current).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
							int num2 = application.Watches.Count - 1;
							int num3 = 0;
							while (true)
							{
								if (num3 <= num2)
								{
									Range range2 = (Range)application.Watches[num3].Source;
									if (range2 != null)
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
										if (Operators.CompareString(range2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) == 0)
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
											application.Watches[num3].Delete();
											num++;
											break;
										}
									}
									range2 = null;
									num3++;
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
								break;
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_015e;
							}
							continue;
							end_IL_015e:
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
					ExcelAddIn1.ExcelApp.StatusBar.SetText(num.ToString(Watches.m_A) + VH.A(52050));
					A(VH.A(52087));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				range = null;
				application = null;
				return;
			}
		}
	}

	public static void GoToNext()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
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
				Microsoft.Office.Interop.Excel.Watches watches = MH.A.Application.Watches;
				if (watches.Count > 0)
				{
					int num = WatchIndex + 1;
					if (num > watches.Count - 1)
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
						num = 0;
					}
					A(num);
					A(VH.A(52112));
				}
				else
				{
					A();
				}
				watches = null;
				return;
			}
		}
	}

	public static void GoToPrevious()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Watches watches = MH.A.Application.Watches;
		checked
		{
			if (watches.Count > 0)
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
				int num = WatchIndex - 1;
				if (num == -1)
				{
					num = watches.Count - 1;
				}
				A(num);
				A(VH.A(52133));
			}
			else
			{
				A();
			}
			watches = null;
		}
	}

	private static void A()
	{
		Forms.InfoMessage(VH.A(52162));
	}

	private static void A(int A)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Range range = null;
		Microsoft.Office.Interop.Excel.Application application2 = application;
		try
		{
			range = (Range)application2.Watches[A].Source;
			range.Worksheet.Activate();
			try
			{
				Ranges.ScrollIntoView(range);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			range.Select();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		application2 = null;
		WatchIndex = A;
		range = null;
	}

	public static void Clear()
	{
		if (MessageBox.Show(VH.A(52247), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
		{
			return;
		}
		checked
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				int num = 0;
				Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
				try
				{
					for (int i = application.Watches.Count - 1; i >= 0; i += -1)
					{
						if (Operators.CompareString(((Microsoft.Office.Interop.Excel.Workbook)((Range)application.Watches[i].Source).Worksheet.Parent).Name, application.ActiveWorkbook.Name, TextCompare: false) != 0)
						{
							continue;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						application.Watches[i].Delete();
						num++;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						Forms.InfoMessage(VH.A(52374) + num + VH.A(52391));
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				application = null;
				WatchIndex = 0;
				A(VH.A(52448));
				return;
			}
		}
	}

	public static void WatchWindow()
	{
		CommandBar commandBar = MH.A.Application.CommandBars[VH.A(52475)];
		commandBar.Visible = !commandBar.Visible;
		if (commandBar.Visible)
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
			A(VH.A(52475));
		}
		_ = null;
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, A);
	}
}
