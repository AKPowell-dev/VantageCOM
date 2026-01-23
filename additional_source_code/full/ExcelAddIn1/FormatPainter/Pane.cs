using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Charts;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FormatPainter;

public sealed class Pane
{
	private static readonly string m_A = VH.A(172876);

	private static readonly string B = VH.A(172905);

	private static Properties m_A = null;

	private static Properties B = null;

	private static Dictionary<int, CustomTaskPane> m_A = null;

	public static Properties CopiedProperties
	{
		get
		{
			return Pane.m_A;
		}
		set
		{
			Pane.m_A = value;
		}
	}

	public static Properties CurrentProperties
	{
		get
		{
			return B;
		}
		set
		{
			B = value;
		}
	}

	private static Dictionary<int, CustomTaskPane> A
	{
		get
		{
			return Pane.m_A;
		}
		set
		{
			Pane.m_A = value;
		}
	}

	public static void Toggle(bool blnPressed)
	{
		Window activeWindow = MH.A.Application.ActiveWindow;
		CustomTaskPane value = null;
		int hwnd = activeWindow.Hwnd;
		bool flag = false;
		if (Pane.A != null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (Pane.A.TryGetValue(hwnd, out value))
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
				value.Visible = blnPressed;
				flag = true;
			}
		}
		if (blnPressed)
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
			if (!flag)
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
				ctpFormatPainter control = new ctpFormatPainter();
				value = MH.A.CustomTaskPanes.Add(control, Pane.m_A, activeWindow);
				value.Width = 400;
				value.VisibleChanged += A;
				value.Visible = true;
				if (Pane.A == null)
				{
					Pane.A = new Dictionary<int, CustomTaskPane>();
				}
				Pane.A.Add(hwnd, value);
			}
		}
		activeWindow = null;
		value = null;
	}

	private static void A(object A, EventArgs B)
	{
		CustomTaskPane customTaskPane = (CustomTaskPane)A;
		ctpFormatPainter ctpFormatPainter = (ctpFormatPainter)customTaskPane.Control;
		_ = MH.A.Application;
		Chart chart = null;
		ctpFormatPainter.A.Visible = customTaskPane.Visible;
		if (customTaskPane.Visible)
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
			try
			{
				chart = Helpers.SelectedChart();
				FormatTree a = ctpFormatPainter.A;
				if (CopiedProperties == null)
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
					if (chart != null)
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
						CopiedProperties = new Properties(chart);
						a.PopulateProperties();
						a.btnCopy.IsEnabled = true;
					}
				}
				else
				{
					a.PopulateProperties();
					if (chart != null)
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
						a.btnApply.IsEnabled = true;
						a.btnCopy.IsEnabled = true;
					}
					else
					{
						a.btnApply.IsEnabled = false;
						a.btnCopy.IsEnabled = false;
					}
				}
				a = null;
				chart = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		else if (!KH.A)
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
			Pane.A();
		}
		customTaskPane = null;
		ctpFormatPainter = null;
	}

	public static bool IsSingleChartSelected()
	{
		return Helpers.SelectedChart() != null;
	}

	public static bool IsChartSelected()
	{
		if (Helpers.SelectedChart() != null)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		if (Operators.CompareString(Versioned.TypeName(RuntimeHelpers.GetObjectValue(MH.A.Application.Selection)), VH.A(56245), TextCompare: false) == 0)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					try
					{
						try
						{
							enumerator = ((IEnumerable)NewLateBinding.LateGet(MH.A.Application.Selection, null, VH.A(56274), new object[0], null, null, null)).GetEnumerator();
							while (enumerator.MoveNext())
							{
								if (((Shape)enumerator.Current).HasChart == MsoTriState.msoTrue)
								{
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											return true;
										}
									}
								}
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_00c7;
								}
								continue;
								end_IL_00c7:
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
										break;
									default:
										(enumerator as IDisposable).Dispose();
										goto end_IL_00db;
									}
									continue;
									end_IL_00db:
									break;
								}
							}
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					return false;
				}
			}
		}
		return false;
	}

	public static bool IsVisible()
	{
		bool result;
		try
		{
			result = GetPaneByHwnd(MH.A.Application.ActiveWindow.Hwnd).Visible;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static CustomTaskPane GetPaneByHwnd(int hwnd)
	{
		CustomTaskPane value = null;
		try
		{
			if (Pane.A != null)
			{
				Pane.A.TryGetValue(hwnd, out value);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return value;
	}

	private static void A()
	{
		KH.A.InvalidateControl(Pane.B);
	}
}
