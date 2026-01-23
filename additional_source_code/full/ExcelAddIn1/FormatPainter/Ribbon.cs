using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Charts;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.FormatPainter;

public sealed class Ribbon
{
	public static void Copy()
	{
		Chart chart = null;
		Window activeWindow;
		try
		{
			activeWindow = MH.A.Application.ActiveWindow;
			chart = MH.A.Application.ActiveChart;
			if (chart != null)
			{
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
					Pane.CopiedProperties = new Properties(chart);
					CustomTaskPane paneByHwnd = Pane.GetPaneByHwnd(activeWindow.Hwnd);
					if (paneByHwnd == null)
					{
						break;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						if (!paneByHwnd.Visible)
						{
							break;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							((ctpFormatPainter)paneByHwnd.Control).A.PopulateProperties();
							break;
						}
						break;
					}
					break;
				}
			}
			else
			{
				MessageBox.Show(VH.A(173074), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			MessageBox.Show(VH.A(173074), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			ProjectData.ClearProjectError();
		}
		activeWindow = null;
		chart = null;
	}

	public static void OpenPane()
	{
	}

	public static void Layout()
	{
		A(B, VH.A(173171));
	}

	private static void A(Chart A, Properties B)
	{
		Ribbon.B(A, B);
		J(A, B);
		G(A, B);
		J(A, B);
	}

	public static void ChartSize()
	{
		A(B, VH.A(173204));
	}

	private static void B(Chart A, Properties B)
	{
		C(A, B);
		D(A, B);
	}

	public static void ChartHeight()
	{
		A(C, VH.A(173237));
	}

	private static void C(Chart A, Properties B)
	{
		((ChartObject)A.Parent).Height = B.ChartObject.Height;
	}

	public static void ChartWidth()
	{
		A(D, VH.A(173274));
	}

	private static void D(Chart A, Properties B)
	{
		((ChartObject)A.Parent).Width = B.ChartObject.Width;
	}

	public static void ChartTop()
	{
		A(E, VH.A(173309));
	}

	private static void E(Chart A, Properties B)
	{
		((ChartObject)A.Parent).Top = B.ChartObject.Top;
	}

	public static void ChartLeft()
	{
		A(F, VH.A(173340));
	}

	private static void F(Chart A, Properties B)
	{
		((ChartObject)A.Parent).Left = B.ChartObject.Left;
	}

	public static void PlotSize()
	{
		A(B, VH.A(173373));
	}

	private static void G(Chart A, Properties B)
	{
		H(A, B);
		I(A, B);
	}

	public static void PlotHeight()
	{
		A(H, VH.A(173414));
	}

	private static void H(Chart A, Properties B)
	{
		A.PlotArea.InsideHeight = B.PlotArea.InsideHeight;
	}

	public static void PlotWidth()
	{
		A(D, VH.A(173459));
	}

	private static void I(Chart A, Properties B)
	{
		A.PlotArea.InsideWidth = B.PlotArea.InsideWidth;
	}

	public static void PlotPosition()
	{
		A(B, VH.A(173502));
	}

	private static void J(Chart A, Properties B)
	{
		H(A, B);
		I(A, B);
	}

	public static void PlotTop()
	{
		A(H, VH.A(173551));
	}

	private static void K(Chart A, Properties B)
	{
		A.PlotArea.InsideTop = B.PlotArea.InsideTop;
	}

	public static void PlotLeft()
	{
		A(D, VH.A(173590));
	}

	private static void L(Chart A, Properties B)
	{
		A.PlotArea.InsideLeft = B.PlotArea.InsideLeft;
	}

	private static void A(Action<Chart, Properties> A, string B)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		int num = 0;
		Chart chart;
		if (Pane.CopiedProperties != null)
		{
			try
			{
				chart = Helpers.SelectedChart();
				if (chart != null)
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
						A(chart, Pane.CopiedProperties);
						break;
					}
				}
				else if (Operators.CompareString(Versioned.TypeName(RuntimeHelpers.GetObjectValue(application.Selection)), VH.A(56245), TextCompare: false) == 0)
				{
					IEnumerator enumerator = default(IEnumerator);
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						try
						{
							try
							{
								enumerator = ((IEnumerable)NewLateBinding.LateGet(application.Selection, null, VH.A(56274), new object[0], null, null, null)).GetEnumerator();
								while (enumerator.MoveNext())
								{
									Shape shape = (Shape)enumerator.Current;
									if (shape.HasChart != MsoTriState.msoTrue)
									{
										continue;
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										break;
									}
									A(shape.Chart, Pane.CopiedProperties);
									num = checked(num + 1);
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_00f8;
									}
									continue;
									end_IL_00f8:
									break;
								}
							}
							finally
							{
								if (enumerator is IDisposable)
								{
									while (true)
									{
										switch (7)
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
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						if (num != 0)
						{
							break;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							Ribbon.A();
							break;
						}
						break;
					}
				}
				else
				{
					Ribbon.A();
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Ribbon.A();
				ProjectData.ClearProjectError();
			}
		}
		else
		{
			MessageBox.Show(VH.A(173631), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}
		application = null;
		chart = null;
	}

	private static void A()
	{
		MessageBox.Show(VH.A(172145), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
	}
}
