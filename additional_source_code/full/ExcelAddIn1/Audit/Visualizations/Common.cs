using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Visualizations;

public sealed class Common
{
	public enum VisualizationType
	{
		FormulaFlow = 1,
		DependencyDensity,
		FunctionalMap,
		MagnitudeHeatmap
	}

	[CompilerGenerated]
	private static int m_A = ColorTranslator.ToOle(Color.FromArgb(0, 200, 255));

	[CompilerGenerated]
	private static List<KeyValuePair<Range, VisualizationType>> m_A;

	internal static int PATTERN_COLOR_BLUE
	{
		[CompilerGenerated]
		get
		{
			return Common.m_A;
		}
	}

	internal static List<KeyValuePair<Range, VisualizationType>> VisualizedRanges
	{
		[CompilerGenerated]
		get
		{
			return Common.m_A;
		}
		[CompilerGenerated]
		set
		{
			Common.m_A = value;
		}
	}

	internal static void A(Range A, VisualizationType B)
	{
		if (VisualizedRanges == null)
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
			VisualizedRanges = new List<KeyValuePair<Range, VisualizationType>>();
		}
		VisualizedRanges.Add(new KeyValuePair<Range, VisualizationType>(A, B));
	}

	public static void ClearVisualizations(Microsoft.Office.Interop.Excel.Application xlApp = null)
	{
		bool flag = false;
		if (xlApp == null)
		{
			xlApp = MH.A.Application;
		}
		if (xlApp.ActiveSheet is Worksheet)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A(xlApp);
			Range usedRange = ((Worksheet)xlApp.ActiveSheet).UsedRange;
			if (Operators.ConditionalCompareObjectGreater(usedRange.Cells.CountLarge, 125000, TextCompare: false))
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
				if (MessageBox.Show(VH.A(50057), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
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
					flag = true;
				}
			}
			if (!flag)
			{
				xlApp.ScreenUpdating = false;
				try
				{
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = usedRange.GetEnumerator();
						while (enumerator.MoveNext())
						{
							A((Range)enumerator.Current);
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (3)
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
				xlApp.ScreenUpdating = true;
			}
			usedRange = null;
		}
		VisualizedRanges = null;
		xlApp = null;
	}

	private static void A(Microsoft.Office.Interop.Excel.Application A)
	{
		FormulaFlow.B(A);
	}

	private static void A(Range A)
	{
		Interior interior = A.Interior;
		if (!Operators.ConditionalCompareObjectEqual(interior.PatternColor, PATTERN_COLOR_BLUE, TextCompare: false))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!Operators.ConditionalCompareObjectEqual(interior.Pattern, XlPattern.xlPatternGray75, TextCompare: false))
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
				if (!Operators.ConditionalCompareObjectEqual(interior.Pattern, XlPattern.xlPatternGray50, TextCompare: false))
				{
					goto IL_0086;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					break;
				}
			}
		}
		B(A);
		goto IL_0086;
		IL_0086:
		interior = null;
	}

	internal static void B(Range A)
	{
		Interior interior = A.Interior;
		if (Operators.ConditionalCompareObjectEqual(interior.Color, ColorTranslator.ToOle(Color.Transparent), TextCompare: false))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			interior.Pattern = XlPattern.xlPatternNone;
		}
		else
		{
			interior.Pattern = XlPattern.xlPatternSolid;
		}
		interior = null;
	}

	internal static Color A(Color A, Color B, double C)
	{
		double a = Common.A((int)A.R, (int)B.R, C);
		double a2 = Common.A((int)A.G, (int)B.G, C);
		checked
		{
			return Color.FromArgb(blue: (int)Math.Round(unchecked(Common.A((int)A.B, (int)B.B, C))), red: (int)Math.Round(a), green: (int)Math.Round(a2));
		}
	}

	private static double A(double A, double B, double C)
	{
		return A - (A - B) * C;
	}

	internal static DialogResult A()
	{
		return MessageBox.Show(VH.A(50883), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);
	}

	internal static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, A);
	}

	public static void RefreshLiveVisualizations(Range rng)
	{
		if (VisualizedRanges == null)
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
			using List<KeyValuePair<Range, VisualizationType>>.Enumerator enumerator = VisualizedRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				KeyValuePair<Range, VisualizationType> current = enumerator.Current;
				try
				{
					if (rng.Application.Intersect(current.Key, rng, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == null)
					{
						continue;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						VisualizationType value = current.Value;
						if (value != VisualizationType.FormulaFlow)
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
							FormulaFlow.A(current.Key.Worksheet, current.Key);
							break;
						}
						break;
					}
					continue;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
	}
}
