using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class FontColor
{
	[CompilerGenerated]
	private static int m_A;

	internal static int AutoColorIndex
	{
		[CompilerGenerated]
		get
		{
			return FontColor.m_A;
		}
		[CompilerGenerated]
		set
		{
			FontColor.m_A = value;
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
				Application application = MH.A.Application;
				bool flag = false;
				if (application.Selection is Range)
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
					ColorCycle fontColorCycle = KH.A.FontColorCycle;
					int count = fontColorCycle.Colors.Count;
					if (count > 0)
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
						Range range = (Range)application.Selection;
						if (!Base.IsWorksheetProtected(range.Worksheet))
						{
							application.ScreenUpdating = false;
							try
							{
								if (KH.A.UndoFont)
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
									flag = JH.A(range);
								}
								range.Font.Color = fontColorCycle.Colors[fontColorCycle.Index].OLE;
								if (fontColorCycle.Index == 0)
								{
									Base.LogActivity(fontColorCycle.Activity);
								}
								if (fontColorCycle.Index < count - 1)
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
									fontColorCycle.Index++;
								}
								else
								{
									fontColorCycle.Index = 0;
								}
								if (flag)
								{
									while (true)
									{
										switch (3)
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
					fontColorCycle = null;
				}
				application = null;
				return;
			}
		}
	}

	public static void CycleAutoColors()
	{
		List<string> list = null;
		bool flag = false;
		if (!(MH.A.Application.Selection is Range))
		{
			return;
		}
		checked
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
				Range range = (Range)MH.A.Application.Selection;
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
					list = new List<string>();
					using (List<string>.Enumerator enumerator = KH.A.AutoColors.GetEnumerator())
					{
						while (enumerator.MoveNext())
						{
							string current = enumerator.Current;
							if (list.Contains(current) || current.Length <= 0)
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
							list.Add(current);
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_00c5;
							}
							continue;
							end_IL_00c5:
							break;
						}
					}
					try
					{
						if (KH.A.UndoFont)
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
							flag = JH.A(range);
						}
						range.Font.Color = clsColors.RGB2Ole(list[AutoColorIndex]);
						if (AutoColorIndex == 0)
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
							Base.LogActivity(VH.A(149675));
						}
						if (AutoColorIndex < list.Count - 1)
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
							AutoColorIndex++;
						}
						else
						{
							AutoColorIndex = 0;
						}
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
					list = null;
				}
				range = null;
				return;
			}
		}
	}

	public static void BlueBlackToggle()
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
			if (!(MH.A.Application.Selection is Range))
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				Range range;
				try
				{
					range = JH.A((Range)null);
					if (!Base.IsWorksheetProtected(range.Worksheet))
					{
						if (KH.A.UndoFont)
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
							flag = JH.A(range);
						}
						string text = KH.A.AutoColors[0];
						int num;
						if (text.Length > 0)
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
							num = clsColors.RGB2Ole(text);
						}
						else
						{
							num = ColorTranslator.ToOle(Color.Blue);
						}
						int defaultFontColor = KH.A.DefaultFontColor;
						try
						{
							Microsoft.Office.Interop.Excel.Font font = range.Font;
							int num2;
							if (!Operators.ConditionalCompareObjectEqual(range.Application.ActiveCell.Font.Color, num, TextCompare: false))
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
								num2 = num;
							}
							else
							{
								num2 = defaultFontColor;
							}
							font.Color = num2;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							range.Font.Color = defaultFontColor;
							ProjectData.ClearProjectError();
						}
						if (flag)
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
							JH.A(range, VH.A(60635));
						}
						Base.LogActivity(VH.A(149706));
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					Base.HandleFormattingException(ex4);
					ProjectData.ClearProjectError();
				}
				range = null;
				return;
			}
		}
	}

	public static void Automatic()
	{
		bool flag = false;
		if (!(MH.A.Application.Selection is Range))
		{
			return;
		}
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
			Range range;
			try
			{
				range = JH.A((Range)null);
				if (KH.A.UndoFont)
				{
					flag = JH.A(range);
				}
				range.Font.ColorIndex = XlColorIndex.xlColorIndexAutomatic;
				if (flag)
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
					JH.A(range, VH.A(60635));
				}
				K.Settings.LastFontColor = ColorTranslator.FromOle(Conversions.ToInteger(range.Font.Color));
				KH.A.InvalidateControl(clsColors.LAST_FONT_COLOR_BUTTON);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Base.HandleFormattingException(ex2);
				ProjectData.ClearProjectError();
			}
			range = null;
			return;
		}
	}

	internal static void A(int A)
	{
		try
		{
			FontColor.A(clsColors.ColorPalette[A].RGB);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	internal static void A(string A)
	{
		try
		{
			Color color = clsColors.RGB2Color(A);
			B(ColorTranslator.ToOle(color));
			K.Settings.LastFontColor = color;
			KH.A.InvalidateControl(clsColors.LAST_FONT_COLOR_BUTTON);
			Base.LogActivity(VH.A(149741));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	internal static void A()
	{
		try
		{
			B(ColorTranslator.ToOle(K.Settings.LastFontColor));
			Base.LogActivity(VH.A(149778));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	private static void B(int A)
	{
		object objectValue = RuntimeHelpers.GetObjectValue(MH.A.Application.Selection);
		if (objectValue is Range)
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
			Range range = JH.A((Range)null);
			if (!Base.IsWorksheetProtected(range.Worksheet))
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
				bool flag = false;
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
				range.Font.Color = A;
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
					JH.A(range, VH.A(60635));
				}
			}
			range = null;
		}
		else if (objectValue is DataLabels)
		{
			FontColor.A(((DataLabels)objectValue).Format, A);
		}
		else if (objectValue is DataLabel)
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
			FontColor.A(((DataLabel)objectValue).Format, A);
		}
		else if (objectValue is Axis)
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
			((Axis)objectValue).TickLabels.Font.Color = A;
		}
		else if (objectValue is AxisTitle)
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
			FontColor.A(((AxisTitle)objectValue).Format, A);
		}
		else if (objectValue is ChartTitle)
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
			FontColor.A(((ChartTitle)objectValue).Format, A);
		}
		else if (objectValue is Legend)
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
			FontColor.A(((Legend)objectValue).Format, A);
		}
		else if (objectValue is LegendEntry)
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
			FontColor.A(((LegendEntry)objectValue).Format, A);
		}
		else if (objectValue is DataTable)
		{
			FontColor.A(((DataTable)objectValue).Format, A);
		}
		objectValue = null;
	}

	private static void A(ChartFormat A, int B)
	{
		A.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = B;
	}
}
