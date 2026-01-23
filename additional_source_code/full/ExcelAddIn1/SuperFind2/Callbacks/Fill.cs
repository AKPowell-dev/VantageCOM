using System;
using System.Collections;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Media;
using A;
using ExcelAddIn1.Format;
using ExcelAddIn1.SuperFind2.Results;
using ExcelAddIn1.SuperFind2.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Callbacks;

public sealed class Fill
{
	internal static void A(WorksheetItem A, Range B)
	{
		Helpers.A(A, B, Fill.A);
	}

	private static bool A(Range A)
	{
		return FillColor.HasFill(A.Interior);
	}

	internal static void B(WorksheetItem A, Range B)
	{
		if (!Props.SearchForm.LookInEmptyCells)
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
			B = RangeHelpers.H(B);
		}
		if (B == null)
		{
			return;
		}
		string input = Props.SearchForm.Input1.Trim();
		Range A2 = null;
		Regex regex = new Regex(VH.A(102671), RegexOptions.IgnoreCase);
		Regex regex2 = new Regex(VH.A(102706));
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = B.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				if (!FillColor.HasFill(range.Interior))
				{
					continue;
				}
				System.Windows.Media.Color? color = null;
				Match match = regex.Match(input);
				if (match.Success)
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
					try
					{
						color = (System.Windows.Media.Color?)System.Windows.Media.ColorConverter.ConvertFromString(VH.A(49303) + match.Groups[1].Value);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				else
				{
					try
					{
						string[] array = Strings.Split(regex2.Replace(input, ""), VH.A(2378));
						color = checked(System.Windows.Media.Color.FromArgb(byte.MaxValue, (byte)Conversions.ToInteger(array[0]), (byte)Conversions.ToInteger(array[1]), (byte)Conversions.ToInteger(array[2])));
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
				}
				match = null;
				if (!color.HasValue)
				{
					continue;
				}
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
					byte value = (color?.R).Value;
					byte? obj;
					if (!color.HasValue)
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
						obj = null;
					}
					else
					{
						obj = color.GetValueOrDefault().G;
					}
					byte? b = obj;
					byte value2 = b.Value;
					byte? obj2;
					if (!color.HasValue)
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
						obj2 = null;
					}
					else
					{
						obj2 = color.GetValueOrDefault().B;
					}
					b = obj2;
					int num = ColorTranslator.ToOle(System.Drawing.Color.FromArgb(value, value2, b.Value));
					if (Operators.ConditionalCompareObjectEqual(range.Interior.Color, num, TextCompare: false))
					{
						RangeHelpers.A(ref A2, range);
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_029e;
				}
				continue;
				end_IL_029e:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		Helpers.B(A, A2);
		A2 = null;
		regex = null;
		regex2 = null;
	}
}
