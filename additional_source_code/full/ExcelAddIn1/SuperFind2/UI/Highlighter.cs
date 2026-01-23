using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.UI;

public sealed class Highlighter
{
	private static readonly int m_A = ColorTranslator.ToOle(Color.LawnGreen);

	private static readonly string m_A = VH.A(124844);

	[CompilerGenerated]
	private static Dictionary<Worksheet, Range> m_A;

	private static Dictionary<Worksheet, Range> HighlightedRanges
	{
		[CompilerGenerated]
		get
		{
			return Highlighter.m_A;
		}
		[CompilerGenerated]
		set
		{
			Highlighter.m_A = value;
		}
	} = null;

	internal static void A(Range A)
	{
		FormatConditions formatConditions = A.FormatConditions;
		formatConditions.Add(XlFormatConditionType.xlExpression, RuntimeHelpers.GetObjectValue(Missing.Value), Highlighter.m_A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		object instance = NewLateBinding.LateGet(formatConditions.Item(formatConditions.Count), null, VH.A(36170), new object[0], null, null, null);
		NewLateBinding.LateSetComplex(instance, null, VH.A(36187), new object[1] { Highlighter.m_A }, null, null, OptimisticSet: false, RValueBase: true);
		NewLateBinding.LateSetComplex(instance, null, VH.A(36212), new object[1] { XlPattern.xlPatternGray50 }, null, null, OptimisticSet: false, RValueBase: true);
		instance = null;
		_ = null;
		Worksheet worksheet = A.Worksheet;
		if (HighlightedRanges == null)
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
			HighlightedRanges = new Dictionary<Worksheet, Range>();
		}
		Dictionary<Worksheet, Range> highlightedRanges = HighlightedRanges;
		if (highlightedRanges.ContainsKey(worksheet))
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
			Dictionary<Worksheet, Range> dictionary = highlightedRanges;
			Worksheet key;
			Range A2 = dictionary[key = worksheet];
			RangeHelpers.A(ref A2, A);
			dictionary[key] = A2;
		}
		else
		{
			highlightedRanges.Add(worksheet, A);
		}
		highlightedRanges = null;
		worksheet = null;
	}

	internal static void A()
	{
		if (HighlightedRanges == null)
		{
			return;
		}
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
			application.ScreenUpdating = false;
			using (Dictionary<Worksheet, Range>.Enumerator enumerator = HighlightedRanges.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					B(enumerator.Current.Value);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0063;
					}
					continue;
					end_IL_0063:
					break;
				}
			}
			application.ScreenUpdating = true;
			application = null;
			HighlightedRanges = null;
			return;
		}
	}

	internal static void B(Range A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Cells.GetEnumerator();
			while (enumerator.MoveNext())
			{
				FormatConditions formatConditions = ((Range)enumerator.Current).FormatConditions;
				int num = formatConditions.Count;
				while (true)
				{
					if (num >= 1)
					{
						try
						{
							FormatCondition formatCondition = (FormatCondition)formatConditions.Item(num);
							if (formatCondition.Type == 2)
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
								if (Operators.CompareString(formatCondition.Formula1, Highlighter.m_A, TextCompare: false) == 0)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											formatCondition.Delete();
											goto end_IL_00a5;
										}
									}
								}
							}
							formatCondition = null;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						num = checked(num + -1);
						continue;
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
					break;
					continue;
					end_IL_00a5:
					break;
				}
				formatConditions = null;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					return;
				}
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
}
