using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Auth;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class AlternateShading
{
	public enum AlternateShadingStyle
	{
		OddRows = 1,
		EvenRows,
		OddColumns,
		EvenColumns,
		None
	}

	[CompilerGenerated]
	private static int m_A;

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return AlternateShading.m_A;
		}
		[CompilerGenerated]
		set
		{
			AlternateShading.m_A = value;
		}
	}

	public static void Cycle()
	{
		if (!A())
		{
			return;
		}
		checked
		{
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
				CycleIndex++;
				AlternateShadingStyle a;
				switch (CycleIndex)
				{
				case 1:
					a = AlternateShadingStyle.OddRows;
					break;
				case 2:
					a = AlternateShadingStyle.EvenRows;
					break;
				case 3:
					a = AlternateShadingStyle.OddColumns;
					break;
				case 4:
					a = AlternateShadingStyle.EvenColumns;
					break;
				default:
					a = AlternateShadingStyle.None;
					CycleIndex = 0;
					break;
				}
				if (!A(a))
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
					if (CycleIndex != 1)
					{
						return;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						Base.LogActivity(VH.A(148093));
						return;
					}
				}
			}
		}
	}

	public static void ShadeRowsColumns(string strTag)
	{
		if (!A())
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
			AlternateShadingStyle alternateShadingStyle = (AlternateShadingStyle)Conversions.ToInteger(strTag);
			string strActivity = alternateShadingStyle switch
			{
				AlternateShadingStyle.OddRows => VH.A(148140), 
				AlternateShadingStyle.EvenRows => VH.A(148169), 
				AlternateShadingStyle.OddColumns => VH.A(148200), 
				AlternateShadingStyle.EvenColumns => VH.A(148235), 
				_ => VH.A(148272), 
			};
			if (!A(alternateShadingStyle))
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
				Base.LogActivity(strActivity);
				return;
			}
		}
	}

	private static bool A(AlternateShadingStyle A)
	{
		Application application = MH.A.Application;
		bool result = false;
		if (application.Selection is Range)
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
			Range range = (Range)application.Selection;
			if (!Base.IsWorksheetProtected(range.Worksheet))
			{
				application.ScreenUpdating = false;
				try
				{
					string text;
					string text2;
					string text3;
					string text4;
					FormatConditions formatConditions;
					checked
					{
						if (application.LanguageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI) == 1033)
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
							text = VH.A(148321);
							text2 = VH.A(47446);
							text3 = VH.A(47410);
							text4 = VH.A(148328);
						}
						else
						{
							Names names = application.ActiveWorkbook.Names;
							Name name = names.Add(VH.A(94040), VH.A(148345), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							string text5 = name.RefersToLocal.ToString();
							text = Strings.Mid(text5, 2, text5.IndexOf(VH.A(39848)) - 1);
							name.RefersTo = VH.A(148364);
							text5 = name.RefersToLocal.ToString();
							text2 = Strings.Mid(text5, 2, text5.IndexOf(VH.A(39848)) - 1);
							name.RefersTo = VH.A(148381);
							text5 = name.RefersToLocal.ToString();
							text3 = Strings.Mid(text5, 2, text5.IndexOf(VH.A(39848)) - 1);
							name.RefersTo = VH.A(148404);
							text5 = name.RefersToLocal.ToString();
							text4 = Strings.Mid(text5, 2, text5.IndexOf(VH.A(39848)) - 1);
							names.Item(VH.A(94040), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
						}
						formatConditions = range.FormatConditions;
						for (int i = formatConditions.Count; i >= 1; i += -1)
						{
							try
							{
								FormatCondition formatCondition = (FormatCondition)formatConditions.Item(i);
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
									if (!formatCondition.Formula1.Contains(VH.A(48936) + text + VH.A(39848) + text2 + VH.A(148435)))
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
										if (!formatCondition.Formula1.Contains(VH.A(48936) + text + VH.A(39848) + text3 + VH.A(148435)))
										{
											goto IL_0372;
										}
									}
									formatCondition.Delete();
								}
								goto IL_0372;
								IL_0372:
								if (formatCondition.Type == 2 && formatCondition.Formula1.Contains(VH.A(48936) + text + VH.A(39848) + text4 + VH.A(148440)))
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
									formatCondition.Delete();
								}
								formatCondition = null;
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
						}
					}
					if (A != AlternateShadingStyle.None)
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
						string listSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
						int num;
						if (!range.Worksheet.FilterMode)
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
							num = ((((Range)range.Cells[1, 1]).ListObject != null) ? 1 : 0);
						}
						else
						{
							num = 1;
						}
						bool flag = (byte)num != 0;
						string formula = default(string);
						switch (A)
						{
						case AlternateShadingStyle.OddRows:
							if (!flag)
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
								formula = VH.A(48936) + text + VH.A(39848) + text2 + VH.A(148435) + listSeparator + VH.A(148445);
							}
							else
							{
								formula = VH.A(48936) + text + VH.A(39848) + text4 + VH.A(148440) + listSeparator + VH.A(148454) + listSeparator + VH.A(148445);
							}
							break;
						case AlternateShadingStyle.EvenRows:
							formula = (flag ? (VH.A(48936) + text + VH.A(39848) + text4 + VH.A(148440) + listSeparator + VH.A(148454) + listSeparator + VH.A(148473)) : (VH.A(48936) + text + VH.A(39848) + text2 + VH.A(148435) + listSeparator + VH.A(148473)));
							break;
						case AlternateShadingStyle.OddColumns:
							formula = VH.A(48936) + text + VH.A(39848) + text3 + VH.A(148435) + listSeparator + VH.A(148445);
							break;
						case AlternateShadingStyle.EvenColumns:
							formula = VH.A(48936) + text + VH.A(39848) + text3 + VH.A(148435) + listSeparator + VH.A(148473);
							break;
						}
						formatConditions.Add(XlFormatConditionType.xlExpression, RuntimeHelpers.GetObjectValue(Missing.Value), formula, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						NewLateBinding.LateSetComplex(NewLateBinding.LateGet(formatConditions.Item(formatConditions.Count), null, VH.A(36170), new object[0], null, null, null), null, VH.A(55331), new object[1] { KH.A.DefaultFillColor }, null, null, OptimisticSet: false, RValueBase: true);
					}
					formatConditions = null;
					result = true;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					Base.LogException(ex4);
					ProjectData.ClearProjectError();
				}
				application.ScreenUpdating = true;
			}
			range = null;
		}
		application = null;
		return result;
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}
}
