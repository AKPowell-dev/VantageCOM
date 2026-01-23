using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Format;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Formulas;

public sealed class AutoFill
{
	[CompilerGenerated]
	private static int A;

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public static void Dates()
	{
		int try0000_dispatch = -1;
		Application application = default(Application);
		int num2 = default(int);
		int num = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					switch (try0000_dispatch)
					{
					default:
					{
						application = MH.A.Application;
						ProjectData.ClearProjectError();
						num2 = 2;
						Application application2 = application;
						application2.ScreenUpdating = false;
						application2.EnableEvents = false;
						if (Operators.ConditionalCompareObjectGreater(NewLateBinding.LateGet(NewLateBinding.LateGet(application2.Selection, null, VH.A(152043), new object[0], null, null, null), null, VH.A(152052), new object[0], null, null, null), 1, TextCompare: false))
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
							if (Operators.ConditionalCompareObjectGreater(NewLateBinding.LateGet(NewLateBinding.LateGet(application2.Selection, null, VH.A(152073), new object[0], null, null, null), null, VH.A(152052), new object[0], null, null, null), 1, TextCompare: false))
							{
								Interaction.MsgBox(VH.A(152088), MsgBoxStyle.Critical, VH.A(40448));
								break;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						if (string.IsNullOrEmpty(Conversions.ToString(application2.ActiveCell.Text)) | !Information.IsDate(RuntimeHelpers.GetObjectValue(application2.ActiveCell.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
							application2.ActiveCell.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)DateAndTime.DateString);
						}
						switch (CycleIndex)
						{
						case 0:
							application2.ActiveCell.AutoFill((Range)application2.Selection, XlAutoFillType.xlFillWeekdays);
							CycleIndex++;
							break;
						case 1:
							application2.ActiveCell.AutoFill((Range)application2.Selection, XlAutoFillType.xlFillMonths);
							CycleIndex++;
							break;
						default:
							application2.ActiveCell.AutoFill((Range)application2.Selection, XlAutoFillType.xlFillYears);
							CycleIndex = 0;
							break;
						}
						application2 = null;
						AutoColor.Selection();
						break;
					}
					case 543:
						num = -1;
						switch (num2)
						{
						case 2:
							break;
						default:
							goto end_IL_0000;
						}
						break;
					}
					application.EnableEvents = true;
					application.ScreenUpdating = true;
					application = null;
					break;
				}
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num2 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 543;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void Dates(IRibbonControl control)
	{
		int try0000_dispatch = -1;
		Application application = default(Application);
		int num2 = default(int);
		int num = default(int);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
				{
					application = MH.A.Application;
					ProjectData.ClearProjectError();
					num2 = 2;
					string tag = control.Tag;
					XlAutoFillType type;
					if (Operators.CompareString(tag, VH.A(152234), TextCompare: false) != 0)
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
						if (Operators.CompareString(tag, VH.A(47072), TextCompare: false) != 0)
						{
							if (Operators.CompareString(tag, VH.A(152251), TextCompare: false) != 0)
							{
								goto end_IL_0000;
							}
							type = XlAutoFillType.xlFillYears;
						}
						else
						{
							type = XlAutoFillType.xlFillMonths;
						}
					}
					else
					{
						type = XlAutoFillType.xlFillWeekdays;
					}
					Application application2 = application;
					application2.ScreenUpdating = false;
					application2.EnableEvents = false;
					if (Operators.ConditionalCompareObjectGreater(NewLateBinding.LateGet(NewLateBinding.LateGet(application2.Selection, null, VH.A(152043), new object[0], null, null, null), null, VH.A(152052), new object[0], null, null, null), 1, TextCompare: false))
					{
						if (Operators.ConditionalCompareObjectGreater(NewLateBinding.LateGet(NewLateBinding.LateGet(application2.Selection, null, VH.A(152073), new object[0], null, null, null), null, VH.A(152052), new object[0], null, null, null), 1, TextCompare: false))
						{
							Interaction.MsgBox(VH.A(152088), MsgBoxStyle.Critical, VH.A(40448));
							break;
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
					if (string.IsNullOrEmpty(Conversions.ToString(application2.ActiveCell.Text)))
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
						application2.ActiveCell.set_Value(RuntimeHelpers.GetObjectValue(Missing.Value), (object)DateAndTime.DateString);
					}
					if (Operators.ConditionalCompareObjectGreater(NewLateBinding.LateGet(NewLateBinding.LateGet(application2.Selection, null, VH.A(62391), new object[0], null, null, null), null, VH.A(152052), new object[0], null, null, null), 1, TextCompare: false))
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
						application2.ActiveCell.AutoFill((Range)application2.Selection, type);
					}
					AutoColor.Selection();
					application2 = null;
					break;
				}
				case 580:
					num = -1;
					switch (num2)
					{
					case 2:
						break;
					default:
						goto end_IL_0000_2;
					}
					break;
				}
				application.EnableEvents = true;
				application.ScreenUpdating = true;
				application = null;
				break;
				end_IL_0000_2:;
			}
			catch (object obj) when (obj is Exception && num2 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 580;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}
}
