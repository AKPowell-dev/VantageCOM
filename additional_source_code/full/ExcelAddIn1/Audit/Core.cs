using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit;

public sealed class Core
{
	public static bool HasDependents(Range rngCell)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		bool flag = default(bool);
		int num = default(int);
		int num3 = default(int);
		Microsoft.Office.Interop.Excel.Workbook workbook = default(Microsoft.Office.Interop.Excel.Workbook);
		int num5 = default(int);
		bool result = default(bool);
		Range range = default(Range);
		Range range2 = default(Range);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					flag = false;
					goto IL_0005;
				case 498:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0005;
						case 3:
							goto IL_000a;
						case 4:
							goto IL_0011;
						case 5:
							goto IL_0024;
						case 6:
							goto IL_002b;
						case 8:
							goto IL_0034;
						case 9:
							goto IL_003c;
						case 10:
							goto IL_004e;
						case 11:
							goto IL_0054;
						case 12:
							goto IL_0067;
						case 13:
							goto IL_008c;
						case 14:
							goto IL_0111;
						case 16:
							goto IL_0119;
						case 15:
						case 17:
							goto IL_012d;
						case 18:
							goto IL_013e;
						case 19:
							goto IL_0141;
						case 20:
							goto IL_0161;
						case 21:
							goto IL_016b;
						case 22:
							goto IL_0171;
						case 7:
						case 23:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 24:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0141:
					num2 = 19;
					if (rngCell.Application.ActiveWorkbook != workbook)
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
						goto IL_0161;
					}
					goto IL_016b;
					IL_0005:
					num2 = 2;
					num5 = 0;
					goto IL_000a;
					IL_000a:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0011;
					IL_0011:
					num2 = 4;
					num5 = rngCell.DirectDependents.Count;
					goto IL_0024;
					IL_0024:
					num2 = 5;
					if (num5 > 0)
					{
						goto IL_002b;
					}
					goto IL_0034;
					IL_002b:
					num2 = 6;
					result = true;
					goto end_IL_0000_3;
					IL_0034:
					num2 = 8;
					if (rngCell == null)
					{
						break;
					}
					goto IL_003c;
					IL_003c:
					num2 = 9;
					workbook = rngCell.Application.ActiveWorkbook;
					goto IL_004e;
					IL_004e:
					num2 = 10;
					range = rngCell;
					goto IL_0054;
					IL_0054:
					num2 = 11;
					range.ShowDependents(false);
					goto IL_0067;
					IL_0067:
					num2 = 12;
					range2 = (Range)range.NavigateArrow(false, 1, 1);
					goto IL_008c;
					IL_008c:
					num2 = 13;
					if (Operators.CompareString(range2.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), rngCell.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false) == 0)
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
						goto IL_0111;
					}
					goto IL_0119;
					IL_0171:
					num2 = 22;
					workbook = null;
					break;
					IL_0161:
					num2 = 20;
					workbook.Activate();
					goto IL_016b;
					IL_016b:
					num2 = 21;
					range2 = null;
					goto IL_0171;
					IL_0111:
					num2 = 14;
					flag = false;
					goto IL_012d;
					IL_0119:
					num2 = 16;
					flag = Information.Err().Number == 0;
					goto IL_012d;
					IL_012d:
					num2 = 17;
					range.ShowDependents(true);
					goto IL_013e;
					IL_013e:
					range = null;
					goto IL_0141;
					end_IL_0000_2:
					break;
				}
				num2 = 23;
				result = flag;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 498;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
