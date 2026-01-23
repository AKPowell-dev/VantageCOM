using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Base
{
	public static bool IsString(Range rng)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		bool result = default(bool);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 114:
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
							goto IL_0007;
						case 3:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 4:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0007:
					num2 = 2;
					if (!Operators.ConditionalCompareObjectNotEqual(Conversion.Val(RuntimeHelpers.GetObjectValue(rng.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))), rng.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)), TextCompare: false))
					{
						goto end_IL_0000_3;
					}
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 3;
				result = true;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 114;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool IsWorksheetProtected(Worksheet ws)
	{
		bool protectContents = ws.ProtectContents;
		if (protectContents)
		{
			Forms.WarningMessage(VH.A(148526));
		}
		return protectContents;
	}

	public static void HandleFormattingException(Exception ex)
	{
		if (ex is COMException)
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
			if (!ex.Message.ToLower().Contains(VH.A(148573)))
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
				if (!ex.Message.ToLower().Contains(VH.A(148636)))
				{
					goto IL_0096;
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
			}
			MessageBox.Show(ex.Message, VH.A(43304), MessageBoxButtons.OK, MessageBoxIcon.Hand);
			return;
		}
		goto IL_0096;
		IL_0096:
		clsReporting.LogException(ex);
	}

	public static void LogActivity(string strActivity)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, strActivity);
	}

	public static void LogException(Exception ex)
	{
		clsReporting.LogException(ex);
	}
}
