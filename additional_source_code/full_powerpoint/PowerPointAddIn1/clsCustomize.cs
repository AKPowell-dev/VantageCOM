using System;
using System.Windows.Forms;
using A;
using MacabacusMacros.Auth;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1;

public sealed class clsCustomize
{
	public static bool SlideMaster()
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
				case 126:
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
							goto IL_0025;
						case 5:
							goto IL_002b;
						case 6:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 4:
						case 7:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0025:
					num2 = 3;
					result = true;
					goto end_IL_0000_3;
					IL_0007:
					num2 = 2;
					if (Base.IsUserAdmin())
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
						goto IL_0025;
					}
					goto IL_002b;
					IL_002b:
					num2 = 5;
					MessageBox.Show(AH.A(151528), AH.A(5874), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 6;
				result = false;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 126;
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
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
