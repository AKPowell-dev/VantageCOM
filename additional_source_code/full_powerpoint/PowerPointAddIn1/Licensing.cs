using System;
using A;
using MacabacusMacros.Auth;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1;

public sealed class Licensing
{
	public static void Authenticate()
	{
		Application application = NG.A.Application;
		Base.Authorize((Action)A, application.Version, (object)application.Build, AH.A(116727));
		application = null;
	}

	public static void Activate()
	{
		if (!Ribbon.ActivateProduct())
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
			A();
			return;
		}
	}

	private static void A()
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
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
					break;
				case 61:
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
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
							goto end_IL_0000_3;
						}
						goto default;
					}
					end_IL_0000_2:
					break;
				}
				num2 = 2;
				KG.A.InvalidateControl(AH.A(116748));
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 61;
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
	}

	public static bool AllowRestrictedMode()
	{
		return Access.AllowPowerPointOperation((PlanType)2, (Restriction)1, false);
	}

	public static bool AllowTemplateOperation()
	{
		return Access.AllowPowerPointOperation((PlanType)5, (Restriction)1, false);
	}

	public static bool AllowAgendaOperation()
	{
		return AllowTemplateOperation();
	}

	public static bool AllowMasterShapesOperation()
	{
		return AllowTemplateOperation();
	}

	public static bool AllowStylesOperation()
	{
		return AllowTemplateOperation();
	}

	public static bool AllowPaginationOperation()
	{
		return Access.AllowPowerPointOperation((PlanType)5, (Restriction)1, false);
	}

	public static bool AllowAdvancedShapeOperation()
	{
		return Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false);
	}

	public static bool AllowAdvancedTextOperation()
	{
		return Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false);
	}
}
