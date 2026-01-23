using System;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Threading;
using A;
using MacabacusMacros.Auth;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

public sealed class Licensing
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static global::A.A A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal void A()
		{
			KB.C();
		}
	}

	public static void Authenticate()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Base.Authorize((Action)A, application.Version, (object)application.Build, VH.A(169659));
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
			switch (6)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			B();
			return;
		}
	}

	private static void A()
	{
		B();
		System.Windows.Application current = System.Windows.Application.Current;
		if (current == null)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return;
				}
			}
		}
		Dispatcher dispatcher = current.Dispatcher;
		if (dispatcher == null)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		global::A.A method;
		if (_Closure_0024__.A == null)
		{
			method = (_Closure_0024__.A = [SpecialName] () =>
			{
				KB.C();
			});
		}
		else
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
			method = _Closure_0024__.A;
		}
		dispatcher.BeginInvoke(method);
	}

	private static void B()
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
				KH.A.InvalidateControl(VH.A(169670));
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static bool AllowRestrictedMode()
	{
		return Access.AllowExcelOperation((PlanType)2, (Restriction)1, false);
	}

	public static bool AllowQuickChartOperation()
	{
		return Access.AllowExcelOperation((PlanType)5, (Restriction)2, false);
	}

	public static bool AllowChartAddOnOperation()
	{
		return Access.AllowExcelOperation((PlanType)5, (Restriction)1, false);
	}

	public static bool AllowOptimizationOperation()
	{
		return Access.AllowExcelOperation((PlanType)5, (Restriction)2, false);
	}

	public static bool AllowVisualizationOperation()
	{
		return Access.AllowExcelOperation((PlanType)5, (Restriction)1, false);
	}

	public static bool AllowAdvancedViewOperation()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}
}
