using System;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

public sealed class Validation
{
	public static void Toggle()
	{
		A((Action<Range>)Validation.ApplyToggle);
	}

	public static void Number()
	{
		A((Action<Range>)Validation.ApplyNumber);
	}

	public static void DateInput()
	{
		A((Action<Range>)Validation.ApplyDateInput);
	}

	public static void Text()
	{
		A((Action<Range>)Validation.ApplyText);
	}

	public static void ZeroOrLarger()
	{
		A((Action<Range>)Validation.ApplyZeroOrLarger);
	}

	public static void ZeroOrSmaller()
	{
		A((Action<Range>)Validation.ApplyZeroOrSmaller);
	}

	public static void PositivePercent()
	{
		A((Action<Range>)Validation.ApplyPositivePercent);
	}

	public static void AnyPercent()
	{
		A((Action<Range>)Validation.ApplyAnyPercent);
	}

	private static void A(Action<Range> A)
	{
		if (!Validation.A())
		{
			return;
		}
		Application application = MH.A.Application;
		if (application.Selection is Range)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Range range = (Range)application.Selection;
			application.ScreenUpdating = false;
			try
			{
				range.Validation.Delete();
				A(range);
				Validation.SetOtherValidationProperties(range);
				Validation.A(VH.A(91151));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			range = null;
		}
		application = null;
	}

	public static void Clear()
	{
		if (!A())
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
			if (application.Selection is Range)
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
				Range range = (Range)application.Selection;
				application.ScreenUpdating = false;
				try
				{
					range.Validation.Delete();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				application.ScreenUpdating = true;
				range = null;
				A(VH.A(91180));
			}
			application = null;
			return;
		}
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)2, A);
	}
}
