using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Media;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs.Precedents;

public sealed class Dialog
{
	public enum CVErrEnum
	{
		ErrDiv0 = -2146826281,
		ErrNA = -2146826246,
		ErrName = -2146826259,
		ErrNull = -2146826288,
		ErrNum = -2146826252,
		ErrRef = -2146826265,
		ErrValue = -2146826273
	}

	[CompilerGenerated]
	private static List<SolidColorBrush> m_A;

	[CompilerGenerated]
	private static Dictionary<CVErrEnum, string> m_A;

	[CompilerGenerated]
	private static Dictionary<string, List<string>> m_A;

	public static List<SolidColorBrush> FormulaBrushes
	{
		[CompilerGenerated]
		get
		{
			return Dialog.m_A;
		}
		[CompilerGenerated]
		set
		{
			Dialog.m_A = value;
		}
	} = null;

	public static Dictionary<CVErrEnum, string> FormulaErrors
	{
		[CompilerGenerated]
		get
		{
			return Dialog.m_A;
		}
		[CompilerGenerated]
		set
		{
			Dialog.m_A = value;
		}
	} = null;

	public static Dictionary<string, List<string>> ArgumentNames
	{
		[CompilerGenerated]
		get
		{
			return Dialog.m_A;
		}
		[CompilerGenerated]
		set
		{
			Dialog.m_A = value;
		}
	} = null;

	public static void Show()
	{
		if (!Licensing.AllowRestrictedMode())
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (Base.CanTrace(application))
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
				try
				{
					if (Conversions.ToBoolean(application.ActiveCell.HasFormula))
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							Base.RecordLastAuditedCell(application);
							break;
						}
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				if (System.Windows.Forms.Application.OpenForms.OfType<frmPrecedentsHost>().Any())
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
					System.Windows.Forms.Application.OpenForms.OfType<frmPrecedentsHost>().ElementAt(0).Close();
				}
				if (FormulaBrushes == null)
				{
					FormulaBrushes = new List<SolidColorBrush>(new SolidColorBrush[7]
					{
						A(95, 140, 237),
						A(235, 94, 96),
						A(141, 97, 194),
						A(45, 150, 57),
						A(191, 76, 145),
						A(227, 130, 34),
						A(55, 127, 158)
					});
				}
				if (FormulaErrors == null)
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
					FormulaErrors = new Dictionary<CVErrEnum, string>();
					Dictionary<CVErrEnum, string> formulaErrors = FormulaErrors;
					formulaErrors.Add(CVErrEnum.ErrDiv0, VH.A(44078));
					formulaErrors.Add(CVErrEnum.ErrValue, VH.A(44093));
					formulaErrors.Add(CVErrEnum.ErrRef, VH.A(44108));
					formulaErrors.Add(CVErrEnum.ErrNA, VH.A(44119));
					formulaErrors.Add(CVErrEnum.ErrNum, VH.A(44128));
					formulaErrors.Add(CVErrEnum.ErrName, VH.A(44139));
					formulaErrors.Add(CVErrEnum.ErrNull, VH.A(44152));
					_ = null;
				}
				Base.DisableNavAid();
				Base.InitializeRegex();
				frmPrecedentsHost frmPrecedentsHost = new frmPrecedentsHost();
				try
				{
					frmPrecedentsHost.Show();
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				frmPrecedentsHost = null;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, VH.A(44165));
			}
			application = null;
			return;
		}
	}

	private static SolidColorBrush A(int A, int B, int C)
	{
		SolidColorBrush solidColorBrush = new SolidColorBrush(checked(Color.FromRgb((byte)A, (byte)B, (byte)C)));
		solidColorBrush.Freeze();
		return solidColorBrush;
	}

	public static bool IsFormulaError(Range rng)
	{
		bool result;
		try
		{
			WorksheetFunction worksheetFunction = rng.Application.WorksheetFunction;
			if (worksheetFunction.IsError(worksheetFunction.Sum(rng, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value))))
			{
				result = true;
				goto IL_0192;
			}
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = true;
			ProjectData.ClearProjectError();
			goto IL_0192;
		}
		result = false;
		goto IL_0192;
		IL_0192:
		return result;
	}
}
