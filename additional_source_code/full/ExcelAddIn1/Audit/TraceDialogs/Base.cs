using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using A;
using ExcelAddIn1.ExcelApp;
using ExcelAddIn1.View;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.TraceDialogs;

public sealed class Base
{
	public enum HighlightMode
	{
		Selection = 1,
		Crosshairs
	}

	public static readonly string CELL_REF_PATTERN = VH.A(41388);

	private static readonly int m_A = ColorTranslator.ToOle(Color.LawnGreen);

	private static readonly string m_A = VH.A(41594);

	[CompilerGenerated]
	private static Microsoft.Office.Interop.Excel.Workbook m_A;

	[CompilerGenerated]
	private static Worksheet m_A;

	[CompilerGenerated]
	private static Range m_A;

	[CompilerGenerated]
	private static Range m_B;

	[CompilerGenerated]
	private static ExpandedRange m_A;

	[CompilerGenerated]
	private static HighlightMode m_A;

	[CompilerGenerated]
	private static Regex m_A;

	[CompilerGenerated]
	private static Regex m_B;

	[CompilerGenerated]
	private static Regex m_C;

	private static bool m_A = false;

	private static bool m_B = false;

	private static readonly int m_B = 88;

	private static readonly int m_C = 90;

	private static Microsoft.Office.Interop.Excel.Workbook LastAuditedWorkbook
	{
		[CompilerGenerated]
		get
		{
			return Base.m_A;
		}
		[CompilerGenerated]
		set
		{
			Base.m_A = value;
		}
	}

	private static Worksheet LastAuditedSheet
	{
		[CompilerGenerated]
		get
		{
			return Base.m_A;
		}
		[CompilerGenerated]
		set
		{
			Base.m_A = value;
		}
	}

	private static Range LastAuditedCell
	{
		[CompilerGenerated]
		get
		{
			return Base.m_A;
		}
		[CompilerGenerated]
		set
		{
			Base.m_A = value;
		}
	}

	private static Range NavigatedCell
	{
		[CompilerGenerated]
		get
		{
			return Base.m_B;
		}
		[CompilerGenerated]
		set
		{
			Base.m_B = value;
		}
	} = null;

	private static ExpandedRange ExpandedCells
	{
		[CompilerGenerated]
		get
		{
			return Base.m_A;
		}
		[CompilerGenerated]
		set
		{
			Base.m_A = value;
		}
	} = null;

	private static HighlightMode SelectedHighlightMode
	{
		[CompilerGenerated]
		get
		{
			return Base.m_A;
		}
		[CompilerGenerated]
		set
		{
			Base.m_A = value;
		}
	} = HighlightMode.Crosshairs;

	private static Regex RegexA
	{
		[CompilerGenerated]
		get
		{
			return Base.m_A;
		}
		[CompilerGenerated]
		set
		{
			Base.m_A = value;
		}
	}

	private static Regex RegexB
	{
		[CompilerGenerated]
		get
		{
			return Base.m_B;
		}
		[CompilerGenerated]
		set
		{
			Base.m_B = value;
		}
	}

	private static Regex RegexC
	{
		[CompilerGenerated]
		get
		{
			return Base.m_C;
		}
		[CompilerGenerated]
		set
		{
			Base.m_C = value;
		}
	}

	private static bool A
	{
		get
		{
			return Base.m_A;
		}
		set
		{
			Base.m_A = value;
		}
	}

	public static bool WorkshareLoaded
	{
		get
		{
			return Base.m_B;
		}
		set
		{
			Base.m_B = value;
		}
	}

	public static void RecordLastAuditedCell(Microsoft.Office.Interop.Excel.Application xlApp)
	{
		LastAuditedWorkbook = xlApp.ActiveWorkbook;
		LastAuditedSheet = (Worksheet)xlApp.ActiveSheet;
		LastAuditedCell = xlApp.ActiveCell;
	}

	public static bool CanTrace(Microsoft.Office.Interop.Excel.Application xlApp)
	{
		if (xlApp.Selection is Range && !EditMode.IsEditMode(xlApp))
		{
			return A(xlApp);
		}
		return false;
	}

	public static void DisableNavAid()
	{
		if (!K.Settings.AuditHighlightCells)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (!NavAid.Enabled)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				NavAid.Disable();
				NavAid.RefreshRibbon();
				return;
			}
		}
	}

	private static bool A(Microsoft.Office.Interop.Excel.Application A)
	{
		if (A.ActiveWindow.SelectedSheets.Count > 1)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Forms.WarningMessage(VH.A(40467));
					return false;
				}
			}
		}
		return true;
	}

	public static void GoToLastAuditedCell()
	{
		try
		{
			LastAuditedWorkbook.Activate();
			LastAuditedSheet.Activate();
			LastAuditedCell.Select();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, VH.A(40588));
	}

	public static bool IsWorkshareLoaded(Microsoft.Office.Interop.Excel.Application xlApp)
	{
		bool result;
		try
		{
			COMAddIns cOMAddIns = xlApp.COMAddIns;
			object Index = VH.A(40623);
			result = cOMAddIns.Item(ref Index).Connect;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void CheckForWorkshare()
	{
		if (!WorkshareLoaded)
		{
			return;
		}
		if (IsWorkshareLoaded(MH.A.Application))
		{
			if (Base.A)
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
				Forms.WarningMessage(VH.A(40678));
				Base.A = true;
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)3, VH.A(41140));
				return;
			}
		}
		WorkshareLoaded = false;
	}

	public static void InitializeRegex()
	{
		RegexA = new Regex(VH.A(41201) + CELL_REF_PATTERN + VH.A(41262));
		RegexB = new Regex(VH.A(41265) + CELL_REF_PATTERN + VH.A(41262));
		RegexC = new Regex(VH.A(41312) + CELL_REF_PATTERN + VH.A(41262));
	}

	public static Range ResolveAddress(string strAddress, Range rngAudit)
	{
		if (RegexA.IsMatch(strAddress))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Match match = RegexA.Matches(strAddress)[0];
					return (Range)NewLateBinding.LateGet(rngAudit.Application.Workbooks[match.Groups[1].ToString()].Worksheets[A(match.Groups[2].ToString())], null, VH.A(41315), new object[1] { match.Groups[4].ToString() }, null, null, null);
				}
				}
			}
		}
		if (RegexB.IsMatch(strAddress))
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
				{
					Match match2 = RegexB.Matches(strAddress)[0];
					return (Range)NewLateBinding.LateGet(NewLateBinding.LateGet(rngAudit.Worksheet.Parent, null, VH.A(41326), new object[1] { A(match2.Groups[1].ToString()) }, null, null, null), null, VH.A(41315), new object[1] { match2.Groups[3].ToString() }, null, null, null);
				}
				}
			}
		}
		if (RegexC.IsMatch(strAddress))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
				{
					Match match3 = RegexC.Matches(strAddress)[0];
					return ((_Worksheet)rngAudit.Worksheet).get_Range((object)match3.Groups[1].ToString(), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				}
			}
		}
		Range result = default(Range);
		return result;
	}

	private static string A(string A)
	{
		return A.Trim('\'').Replace(VH.A(39854), VH.A(39851));
	}

	public static void GoToCell(Range rng, Range rngVisible, System.Windows.Controls.TreeView tree, bool blnUnhide, ref List<Worksheet> listHiddenSheets)
	{
		Microsoft.Office.Interop.Excel.Application application = rng.Application;
		Worksheet worksheet = rng.Worksheet;
		Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent;
		try
		{
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			if (application.ActiveWorkbook != workbook)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				workbook.Activate();
			}
			Worksheet worksheet2 = worksheet;
			if (worksheet2.Visible == XlSheetVisibility.xlSheetHidden)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				worksheet2.Visible = XlSheetVisibility.xlSheetVisible;
				listHiddenSheets.Add(worksheet);
			}
			if (application.ActiveSheet != worksheet)
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
				worksheet2.Select(RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			worksheet2 = null;
			if (blnUnhide)
			{
				A(rng);
			}
			if (((TraceItem)tree.SelectedItem).Level == 0)
			{
				ScrollTo(rngVisible);
			}
			else
			{
				Ranges.ScrollIntoView(rng);
			}
			rng.Select();
			if (!K.Settings.AuditHighlightCells)
			{
				return;
			}
			RemoveHighlight();
			if (Operators.ConditionalCompareObjectLess(rng.Rows.CountLarge, rng.Worksheet.Rows.CountLarge, TextCompare: false))
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
				if (Operators.ConditionalCompareObjectLess(rng.Columns.CountLarge, rng.Worksheet.Columns.CountLarge, TextCompare: false))
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
						{
							Range range;
							if (SelectedHighlightMode == HighlightMode.Crosshairs)
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
								range = application.Union(rng.EntireRow, rng.EntireColumn, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							else
							{
								range = rng;
							}
							FormatConditions formatConditions = range.FormatConditions;
							formatConditions.Add(XlFormatConditionType.xlExpression, RuntimeHelpers.GetObjectValue(Missing.Value), Base.m_A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							object instance = NewLateBinding.LateGet(formatConditions.Item(formatConditions.Count), null, VH.A(36170), new object[0], null, null, null);
							if (SelectedHighlightMode == HighlightMode.Crosshairs)
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
								NewLateBinding.LateSetComplex(instance, null, VH.A(36187), new object[1] { NavAid.PATTERN_COLOR_GREEN }, null, null, OptimisticSet: false, RValueBase: true);
							}
							else
							{
								NewLateBinding.LateSetComplex(instance, null, VH.A(36187), new object[1] { Base.m_A }, null, null, OptimisticSet: false, RValueBase: true);
							}
							NewLateBinding.LateSetComplex(instance, null, VH.A(36212), new object[1] { XlPattern.xlPatternGray50 }, null, null, OptimisticSet: false, RValueBase: true);
							instance = null;
							_ = null;
							range = null;
							NavigatedCell = rng;
							return;
						}
						}
					}
				}
			}
			NavigatedCell = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		finally
		{
			application.EnableEvents = true;
			application.ScreenUpdating = true;
			worksheet = null;
			workbook = null;
			application = null;
		}
	}

	public static void ScrollTo(Range rngVisible)
	{
		Pane activePane = rngVisible.Application.ActiveWindow.ActivePane;
		activePane.ScrollRow = Conversions.ToInteger(NewLateBinding.LateGet(rngVisible.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41347), new object[0], null, null, null));
		activePane.ScrollColumn = Conversions.ToInteger(NewLateBinding.LateGet(rngVisible.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41354), new object[0], null, null, null));
		_ = null;
	}

	public static void ReturnToPreviousRange(Range rng)
	{
		try
		{
			((Microsoft.Office.Interop.Excel.Workbook)rng.Worksheet.Parent).Activate();
			rng.Worksheet.Activate();
			rng.Select();
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static void RemoveHighlight()
	{
		if (NavigatedCell == null)
		{
			return;
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Range range;
			if (SelectedHighlightMode == HighlightMode.Crosshairs)
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
				range = NavigatedCell.Application.Union(NavigatedCell.EntireRow, NavigatedCell.EntireColumn, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			else
			{
				range = NavigatedCell;
			}
			FormatConditions formatConditions = range.FormatConditions;
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
								switch (5)
								{
								case 0:
									continue;
								}
								break;
							}
							if (Operators.CompareString(formatCondition.Formula1, Base.m_A, TextCompare: false) == 0)
							{
								formatCondition.Delete();
								break;
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
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
				break;
			}
			formatConditions = null;
			range = null;
			NavigatedCell = null;
			return;
		}
	}

	private static void A(Range A)
	{
		CollapseExpandedCells();
		bool flag = Ranges.HasHiddenRows(A);
		bool flag2 = Ranges.HasHiddenColumns(A);
		if (!flag)
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
			if (!flag2)
			{
				return;
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
		ExpandedCells = new ExpandedRange(A, flag, flag2);
		D(A);
	}

	public static void CollapseExpandedCells()
	{
		if (ExpandedCells == null)
		{
			return;
		}
		try
		{
			ExpandedRange expandedCells = ExpandedCells;
			if (expandedCells.RowsExpanded)
			{
				B(expandedCells.Range);
			}
			if (expandedCells.ColumnsExpanded)
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
				C(expandedCells.Range);
			}
			expandedCells = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		ExpandedCells = null;
	}

	private static void B(Range A)
	{
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		XlCalculation calculation = application.Calculation;
		application.Calculation = XlCalculation.xlCalculationManual;
		Range range = JH.A(A, application);
		if (range != null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range.Rows.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range2 = (Range)enumerator.Current;
					if (!Operators.ConditionalCompareObjectGreater(range2.OutlineLevel, 1, TextCompare: false))
					{
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
					try
					{
						range2.ShowDetail = false;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_00a6;
					}
					continue;
					end_IL_00a6:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			range = null;
		}
		application.Calculation = calculation;
		application = null;
	}

	private static void C(Range A)
	{
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		XlCalculation calculation = application.Calculation;
		application.Calculation = XlCalculation.xlCalculationManual;
		Range range = JH.A(A, application);
		if (range != null)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			foreach (Range column in range.Columns)
			{
				if (!Operators.ConditionalCompareObjectGreater(column.OutlineLevel, 1, TextCompare: false))
				{
					continue;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					break;
				}
				try
				{
					column.ShowDetail = false;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			range = null;
		}
		application.Calculation = calculation;
		application = null;
	}

	internal static void D(Range A)
	{
		Range range = null;
		bool flag = false;
		Range range2 = A;
		try
		{
			range2.EntireRow.ShowDetail = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			try
			{
				range = JH.A(A, (Microsoft.Office.Interop.Excel.Application)null);
				flag = true;
				if (range != null)
				{
					IEnumerator enumerator = default(IEnumerator);
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
						try
						{
							enumerator = range.Rows.GetEnumerator();
							while (enumerator.MoveNext())
							{
								((Range)enumerator.Current).ShowDetail = true;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_007d;
								}
								continue;
								end_IL_007d:
								break;
							}
						}
						finally
						{
							if (enumerator is IDisposable)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									(enumerator as IDisposable).Dispose();
									break;
								}
							}
						}
						break;
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			ProjectData.ClearProjectError();
		}
		range2.EntireRow.Hidden = false;
		try
		{
			range2.EntireColumn.ShowDetail = true;
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			try
			{
				if (!flag)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
					range = JH.A(A, (Microsoft.Office.Interop.Excel.Application)null);
				}
				if (range != null)
				{
					IEnumerator enumerator2 = default(IEnumerator);
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						try
						{
							enumerator2 = range.Columns.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								((Range)enumerator2.Current).ShowDetail = true;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
						break;
					}
				}
			}
			catch (Exception ex7)
			{
				ProjectData.SetProjectError(ex7);
				Exception ex8 = ex7;
				ProjectData.ClearProjectError();
			}
			ProjectData.ClearProjectError();
		}
		range2.EntireColumn.Hidden = false;
		range = null;
		range2 = null;
	}

	public static List<string> UnhideHiddenSheets()
	{
		List<string> list = new List<string>();
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = MH.A.Application.ActiveWorkbook;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = activeWorkbook.Worksheets.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Worksheet worksheet = (Worksheet)enumerator.Current;
				if (worksheet.Visible == XlSheetVisibility.xlSheetHidden)
				{
					try
					{
						worksheet.Visible = XlSheetVisibility.xlSheetVisible;
						list.Add(worksheet.Name);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		activeWorkbook = null;
		return list;
	}

	public static void HidePreviouslyHiddenSheets(List<string> HiddenSheetNames)
	{
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = MH.A.Application.ActiveWorkbook;
		using (List<string>.Enumerator enumerator = HiddenSheetNames.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				string current = enumerator.Current;
				try
				{
					NewLateBinding.LateSetComplex(activeWorkbook.Worksheets[current], null, VH.A(41367), new object[1] { XlSheetVisibility.xlSheetHidden }, null, null, OptimisticSet: false, RValueBase: true);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
		}
		activeWorkbook = null;
	}

	public static void HideSheets(List<Worksheet> listHiddenSheets, Microsoft.Office.Interop.Excel.Application xlApp, bool blnLeaveActiveSheetOpen)
	{
		xlApp.ScreenUpdating = false;
		foreach (Worksheet listHiddenSheet in listHiddenSheets)
		{
			if (blnLeaveActiveSheetOpen)
			{
				if (xlApp.ActiveSheet == listHiddenSheet)
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
			}
			try
			{
				listHiddenSheet.Visible = XlSheetVisibility.xlSheetHidden;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		xlApp.ScreenUpdating = true;
	}

	public static void SetFormSizeAndPosition(Form frm)
	{
		MySettings settings = K.Settings;
		frm.Height = settings.AuditFormHeight;
		frm.Width = settings.AuditFormWidth;
		frm.Top = settings.AuditFormTop;
		frm.Left = settings.AuditFormLeft;
		if (!A(frm.RectangleToScreen(frm.ClientRectangle)))
		{
			frm.Top = 0;
			frm.Left = 0;
		}
		settings = null;
	}

	public static void SaveFormSizeAndPosition(Form frm)
	{
		MySettings settings = K.Settings;
		settings.AuditFormHeight = frm.Height;
		settings.AuditFormWidth = frm.Width;
		settings.AuditFormTop = Math.Max(frm.Top, 0);
		settings.AuditFormLeft = Math.Max(frm.Left, 0);
		_ = null;
	}

	private static bool A(Rectangle A)
	{
		Screen[] allScreens = Screen.AllScreens;
		for (int i = 0; i < allScreens.Length; i = checked(i + 1))
		{
			if (!allScreens[i].Bounds.IntersectsWith(A))
			{
				continue;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				return true;
			}
		}
		while (true)
		{
			switch (1)
			{
			case 0:
				continue;
			}
			return false;
		}
	}

	public static void MoveFormAsNeeded(Form frm, Range rng)
	{
		if (rng.Application.ActiveWindow.Split)
		{
			return;
		}
		checked
		{
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
				Rectangle rect = frm.RectangleToScreen(frm.ClientRectangle);
				Rectangle rect2 = A(rng);
				if (!rect2.IntersectsWith(rect))
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
					if (!rect.IntersectsWith(rect2))
					{
						goto IL_02f7;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				int num = rect.X;
				int width = rect.Width;
				int num2 = rect.Y;
				int height = rect.Height;
				Rectangle workingArea = Screen.FromControl(frm).WorkingArea;
				bool flag = rect2.X + rect2.Width + 5 + width <= workingArea.Width;
				bool flag2 = rect2.X - 5 - width >= 0;
				bool flag3 = rect2.Y + rect2.Height + 5 + height <= workingArea.Height;
				bool flag4 = rect2.Y - 5 - height >= 0;
				workingArea = default(Rectangle);
				int num3 = Math.Abs(num - (rect2.X + rect2.Width + 5));
				int num4 = Math.Abs(num - (rect2.X - 5 - width));
				int num5 = Math.Abs(num2 - (rect2.Y + rect2.Height + 5));
				int num6 = Math.Abs(num2 - (rect2.Y - 5 - height));
				List<int> list = new List<int>(new int[4] { num3, num4, num5, num6 });
				list.Sort();
				foreach (int item in list)
				{
					if (item == num3)
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
						if (flag)
						{
							num = rect2.X + rect2.Width + 5;
							goto IL_0289;
						}
					}
					if (item == num4 && flag2)
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
						num = rect2.X - 5 - width;
					}
					else
					{
						if (item == num5)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								break;
							}
							if (flag3)
							{
								num2 = rect2.Y + rect2.Height + 5;
								goto IL_0289;
							}
						}
						if (item == num6)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								break;
							}
							if (flag4)
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
								num2 = rect2.Y - 5 - height;
							}
						}
					}
					goto IL_0289;
					IL_0289:
					rect.X = num;
					rect.Y = num2;
					if (rect2.IntersectsWith(rect))
					{
						continue;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						break;
					}
					if (rect.IntersectsWith(rect2))
					{
						continue;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_02bd;
						}
						continue;
						end_IL_02bd:
						break;
					}
					break;
				}
				list = null;
				frm.Location = new System.Drawing.Point(num, num2);
				goto IL_02f7;
				IL_02f7:
				rect = default(Rectangle);
				rect2 = default(Rectangle);
				return;
			}
		}
	}

	[DllImport("gdi32.dll", EntryPoint = "GetDeviceCaps")]
	private static extern int A(IntPtr A, int B);

	[DllImport("user32.dll", EntryPoint = "GetDC")]
	private static extern IntPtr B(IntPtr A);

	[DllImport("user32.dll", EntryPoint = "ReleaseDC")]
	private static extern bool B(IntPtr A, IntPtr B);

	private static Rectangle A(Range A)
	{
		Microsoft.Office.Interop.Excel.Application application = A.Application;
		_ = A.Worksheet;
		Microsoft.Office.Interop.Excel.Window activeWindow = application.ActiveWindow;
		IntPtr intPtr = B(IntPtr.Zero);
		long num = Base.A(intPtr, Base.m_B);
		long num2 = Base.A(intPtr, Base.m_C);
		B(IntPtr.Zero, intPtr);
		double num3 = application.InchesToPoints(1.0);
		double num4 = Conversions.ToDouble(Operators.DivideObject(activeWindow.Zoom, 100));
		int num5 = activeWindow.PointsToScreenPixelsX(0);
		int num6 = activeWindow.PointsToScreenPixelsY(0);
		int num7 = Convert.ToInt32(Operators.AddObject(num5, Operators.DivideObject(Operators.MultiplyObject(Operators.MultiplyObject(A.Left, num4), num), num3)));
		checked
		{
			int width;
			try
			{
				width = Convert.ToInt32(Operators.AddObject(num5, Operators.DivideObject(Operators.MultiplyObject(Operators.MultiplyObject(((Range)A.Cells[1, RuntimeHelpers.GetObjectValue(A.Columns.CountLarge)]).get_Offset((object)0, (object)1).Left, num4), num), num3))) - num7;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				width = 10000;
				ProjectData.ClearProjectError();
			}
			int num8 = Convert.ToInt32(Operators.AddObject(num6, Operators.DivideObject(Operators.MultiplyObject(Operators.MultiplyObject(A.Top, num4), num2), num3)));
			int height;
			try
			{
				height = Convert.ToInt32(Operators.AddObject(num6, Operators.DivideObject(Operators.MultiplyObject(Operators.MultiplyObject(((Range)A.Cells[RuntimeHelpers.GetObjectValue(A.Rows.CountLarge), 1]).get_Offset((object)1, (object)0).Top, num4), num2), num3))) - num8;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				height = 10000;
				ProjectData.ClearProjectError();
			}
			application = null;
			return new Rectangle(num7, num8, width, height);
		}
	}

	public static bool ProcessShortcut(Form frm, System.Windows.Input.KeyEventArgs e)
	{
		Form form = frm;
		checked
		{
			if (System.Windows.Input.Keyboard.Modifiers == ModifierKeys.Control)
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
				switch (e.Key)
				{
				case Key.Up:
					form.Top = Math.Max(0, form.Top - 50);
					return true;
				case Key.Down:
					form.Top = Math.Min(Screen.FromControl(frm).WorkingArea.Height - form.Height, form.Top + 50);
					return true;
				case Key.Left:
					form.Left = Math.Max(0, form.Left - 50);
					return true;
				case Key.Right:
					form.Left = Math.Min(Screen.FromControl(frm).WorkingArea.Width - form.Width, form.Left + 50);
					return true;
				case Key.Home:
				{
					Rectangle workingArea2 = Screen.FromControl(frm).WorkingArea;
					frm.Top = workingArea2.Top;
					frm.Left = workingArea2.Left;
					return true;
				}
				case Key.End:
				{
					Rectangle workingArea = Screen.FromControl(frm).WorkingArea;
					frm.Top = workingArea.Top + workingArea.Height - frm.Height;
					frm.Left = workingArea.Left + workingArea.Width - frm.Width;
					return true;
				}
				}
			}
			else if (System.Windows.Input.Keyboard.Modifiers == ModifierKeys.Shift)
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
				switch (e.Key)
				{
				case Key.Down:
					form.Height += 50;
					form.Invalidate();
					return true;
				case Key.Up:
					form.Height -= 50;
					form.Invalidate();
					return true;
				case Key.Right:
					form.Width += 50;
					form.Invalidate();
					return true;
				case Key.Left:
					form.Width -= 50;
					form.Invalidate();
					return true;
				}
			}
			form = null;
			return false;
		}
	}

	public static double TreeViewSizeChanged(SizeChangedEventArgs e, ScrollViewer scroll)
	{
		return scroll.ViewportWidth - e.PreviousSize.Width + e.NewSize.Width;
	}

	public static bool ScrollRightRequired(TextPointer tp, ScrollViewer sv)
	{
		return tp.GetCharacterRect(LogicalDirection.Forward).Location.X > sv.ActualWidth + sv.HorizontalOffset;
	}

	public static bool ScrollLeftRequired(TextPointer tp, ScrollViewer sv)
	{
		return tp.GetCharacterRect(LogicalDirection.Forward).Location.X < sv.HorizontalOffset;
	}

	public static void ScrollFormula(TextPointer tp, ScrollViewer sv)
	{
		sv.ScrollToHorizontalOffset(tp.GetPositionAtOffset(-2).GetCharacterRect(LogicalDirection.Forward).Location.X);
	}

	public static string CleanValueText(string str)
	{
		return Strings.Trim(str).Replace(VH.A(41382), VH.A(41385));
	}

	public static string RemoveNewlines(string strFormula)
	{
		return strFormula.Replace(VH.A(41382), "");
	}
}
