using System;
using System.Collections;
using System.Runtime.CompilerServices;
using A;
using ExcelAddIn1.Shapes;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class CellSize
{
	public enum ConformSizeBehavior
	{
		Columns,
		Rows,
		Both
	}

	[CompilerGenerated]
	private static int m_A;

	[CompilerGenerated]
	private static int m_B;

	internal static int RowCycleIndex
	{
		[CompilerGenerated]
		get
		{
			return CellSize.m_A;
		}
		[CompilerGenerated]
		set
		{
			CellSize.m_A = value;
		}
	}

	internal static int ColumnCycleIndex
	{
		[CompilerGenerated]
		get
		{
			return CellSize.m_B;
		}
		[CompilerGenerated]
		set
		{
			CellSize.m_B = value;
		}
	}

	public static void CycleRowHeight()
	{
		checked
		{
			try
			{
				if (RowCycleIndex > KH.A.RowHeightCycle.Count - 1)
				{
					RowCycleIndex = 0;
				}
				A(KH.A.RowHeightCycle[RowCycleIndex]);
				RowCycleIndex++;
				if (RowCycleIndex != 1)
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
					Base.LogActivity(VH.A(148906));
					return;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	public static void CycleColumnWidth()
	{
		checked
		{
			try
			{
				if (ColumnCycleIndex > KH.A.ColumnWidthCycle.Count - 1)
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
					ColumnCycleIndex = 0;
				}
				B(KH.A.ColumnWidthCycle[ColumnCycleIndex]);
				ColumnCycleIndex++;
				if (ColumnCycleIndex == 1)
				{
					Base.LogActivity(VH.A(148939));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	public static void DoRowHeight(IRibbonControl control)
	{
		A(Conversions.ToSingle(control.Tag));
	}

	public static void DoColumnWidth(IRibbonControl control)
	{
		B(Conversions.ToSingle(control.Tag));
	}

	private static void A(float A)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		try
		{
			if (application.Selection is Range)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Range range = JH.A((Range)null);
					if (!Base.IsWorksheetProtected(range.Worksheet))
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
						application.ScreenUpdating = false;
						try
						{
							bool num = JH.A(range);
							range.RowHeight = A;
							if (num)
							{
								JH.A(range, VH.A(148976));
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception b = ex;
							CellSize.A(application, b);
							ProjectData.ClearProjectError();
						}
						application.ScreenUpdating = true;
					}
					range = null;
					break;
				}
			}
		}
		catch (Exception ex2)
		{
			ProjectData.SetProjectError(ex2);
			Exception ex3 = ex2;
			Base.LogException(ex3);
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	private static void B(float A)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		try
		{
			if (application.Selection is Range)
			{
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
					Range range = JH.A((Range)null);
					if (!Base.IsWorksheetProtected(range.Worksheet))
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
						application.ScreenUpdating = false;
						try
						{
							bool num = JH.A(range);
							range.ColumnWidth = A;
							if (num)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									JH.A(range, VH.A(148997));
									break;
								}
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception b = ex;
							CellSize.A(application, b);
							ProjectData.ClearProjectError();
						}
						application.ScreenUpdating = true;
					}
					range = null;
					break;
				}
			}
		}
		catch (Exception ex2)
		{
			ProjectData.SetProjectError(ex2);
			Exception ex3 = ex2;
			Base.LogException(ex3);
			ProjectData.ClearProjectError();
		}
		application = null;
	}

	private static void A(Microsoft.Office.Interop.Excel.Application A, Exception B)
	{
		if (A.ActiveWorkbook.DisplayDrawingObjects == XlDisplayDrawingObjects.xlHide)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Forms.ErrorMessage(VH.A(149022));
					return;
				}
			}
		}
		Base.LogException(B);
	}

	public static void ConformSize(int i)
	{
		//IL_0072: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_003a: Unknown result type (might be due to invalid IL or missing references)
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		if (!B())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		StandardSize standardSize;
		if (A())
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
			try
			{
				standardSize = clsPublish.GetStandardSize(i);
				A(application.InchesToPoints(standardSize.Width), application.InchesToPoints(standardSize.Height));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		application = null;
		standardSize = default(StandardSize);
	}

	public static void ConformPowerPoint()
	{
		if (!B())
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
			Microsoft.Office.Interop.PowerPoint.Shape shape = null;
			if (!A())
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
				try
				{
					Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = ((Microsoft.Office.Interop.PowerPoint.Application)Interaction.GetObject(null, VH.A(62824))).ActiveWindow.Selection.ShapeRange;
					if (shapeRange.Count == 1)
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
						shape = shapeRange[1];
					}
					shapeRange = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				if (shape == null)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							Forms.WarningMessage(VH.A(62869));
							return;
						}
					}
				}
				try
				{
					A(shape.Width, shape.Height);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				shape = null;
				return;
			}
		}
	}

	public static void ConformWord()
	{
		if (!B())
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
				try
				{
					float[] array = clsPublish.SelectedWordShapeSize((Microsoft.Office.Interop.Word.Application)Interaction.GetObject(null, VH.A(62984)));
					if (array[0] > 0f)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								try
								{
									A(array[0], array[1]);
									return;
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
									return;
								}
							}
						}
					}
					throw new Exception();
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					Forms.WarningMessage(VH.A(63017));
					ProjectData.ClearProjectError();
					return;
				}
			}
		}
	}

	private static bool A()
	{
		bool result = true;
		if (!(MH.A.Application.Selection is Range))
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
			Forms.WarningMessage(VH.A(149312));
			result = false;
		}
		return result;
	}

	private static void A(double A, double B)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Range range;
		try
		{
			wpfConformSize wpfConformSize2 = new wpfConformSize();
			switch ((ConformSizeBehavior)K.Settings.ConformSizeBehavior)
			{
			case ConformSizeBehavior.Columns:
				wpfConformSize2.optColumns.IsChecked = true;
				break;
			case ConformSizeBehavior.Rows:
				wpfConformSize2.optRows.IsChecked = true;
				break;
			case ConformSizeBehavior.Both:
				wpfConformSize2.optBoth.IsChecked = true;
				break;
			}
			wpfConformSize2.chkGuide.IsChecked = K.Settings.ConformSizeShowGuide;
			wpfConformSize2.ShowDialog();
			bool? obj;
			bool? isChecked2;
			bool? flag;
			bool? isChecked;
			if (wpfConformSize2.DialogResult.HasValue)
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
				if (wpfConformSize2.DialogResult.Value)
				{
					flag = (isChecked = wpfConformSize2.optColumns.IsChecked);
					if (flag.HasValue)
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
						if (isChecked == true)
						{
							obj = true;
							goto IL_017b;
						}
					}
					flag = (isChecked2 = wpfConformSize2.optBoth.IsChecked);
					if (!flag.HasValue)
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
						obj = null;
					}
					else if (isChecked2 != true)
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
						obj = isChecked;
					}
					else
					{
						obj = true;
					}
					goto IL_017b;
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
			throw new Exception();
			IL_0210:
			bool? obj2;
			isChecked = (bool?)obj2;
			bool value = isChecked.Value;
			bool value2 = wpfConformSize2.chkGuide.IsChecked.Value;
			if (wpfConformSize2.optColumns.IsChecked == true)
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
				K.Settings.ConformSizeBehavior = 0;
			}
			else if (wpfConformSize2.optRows.IsChecked == true)
			{
				K.Settings.ConformSizeBehavior = 1;
			}
			else if (wpfConformSize2.optBoth.IsChecked == true)
			{
				K.Settings.ConformSizeBehavior = 2;
			}
			K.Settings.ConformSizeShowGuide = value2;
			wpfConformSize2 = null;
			application.ScreenUpdating = false;
			application.EnableCancelKey = XlEnableCancelKey.xlErrorHandler;
			range = (Range)application.Selection;
			bool flag2 = JH.A(range);
			Range range2 = range;
			int value3;
			if (value3 != 0)
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
				double num = Conversions.ToDouble(range2.Width);
				Range columns = range2.Columns;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = columns.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range3 = (Range)enumerator.Current;
						double num2 = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(A, range3.Width), num));
						while (Operators.ConditionalCompareObjectGreater(Operators.SubtractObject(Operators.SubtractObject(range3.get_Offset((object)0, (object)1).Left, range3.Left), 0.1), num2, TextCompare: false))
						{
							range3.ColumnWidth = Operators.SubtractObject(range3.ColumnWidth, 0.1);
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
						while (Operators.ConditionalCompareObjectLess(Operators.AddObject(Operators.SubtractObject(range3.get_Offset((object)0, (object)1).Left, range3.Left), 0.1), num2, TextCompare: false))
						{
							range3.ColumnWidth = Operators.AddObject(range3.ColumnWidth, 0.1);
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_045a;
							}
							continue;
							end_IL_045a:
							break;
						}
						range3 = null;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_0473;
						}
						continue;
						end_IL_0473:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			if (value)
			{
				double num = Conversions.ToDouble(range2.Height);
				Range columns = range2.Rows;
				foreach (Range item in columns)
				{
					item.RowHeight = Operators.DivideObject(Operators.MultiplyObject(B, item.Height), num);
					Range range4 = null;
				}
			}
			if (flag2)
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
				JH.A(range, VH.A(149395));
			}
			if (value2)
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
				Guides.Remove();
				Guides.Create((Range)range2.Cells[1, 1], (float)A, (float)B);
			}
			range2 = null;
			goto end_IL_000f;
			IL_017b:
			isChecked2 = obj;
			value3 = (isChecked2.Value ? 1 : 0);
			flag = (isChecked2 = wpfConformSize2.optRows.IsChecked);
			if (flag.HasValue)
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
				if (isChecked2 == true)
				{
					obj2 = true;
					goto IL_0210;
				}
			}
			flag = (isChecked = wpfConformSize2.optBoth.IsChecked);
			if (!flag.HasValue)
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
				obj2 = null;
			}
			else if (isChecked != true)
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
				obj2 = isChecked2;
			}
			else
			{
				obj2 = true;
			}
			goto IL_0210;
			end_IL_000f:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application.EnableCancelKey = XlEnableCancelKey.xlInterrupt;
		range = null;
		application = null;
	}

	private static bool B()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}
}
