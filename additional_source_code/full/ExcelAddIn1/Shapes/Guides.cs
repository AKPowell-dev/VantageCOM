using System;
using System.Collections;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using ExcelAddIn1.Workbook;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Shapes;

public sealed class Guides
{
	private static readonly string m_A = VH.A(102058);

	public static void ShowExcel()
	{
		if (!A())
		{
			return;
		}
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (A(application))
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
				try
				{
					enumerator = ((Range)application.Selection).Areas.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Range range = (Range)enumerator.Current;
							Create((Range)range.Cells[1, 1], Conversions.ToSingle(range.Width), Conversions.ToSingle(range.Height));
							range = null;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_00b9;
							}
							continue;
							end_IL_00b9:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Forms.WarningMessage(VH.A(101703));
					ProjectData.ClearProjectError();
				}
			}
			application = null;
			return;
		}
	}

	public static void ShowPowerPoint()
	{
		if (!A())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			if (A(application))
			{
				try
				{
					Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = ((Microsoft.Office.Interop.PowerPoint.Application)Interaction.GetObject(null, VH.A(62824))).ActiveWindow.Selection.ShapeRange;
					if (shapeRange.Count != 1)
					{
						throw new Exception();
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						Create((Range)((Range)application.Selection).Cells[1, 1], shapeRange[1].Width, shapeRange[1].Height);
						shapeRange = null;
						break;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Forms.WarningMessage(VH.A(62869));
					ProjectData.ClearProjectError();
				}
			}
			application = null;
			return;
		}
	}

	public static void ShowWord()
	{
		if (!A())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (A(application))
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
				float[] array = clsPublish.SelectedWordShapeSize((Microsoft.Office.Interop.Word.Application)Interaction.GetObject(null, VH.A(62984)));
				if (!(array[0] > 0f))
				{
					throw new Exception();
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					Create((Range)((Range)application.Selection).Cells[1, 1], array[0], array[1]);
					break;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.WarningMessage(VH.A(63017));
				ProjectData.ClearProjectError();
			}
		}
		application = null;
	}

	public static void Show(int i)
	{
		//IL_00bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_0089: Unknown result type (might be due to invalid IL or missing references)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			StandardSize standardSize;
			if (A(application))
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
				try
				{
					standardSize = clsPublish.GetStandardSize(i);
					Create((Range)((Range)application.Selection).Cells[1, 1], (float)application.InchesToPoints(standardSize.Width), (float)application.InchesToPoints(standardSize.Height));
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					Forms.ErrorMessage(VH.A(101756));
					ProjectData.ClearProjectError();
				}
			}
			standardSize = default(StandardSize);
			application = null;
			return;
		}
	}

	public static Microsoft.Office.Interop.Excel.Shape Create(Range rng, float Width, float Height)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Microsoft.Office.Interop.Excel.Shape shape = default(Microsoft.Office.Interop.Excel.Shape);
		Microsoft.Office.Interop.Excel.Shape shape2 = default(Microsoft.Office.Interop.Excel.Shape);
		Color gUIDE_COLOR = default(Color);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		Microsoft.Office.Interop.Excel.Shape result = default(Microsoft.Office.Interop.Excel.Shape);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				object instance;
				string memberName;
				object[] obj;
				object[] array;
				bool[] obj2;
				bool[] array2;
				object obj3;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 582:
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
							goto IL_0010;
						case 4:
							goto IL_0022;
						case 5:
							goto IL_004c;
						case 6:
							goto IL_0056;
						case 7:
							goto IL_014a;
						case 8:
							goto IL_0150;
						case 9:
							goto IL_0161;
						case 10:
							goto IL_0197;
						case 11:
							goto IL_01ad;
						case 12:
							goto IL_01b8;
						case 13:
							goto IL_01c3;
						case 14:
							goto IL_01d2;
						case 15:
							goto IL_01d5;
						case 16:
							goto IL_01e0;
						case 17:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 18:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_014a:
					num2 = 7;
					shape = shape2;
					goto IL_0150;
					IL_0007:
					num2 = 2;
					gUIDE_COLOR = clsGuides.GUIDE_COLOR;
					goto IL_0010;
					IL_0010:
					num2 = 3;
					application = MH.A.Application;
					goto IL_0022;
					IL_0022:
					num2 = 4;
					if (!Workbooks.IsShared(application.ActiveWorkbook, true, (System.Windows.Window)null))
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
						goto IL_004c;
					}
					goto IL_01e0;
					IL_0150:
					num2 = 8;
					shape.Fill.Visible = MsoTriState.msoFalse;
					goto IL_0161;
					IL_0161:
					num2 = 9;
					shape.Line.ForeColor.RGB = Information.RGB(gUIDE_COLOR.R, gUIDE_COLOR.G, gUIDE_COLOR.B);
					goto IL_0197;
					IL_0197:
					num2 = 10;
					shape.Line.Weight = 1f;
					goto IL_01ad;
					IL_004c:
					num2 = 5;
					application.ScreenUpdating = false;
					goto IL_0056;
					IL_0056:
					num2 = 6;
					instance = NewLateBinding.LateGet(application.ActiveSheet, null, VH.A(101831), new object[0], null, null, null);
					memberName = VH.A(101844);
					obj = new object[5]
					{
						MsoAutoShapeType.msoShapeRectangle,
						rng.Left,
						rng.Top,
						Width,
						Height
					};
					array = obj;
					obj2 = new bool[5] { false, false, false, true, true };
					array2 = obj2;
					obj3 = NewLateBinding.LateGet(instance, null, memberName, obj, null, null, obj2);
					if (array2[3])
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
						Width = (float)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[3]), typeof(float));
					}
					if (array2[4])
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
						Height = (float)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[4]), typeof(float));
					}
					shape2 = (Microsoft.Office.Interop.Excel.Shape)obj3;
					goto IL_014a;
					IL_01b8:
					num2 = 12;
					shape.Placement = XlPlacement.xlMove;
					goto IL_01c3;
					IL_01c3:
					num2 = 13;
					shape.Name = Guides.m_A;
					goto IL_01d2;
					IL_01ad:
					num2 = 11;
					shape.ZOrder(MsoZOrderCmd.msoBringToFront);
					goto IL_01b8;
					IL_01d5:
					num2 = 15;
					application.ScreenUpdating = true;
					goto IL_01e0;
					IL_01e0:
					application = null;
					break;
					IL_01d2:
					shape = null;
					goto IL_01d5;
					end_IL_0000_2:
					break;
				}
				num2 = 17;
				result = shape2;
				break;
				end_IL_0000:;
			}
			catch (object obj4) when (obj4 is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj4);
				try0000_dispatch = 582;
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
				switch (4)
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

	public static void Remove()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (application.ActiveSheet is Worksheet)
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
			Worksheet worksheet = (Worksheet)application.ActiveSheet;
			application.ScreenUpdating = false;
			try
			{
				for (int i = worksheet.Shapes.Count; i >= 1; i = checked(i + -1))
				{
					Microsoft.Office.Interop.Excel.Shape shape = worksheet.Shapes.Item(i);
					if (Operators.CompareString(shape.Name, Guides.m_A, TextCompare: false) == 0)
					{
						shape.Delete();
					}
					shape = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_0096;
					}
					continue;
					end_IL_0096:
					break;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			worksheet = null;
		}
		application = null;
	}

	private static bool A(Microsoft.Office.Interop.Excel.Application A)
	{
		bool flag = true;
		if (Operators.ConditionalCompareObjectNotEqual(NewLateBinding.LateGet(A.ActiveSheet, null, VH.A(101861), new object[0], null, null, null), XlSheetType.xlWorksheet, TextCompare: false))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Forms.WarningMessage(VH.A(101870));
			flag = false;
		}
		if (!(A.Selection is Range))
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
			Forms.WarningMessage(VH.A(101973));
			flag = false;
		}
		if (flag)
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
			flag = Miscellaneous.DisplayObjects(A.ActiveWorkbook);
		}
		return flag;
	}

	private static bool A()
	{
		return Access.AllowExcelOperation((PlanType)4, (Restriction)1, false);
	}
}
