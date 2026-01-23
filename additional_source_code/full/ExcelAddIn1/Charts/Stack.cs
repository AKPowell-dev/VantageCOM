using System;
using System.Collections;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

public sealed class Stack
{
	public static void Initiate()
	{
		if (!Helpers.A())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		int result = default(int);
		float num2 = default(float);
		float num3 = default(float);
		float num4 = default(float);
		float num5 = default(float);
		float num6 = default(float);
		XlPlacement xlPlacement = default(XlPlacement);
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
			Application application = MH.A.Application;
			int num = 0;
			ShapeRange shapeRange;
			try
			{
				if (Operators.CompareString(Versioned.TypeName(RuntimeHelpers.GetObjectValue(application.Selection)), VH.A(56245), TextCompare: false) != 0)
				{
					A();
					goto IL_0259;
				}
				shapeRange = (ShapeRange)NewLateBinding.LateGet(application.Selection, null, VH.A(56274), new object[0], null, null, null);
				if (shapeRange.Count <= 1)
				{
					A();
					goto IL_0259;
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
				application.ScreenUpdating = false;
				enumerator = shapeRange.GetEnumerator();
				try
				{
					while (true)
					{
						if (!enumerator.MoveNext())
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_0220;
								}
								continue;
								end_IL_0220:
								break;
							}
							break;
						}
						Shape shape = (Shape)enumerator.Current;
						if (shape.HasChart != MsoTriState.msoTrue)
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
						ChartObject chartObject = (ChartObject)shape.Chart.Parent;
						if (num != 0)
						{
							if (num % result == 0)
							{
								num2 += num3;
								num4 = num5;
							}
							else
							{
								num4 += num6;
							}
							ChartObject chartObject2 = chartObject;
							chartObject2.Width = num6;
							chartObject2.Height = num3;
							chartObject2.Left = num4;
							chartObject2.Top = num2;
							chartObject2.Placement = xlPlacement;
							_ = null;
							goto IL_020c;
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
						string text = Forms.InputBox2(VH.A(63767), VH.A(63792), Conversions.ToString(3));
						if (modFunctionsStr.IsBlank(text))
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_0133;
								}
								continue;
								end_IL_0133:
								break;
							}
						}
						else
						{
							if (int.TryParse(text, out result))
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
								if (result > 0)
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
									ChartObject chartObject3 = chartObject;
									num6 = (float)chartObject3.Width;
									num3 = (float)chartObject3.Height;
									num4 = (float)chartObject3.Left;
									num2 = (float)chartObject3.Top;
									xlPlacement = (XlPlacement)Conversions.ToInteger(chartObject3.Placement);
									_ = null;
									num5 = num4;
									goto IL_020c;
								}
							}
							Forms.WarningMessage(VH.A(63823));
						}
						goto end_IL_00a5;
						IL_020c:
						chartObject = null;
						num = checked(num + 1);
					}
					goto IL_0242;
					end_IL_00a5:;
				}
				finally
				{
					IDisposable disposable = enumerator as IDisposable;
					if (disposable != null)
					{
						disposable.Dispose();
					}
				}
				goto end_IL_002e;
				IL_0259:
				clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)5, VH.A(63767));
				goto end_IL_002e;
				IL_0242:
				if (num < 2)
				{
					A();
				}
				goto IL_0259;
				end_IL_002e:;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			application.ScreenUpdating = true;
			application = null;
			shapeRange = null;
			return;
		}
	}

	private static void A()
	{
		Forms.WarningMessage(VH.A(63884));
	}
}
