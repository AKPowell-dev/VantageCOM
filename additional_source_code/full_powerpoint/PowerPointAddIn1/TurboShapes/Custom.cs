using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TurboShapes;

public sealed class Custom
{
	public static void Convert()
	{
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Application application = NG.A.Application;
			Selection selection;
			Microsoft.Office.Interop.PowerPoint.Shape shape;
			try
			{
				selection = application.ActiveWindow.Selection;
				if (selection.Type == PpSelectionType.ppSelectionShapes)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						if (selection.ShapeRange.Count == 1)
						{
							shape = selection.ShapeRange[1];
							if (shape.Type == MsoShapeType.msoGroup)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									new List<string>();
									try
									{
										bool flag = true;
										try
										{
											enumerator = shape.GroupItems.GetEnumerator();
											while (true)
											{
												if (enumerator.MoveNext())
												{
													if (Operators.CompareString(A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current), string.Empty, TextCompare: false) != 0)
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
														flag = false;
														break;
													}
													break;
												}
												while (true)
												{
													switch (3)
													{
													case 0:
														break;
													default:
														goto end_IL_00e8;
													}
													continue;
													end_IL_00e8:
													break;
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
										if (flag)
										{
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												shape.LockAspectRatio = MsoTriState.msoTrue;
												shape.Tags.Add(Base.TAG_TYPE, 0.ToString());
												Forms.InfoMessage(AH.A(158694));
												ApplyState(shape, 1);
												Edit(shape, 1);
												Base.LogActivity(AH.A(158795));
												break;
											}
										}
										else
										{
											Forms.WarningMessage(AH.A(158830));
										}
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										Forms.ErrorMessage(ex2.Message);
										Base.LogException(ex2);
										ProjectData.ClearProjectError();
									}
									break;
								}
							}
							else
							{
								A();
							}
						}
						else
						{
							A();
						}
						break;
					}
				}
				else
				{
					A();
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				A();
				ProjectData.ClearProjectError();
			}
			application = null;
			selection = null;
			shape = null;
			return;
		}
	}

	private static string A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		string result = string.Empty;
		try
		{
			result = Regex.Match(A.Name, AH.A(159014)).Groups[1].ToString();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void ApplyState(Microsoft.Office.Interop.PowerPoint.Shape shpGroup, int state)
	{
		shpGroup.Tags.Add(Base.TAG_VALUE, state.ToString());
		shpGroup.Select();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = shpGroup.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				string text = A(shape);
				bool flag;
				if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
				{
					flag = false;
					if (!text.Contains(AH.A(12717)))
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (!text.Contains(AH.A(14417)))
						{
							if (Conversions.ToInteger(text) != state)
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
								if (Conversions.ToInteger(text) != 0)
								{
									goto IL_0166;
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
							flag = true;
							goto IL_0166;
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
					string value = ((!text.Contains(AH.A(12717))) ? AH.A(14417) : AH.A(12717));
					string[] array = text.Split(Conversions.ToChar(value));
					int num = 0;
					while (true)
					{
						if (num < array.Length)
						{
							string value2 = array[num];
							if (Conversions.ToInteger(value2) != state)
							{
								if (Conversions.ToInteger(value2) != 0)
								{
									num = checked(num + 1);
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
							}
							flag = true;
							break;
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
						break;
					}
					goto IL_0166;
				}
				B();
				break;
				IL_0166:
				if (flag)
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
					shape.Visible = MsoTriState.msoTrue;
				}
				else
				{
					shape.Visible = MsoTriState.msoFalse;
				}
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
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpGroup, int state)
	{
		int num = 0;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = shpGroup.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				string text = A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
				if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
					if (!text.Contains(AH.A(12717)))
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
						if (!text.Contains(AH.A(14417)))
						{
							if (Conversions.ToInteger(text) <= num)
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
							num = Conversions.ToInteger(text);
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
					}
					string value = ((!text.Contains(AH.A(12717))) ? AH.A(14417) : AH.A(12717));
					string[] array = text.Split(Conversions.ToChar(value));
					int num2 = 0;
					while (true)
					{
						if (num2 < array.Length)
						{
							string value2 = array[num2];
							if (Conversions.ToInteger(value2) <= num)
							{
								break;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_00f2;
								}
								continue;
								end_IL_00f2:
								break;
							}
							num = Conversions.ToInteger(value2);
							num2 = checked(num2 + 1);
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
						break;
					}
					continue;
				}
				B();
				return;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpGroup, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		wpfCustom wpfCustom2 = new wpfCustom();
		wpfCustom2.EditedShape = shpGroup;
		wpfCustom2.MaxState = num;
		wpfCustom2.CurrentState = state;
		wpfCustom2.Top = unitY - wpfCustom2.Height;
		wpfCustom2.Left = unitX;
		wpfCustom2.ShowActivated = false;
		wpfCustom2.Show();
		wpfCustom2 = null;
	}

	private static void A()
	{
		Forms.WarningMessage(AH.A(159075));
	}

	private static void B()
	{
		Forms.ErrorMessage(AH.A(159184));
	}
}
