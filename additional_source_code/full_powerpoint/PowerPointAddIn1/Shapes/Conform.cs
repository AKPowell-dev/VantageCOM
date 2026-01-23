using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Conform
{
	public static void Width()
	{
		A(A: true, B: false);
		Base.LogActivity(AH.A(81349));
	}

	public static void Height()
	{
		A(A: false, B: true);
		Base.LogActivity(AH.A(81376));
	}

	public static void Size()
	{
		A(A: true, B: true);
		Base.LogActivity(AH.A(81405));
	}

	private static void A(bool A, bool B)
	{
		Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
		try
		{
			shapeRange = Base.SelectedShapes();
			NG.A.Application.StartNewUndoEntry();
			if (shapeRange.Count > 1)
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
					if (A)
					{
						shapeRange.Width = shapeRange[1].Width;
					}
					if (!B)
					{
						break;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						shapeRange.Height = shapeRange[1].Height;
						break;
					}
					break;
				}
			}
			else if (shapeRange.Count == 1)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					PageSetup pageSetup = NG.A.Application.ActivePresentation.PageSetup;
					if (A)
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
						shapeRange.Width = pageSetup.SlideWidth;
					}
					if (B)
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
						shapeRange.Height = pageSetup.SlideHeight;
					}
					pageSetup = null;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Base.AlignError();
			ProjectData.ClearProjectError();
		}
		shapeRange = null;
	}

	public static void Adjustments()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			List<MsoAutoShapeType> list = new List<MsoAutoShapeType>();
			Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
			try
			{
				shapeRange = Base.SelectedShapes();
				if (shapeRange.Count > 1)
				{
					{
						enumerator = shapeRange.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
								if (shape.Type != MsoShapeType.msoAutoShape)
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
								list.Add(shape.AutoShapeType);
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_0085;
								}
								continue;
								end_IL_0085:
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
					if (list.Count == shapeRange.Count && list.Distinct().ToList().Count == 1)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							try
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape2 = shapeRange[1];
								int count = shape2.Adjustments.Count;
								NG.A.Application.StartNewUndoEntry();
								try
								{
									enumerator2 = shapeRange.GetEnumerator();
									while (enumerator2.MoveNext())
									{
										Microsoft.Office.Interop.PowerPoint.Shape shape3 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
										if (shape3 == shape2)
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
										int num = count;
										for (int i = 1; i <= num; i = checked(i + 1))
										{
											shape3.Adjustments[i] = shape2.Adjustments[i];
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
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_0181;
										}
										continue;
										end_IL_0181:
										break;
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
								Base.LogActivity(AH.A(81430));
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								B();
								ProjectData.ClearProjectError();
							}
							finally
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape2 = null;
							}
							break;
						}
					}
					else
					{
						B();
					}
				}
				else
				{
					B();
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				B();
				ProjectData.ClearProjectError();
			}
			shapeRange = null;
			list = null;
			return;
		}
	}

	private static void B()
	{
		Forms.WarningMessage(AH.A(81469));
	}

	public static void Points()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
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
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
				try
				{
					shapeRange = Base.SelectedShapes();
					if (shapeRange.Count > 1)
					{
						try
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = shapeRange[1];
							if (shape.Type == MsoShapeType.msoFreeform && shape.Nodes.Count > 0)
							{
								int count = shape.Nodes.Count;
								int num = 0;
								NG.A.Application.StartNewUndoEntry();
								try
								{
									enumerator = shapeRange.GetEnumerator();
									while (enumerator.MoveNext())
									{
										Microsoft.Office.Interop.PowerPoint.Shape shape2;
										if ((shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current) != shape)
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
											if (shape2.Type == MsoShapeType.msoFreeform)
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
												if (shape2.Nodes.Count == count)
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
													int num2 = count;
													for (int i = 1; i <= num2; i++)
													{
														shape2.Nodes.SetPosition(i, Conversions.ToSingle(Operators.AddObject(Operators.SubtractObject(NewLateBinding.LateIndexGet(shape.Nodes[i].Points, new object[2] { 1, 1 }, null), shape.Left), shape2.Left)), Conversions.ToSingle(Operators.AddObject(Operators.SubtractObject(NewLateBinding.LateIndexGet(shape.Nodes[i].Points, new object[2] { 1, 2 }, null), shape.Top), shape2.Top)));
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
													num++;
												}
											}
										}
										shape2 = null;
									}
									while (true)
									{
										switch (4)
										{
										case 0:
											break;
										default:
											goto end_IL_020d;
										}
										continue;
										end_IL_020d:
										break;
									}
								}
								finally
								{
									if (enumerator is IDisposable)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											(enumerator as IDisposable).Dispose();
											break;
										}
									}
								}
								if (num == 0)
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
									Forms.InfoMessage(AH.A(81590));
								}
								Base.LogActivity(AH.A(81629));
							}
							else
							{
								Forms.WarningMessage(AH.A(81658));
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							Forms.ErrorMessage(ex2.Message);
							ProjectData.ClearProjectError();
						}
						finally
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = null;
						}
					}
					else
					{
						C();
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					C();
					ProjectData.ClearProjectError();
				}
				shapeRange = null;
				return;
			}
		}
	}

	private static void C()
	{
		Forms.WarningMessage(AH.A(81757));
	}
}
