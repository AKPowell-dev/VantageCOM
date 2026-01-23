using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class Effects
{
	public static void Animation(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		bool flag;
		try
		{
			shp.PickupAnimation();
			flag = true;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			flag = false;
			ProjectData.ClearProjectError();
		}
		if (!flag)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Main.Analysis.Errors.Add(new Animation(sld, shp));
			return;
		}
	}

	public static void AllEffects(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, Severity sevShapeEffects, Severity sevTextEffects)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Invalid comparison between Unknown and I4
		//IL_0005: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Invalid comparison between Unknown and I4
		bool flag = (int)sevShapeEffects > 0;
		bool flag2 = (int)sevTextEffects > 0;
		checked
		{
			try
			{
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (A(shp))
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
						Main.Analysis.Errors.Add(new ShapeEffects(sld, shp));
					}
				}
				if (shp.HasTable == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
						{
							List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
							List<TextRange2> list2 = new List<TextRange2>();
							Table table = shp.Table;
							int count = table.Rows.Count;
							int count2 = table.Columns.Count;
							int num = count;
							for (int i = 1; i <= num; i++)
							{
								int num2 = count2;
								for (int j = 1; j <= num2; j++)
								{
									Cell cell = table.Cell(i, j);
									if (flag)
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
										if (G(cell.Shape))
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
											list.Add(cell.Shape);
										}
									}
									if (flag2)
									{
										Microsoft.Office.Interop.PowerPoint.Shape shape = cell.Shape;
										if (shape.HasTextFrame == MsoTriState.msoTrue)
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
											if (shape.TextFrame2.HasText == MsoTriState.msoTrue && A(shape.TextFrame2))
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
												list2.Add(shape.TextFrame2.TextRange);
											}
										}
										shape = null;
									}
									cell = null;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_0171;
									}
									continue;
									end_IL_0171:
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
									table = null;
									if (list.Count > 0)
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
										Main.Analysis.Errors.Add(new ShapeEffects(sld, shp, list));
									}
									if (list2.Count > 0)
									{
										A(sld, shp, list2);
									}
									list2 = null;
									list = null;
									return;
								}
							}
						}
						}
					}
				}
				IEnumerator enumerator = default(IEnumerator);
				IEnumerator enumerator2 = default(IEnumerator);
				if (shp.HasSmartArt == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
						{
							List<Microsoft.Office.Core.Shape> list3 = new List<Microsoft.Office.Core.Shape>();
							List<TextRange2> list2 = new List<TextRange2>();
							{
								enumerator = shp.SmartArt.AllNodes.GetEnumerator();
								try
								{
									while (enumerator.MoveNext())
									{
										SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
										try
										{
											enumerator2 = smartArtNode.Shapes.GetEnumerator();
											while (enumerator2.MoveNext())
											{
												Microsoft.Office.Core.Shape shape2 = (Microsoft.Office.Core.Shape)enumerator2.Current;
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
													if (A(shape2))
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
														list3.Add(shape2);
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
													if (shape2.TextFrame2.HasText == MsoTriState.msoTrue && A(shape2.TextFrame2))
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
														list2.Add(shape2.TextFrame2.TextRange);
													}
												}
											}
											while (true)
											{
												switch (6)
												{
												case 0:
													break;
												default:
													goto end_IL_02de;
												}
												continue;
												end_IL_02de:
												break;
											}
										}
										finally
										{
											if (enumerator2 is IDisposable)
											{
												while (true)
												{
													switch (1)
													{
													case 0:
														break;
													default:
														(enumerator2 as IDisposable).Dispose();
														goto end_IL_02f3;
													}
													continue;
													end_IL_02f3:
													break;
												}
											}
										}
										if (flag2 && smartArtNode.TextFrame2.HasText == MsoTriState.msoTrue && A(smartArtNode.TextFrame2))
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
											list2.Add(smartArtNode.TextFrame2.TextRange);
										}
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_035c;
										}
										continue;
										end_IL_035c:
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
							if (list3.Count > 0)
							{
								Main.Analysis.Errors.Add(new ShapeEffects(sld, shp, list3));
							}
							if (list2.Count > 0)
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
								A(sld, shp, list2);
							}
							list3 = null;
							list2 = null;
							return;
						}
						}
					}
				}
				IEnumerator enumerator3 = default(IEnumerator);
				if (shp.HasChart == MsoTriState.msoTrue)
				{
					while (true)
					{
						Chart chart;
						TextRange2 textRange;
						switch (6)
						{
						case 0:
							break;
						default:
							{
								chart = shp.Chart;
								if (A(chart.PlotArea.Format))
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
									Main.Analysis.Errors.Add(new ShapeEffects(sld, shp, chart.PlotArea));
								}
								try
								{
									enumerator3 = ((IEnumerable)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
									while (enumerator3.MoveNext())
									{
										IMsoSeries msoSeries = (IMsoSeries)enumerator3.Current;
										if (A(msoSeries.Format))
										{
											Main.Analysis.Errors.Add(new ShapeEffects(sld, shp, msoSeries));
										}
									}
								}
								finally
								{
									if (enumerator3 is IDisposable)
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												(enumerator3 as IDisposable).Dispose();
												goto end_IL_04a3;
											}
											continue;
											end_IL_04a3:
											break;
										}
									}
								}
								if (chart.HasTitle)
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
									if (A(chart.ChartTitle.Format))
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
										Main.Analysis.Errors.Add(new ShapeEffects(sld, shp, chart.ChartTitle));
									}
									if (A(chart.ChartTitle.Format.TextFrame2))
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
										Main.Analysis.Errors.Add(new TextEffects(sld, shp, chart.ChartTitle));
									}
								}
								if (chart.HasLegend)
								{
									if (A(chart.Legend.Format))
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
										Main.Analysis.Errors.Add(new ShapeEffects(sld, shp, chart.Legend));
									}
									if (A(chart.Legend.Format.TextFrame2))
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
										Main.Analysis.Errors.Add(new TextEffects(sld, shp, chart.Legend));
									}
								}
								if (chart.HasDataTable)
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
									if (A(chart.DataTable.Format))
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
										Main.Analysis.Errors.Add(new ShapeEffects(sld, shp, chart.DataTable));
									}
									textRange = chart.DataTable.Format.TextFrame2.TextRange;
									if (!B(textRange.Font) && !D(textRange.Font))
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
										if (!A(textRange.Font))
										{
											goto IL_06db;
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
									}
									Main.Analysis.Errors.Add(new TextEffects(sld, shp, shp.Chart.DataTable));
									goto IL_06db;
								}
								goto IL_06df;
							}
							IL_06df:
							foreach (Axis item in modCharts.AxesList(shp.Chart))
							{
								try
								{
									Axis axis = item;
									if (A(axis.Format))
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
										Main.Analysis.Errors.Add(new ShapeEffects(sld, shp, item));
									}
									if (axis.HasTitle)
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
										if (A(axis.AxisTitle.Format))
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
											Main.Analysis.Errors.Add(new ShapeEffects(sld, shp, axis.AxisTitle));
										}
										if (A(axis.AxisTitle.Format.TextFrame2))
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
											Main.Analysis.Errors.Add(new TextEffects(sld, shp, axis.AxisTitle));
										}
									}
									axis = null;
								}
								finally
								{
									Axis current = null;
								}
							}
							chart = null;
							return;
							IL_06db:
							textRange = null;
							goto IL_06df;
						}
					}
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
					if (shp.HasTextFrame != MsoTriState.msoTrue)
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
						if (shp.TextFrame2.HasText != MsoTriState.msoTrue)
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
							List<TextRange2> list2 = new List<TextRange2>();
							if (A(shp.TextFrame2))
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
								A(sld, shp, new List<TextRange2>(new TextRange2[1] { shp.TextFrame2.TextRange }));
							}
							list2 = null;
							return;
						}
					}
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

	private static void A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, List<TextRange2> C)
	{
		Main.Analysis.Errors.Add(new TextEffects(A, B, C));
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (!B(A) && !E(A) && !F(A))
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
			if (!D(A))
			{
				return C(A);
			}
		}
		return true;
	}

	private static bool A(Microsoft.Office.Core.Shape A)
	{
		if (!B(A))
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
			if (!E(A))
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
				if (!F(A))
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
					if (!D(A))
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								return C(A);
							}
						}
					}
				}
			}
		}
		return true;
	}

	private static bool A(ChartFormat A)
	{
		if (!B(A) && !E(A) && !F(A) && !D(A))
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
			if (!C(A))
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						return G(A);
					}
				}
			}
		}
		return true;
	}

	private static bool A(IMsoChartFormat A)
	{
		if (!B(A))
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
			if (!E(A) && !F(A) && !D(A))
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
				if (!C(A))
				{
					return G(A);
				}
			}
		}
		return true;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.TextFrame2 A)
	{
		bool result = false;
		if (B(A))
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
			result = true;
		}
		else
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.TextRange.get_Runs(-1, -1).GetEnumerator();
				while (enumerator.MoveNext())
				{
					TextRange2 textRange = (TextRange2)enumerator.Current;
					if (!Effects.A(textRange.Font))
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
						if (!D(textRange.Font))
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
							if (!C(textRange.Font))
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
								if (!B(textRange.Font))
								{
									textRange = null;
									continue;
								}
							}
						}
					}
					result = true;
					break;
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
		}
		return result;
	}

	private static bool A(Microsoft.Office.Core.TextFrame2 A)
	{
		bool result = false;
		if (B(A))
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
			result = true;
		}
		else
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.TextRange.get_Runs(-1, -1).GetEnumerator();
				while (true)
				{
					if (enumerator.MoveNext())
					{
						TextRange2 textRange = (TextRange2)enumerator.Current;
						if (!Effects.A(textRange.Font))
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
							if (!D(textRange.Font))
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
								if (!C(textRange.Font))
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
									if (!B(textRange.Font))
									{
										textRange = null;
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
							}
						}
						result = true;
						break;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_00c8;
						}
						continue;
						end_IL_00c8:
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
		}
		return result;
	}

	private static bool B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Shadow);
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

	private static bool B(Microsoft.Office.Core.Shape A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Shadow);
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

	private static bool B(ChartFormat A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Shadow);
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

	private static bool B(IMsoChartFormat A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Shadow);
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

	private static bool A(Microsoft.Office.Interop.PowerPoint.ShadowFormat A)
	{
		if (A.Visible == MsoTriState.msoTrue)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return A.Transparency < 1f;
				}
			}
		}
		return false;
	}

	private static bool A(Microsoft.Office.Core.ShadowFormat A)
	{
		if (A.Visible == MsoTriState.msoTrue)
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
					return A.Transparency < 1f;
				}
			}
		}
		return false;
	}

	private static bool A(Font2 A)
	{
		bool result;
		try
		{
			result = A.Shadow.Visible == MsoTriState.msoTrue;
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

	private static bool C(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Glow);
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

	private static bool C(Microsoft.Office.Core.Shape A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Glow);
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

	private static bool C(ChartFormat A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Glow);
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

	private static bool C(IMsoChartFormat A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Glow);
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

	private static bool B(Font2 A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Glow);
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

	private static bool A(GlowFormat A)
	{
		return A.Radius > 0f;
	}

	private static bool D(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		bool result;
		try
		{
			result = Effects.A(A.SoftEdge);
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

	private static bool D(Microsoft.Office.Core.Shape A)
	{
		bool result;
		try
		{
			result = Effects.A(A.SoftEdge);
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

	private static bool D(ChartFormat A)
	{
		bool result;
		try
		{
			result = Effects.A(A.SoftEdge);
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

	private static bool D(IMsoChartFormat A)
	{
		bool result;
		try
		{
			result = Effects.A(A.SoftEdge);
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

	private static bool A(SoftEdgeFormat A)
	{
		return A.Type != MsoSoftEdgeType.msoSoftEdgeTypeNone;
	}

	private static bool C(Font2 A)
	{
		bool result;
		try
		{
			result = A.SoftEdgeFormat != MsoSoftEdgeType.msoSoftEdgeTypeNone;
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

	private static bool E(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Reflection);
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

	private static bool E(Microsoft.Office.Core.Shape A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Reflection);
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

	private static bool E(ChartFormat A)
	{
		return false;
	}

	private static bool E(IMsoChartFormat A)
	{
		return false;
	}

	private static bool D(Font2 A)
	{
		bool result;
		try
		{
			result = Effects.A(A.Reflection);
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

	private static bool A(ReflectionFormat A)
	{
		if (A.Type != MsoReflectionType.msoReflectionTypeNone)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return A.Size > 0f;
				}
			}
		}
		return false;
	}

	private static bool F(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		bool result;
		try
		{
			result = A.ThreeD.Visible == MsoTriState.msoTrue;
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

	private static bool F(Microsoft.Office.Core.Shape A)
	{
		bool result;
		try
		{
			result = A.ThreeD.Visible == MsoTriState.msoTrue;
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

	private static bool F(ChartFormat A)
	{
		return false;
	}

	private static bool F(IMsoChartFormat A)
	{
		return false;
	}

	private static bool B(Microsoft.Office.Interop.PowerPoint.TextFrame2 A)
	{
		bool result;
		try
		{
			result = Effects.A(A.ThreeD);
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

	private static bool B(Microsoft.Office.Core.TextFrame2 A)
	{
		bool result;
		try
		{
			result = Effects.A(A.ThreeD);
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

	private static bool G(ChartFormat A)
	{
		bool result;
		try
		{
			result = Effects.A(A.ThreeD);
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

	private static bool G(IMsoChartFormat A)
	{
		bool result;
		try
		{
			result = Effects.A(A.ThreeD);
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

	private static bool G(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		bool result;
		try
		{
			result = Effects.A(A.ThreeD);
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

	private static bool A(Microsoft.Office.Interop.PowerPoint.ThreeDFormat A)
	{
		if (A.BevelTopType == MsoBevelType.msoBevelNone)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return A.BevelBottomType != MsoBevelType.msoBevelNone;
				}
			}
		}
		return true;
	}

	private static bool A(Microsoft.Office.Core.ThreeDFormat A)
	{
		if (A.BevelTopType == MsoBevelType.msoBevelNone)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return A.BevelBottomType != MsoBevelType.msoBevelNone;
				}
			}
		}
		return true;
	}
}
