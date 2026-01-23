using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using A;
using MacabacusMacros.Proofing.UI;
using MacabacusMacros.Proofing.UI.Reformat;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public class BaseItem : BaseItem
{
	[CompilerGenerated]
	private List<NavigationItem> m_A;

	public List<NavigationItem> Objects
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public BaseItem(int intTotal, int intIndex, List<IndexedObject> IndexedObjects, DataTemplate template, DataTemplate navItemTemplate, string strHeader)
		: base(IndexedObjects.Count, intTotal, intIndex, template, navItemTemplate, strHeader)
	{
		List<NavigationItem> list = new List<NavigationItem>();
		using (List<IndexedObject>.Enumerator enumerator = IndexedObjects.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				IndexedObject current = enumerator.Current;
				try
				{
					list.Add(new NavigationItem(this, current, navItemTemplate, strHeader));
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
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
				break;
			}
		}
		Objects = list;
		list = null;
	}

	public Tuple<List<NavigationItem>, int> GenerateChildren(List<IndexedObject> objects, string strHeader)
	{
		List<NavigationItem> list = new List<NavigationItem>();
		int num = 0;
		int num2 = 0;
		checked
		{
			NavigationItem item;
			using (List<IndexedObject>.Enumerator enumerator = objects.GetEnumerator())
			{
				IEnumerator enumerator2 = default(IEnumerator);
				IEnumerator enumerator3 = default(IEnumerator);
				IEnumerator enumerator4 = default(IEnumerator);
				IEnumerator enumerator5 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					IndexedObject current = enumerator.Current;
					int num3 = 0;
					IndexedObject indexedObject = current;
					try
					{
						if (indexedObject.Child is TextRange2)
						{
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
								num3 = A(((TextRange2)indexedObject.Child).Text);
								list.Add(new NavigationItem(this, current, ((BaseItem)this).NavItemDataTemplate, strHeader));
								break;
							}
						}
						else if (indexedObject.Child is BulletFormat2)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								num3 = 0;
								item = new NavigationItem(this, current, ((BaseItem)this).NavItemDataTemplate, strHeader, current.Shape.Name + AH.A(49166));
								list.Add(item);
								break;
							}
						}
						else if (indexedObject.Child is ChartTitle)
						{
							string text = ((ChartTitle)indexedObject.Child).Text;
							num3 = A(text);
							list.Add(A(current, strHeader, text, Icons.CHART));
						}
						else if (indexedObject.Child is AxisTitle)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								string text = ((AxisTitle)indexedObject.Child).Text;
								num3 = A(text);
								list.Add(A(current, strHeader, text, Icons.CHART));
								break;
							}
						}
						else if (indexedObject.Child is TickLabels)
						{
							try
							{
								if (((TickLabels)indexedObject.Child).Parent is Axis axis)
								{
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										Axis axis2 = axis;
										if (axis2.Type != Microsoft.Office.Interop.PowerPoint.XlAxisType.xlValue)
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
											if (!A(axis))
											{
												if (axis2.Type != Microsoft.Office.Interop.PowerPoint.XlAxisType.xlCategory)
												{
													if (axis2.Type != Microsoft.Office.Interop.PowerPoint.XlAxisType.xlSeriesAxis)
													{
														goto IL_0359;
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
												try
												{
													enumerator2 = ((IEnumerable)axis2.CategoryNames).GetEnumerator();
													while (enumerator2.MoveNext())
													{
														string text2 = Conversions.ToString(enumerator2.Current);
														num3 += A(text2);
														list.Add(A(current, strHeader, text2, Icons.CHART));
													}
													while (true)
													{
														switch (3)
														{
														case 0:
															break;
														default:
															goto end_IL_032d;
														}
														continue;
														end_IL_032d:
														break;
													}
												}
												finally
												{
													if (enumerator2 is IDisposable)
													{
														while (true)
														{
															switch (3)
															{
															case 0:
																continue;
															}
															(enumerator2 as IDisposable).Dispose();
															break;
														}
													}
												}
												goto IL_0359;
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
										num3 = (int)Math.Round((axis2.MaximumScale - axis2.MinimumScale) / axis2.MajorUnit + 1.0);
										double maximumScale = axis2.MaximumScale;
										double minimumScale = axis2.MinimumScale;
										double num4 = -1.0 * axis2.MajorUnit;
										bool flag = num4 >= 0.0;
										double num5 = maximumScale;
										while (true)
										{
											bool num6;
											if (!flag)
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
												num6 = num5 >= minimumScale;
											}
											else
											{
												num6 = num5 <= minimumScale;
											}
											if (num6)
											{
												list.Add(A(current, strHeader, num5.ToString(), Icons.CHART));
												num5 += num4;
												continue;
											}
											break;
										}
										goto IL_0359;
										IL_0359:
										axis2 = null;
										Axis axis3 = null;
										break;
									}
								}
								else if (((TickLabels)indexedObject.Child).Parent is ChartGroup)
								{
									list.Add(A(current, strHeader, indexedObject.AreRadarLabels ? AH.A(49202) : AH.A(49185), Icons.CHART));
								}
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
						}
						else if (indexedObject.Child is IMsoDataLabel)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								string text = ((IMsoDataLabel)indexedObject.Child).Text;
								num3 = A(text);
								list.Add(A(current, strHeader, text, Icons.CHART));
								break;
							}
						}
						else if (indexedObject.Child is IMsoDataLabels)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								try
								{
									enumerator3 = ((IMsoDataLabels)indexedObject.Child).GetEnumerator();
									while (enumerator3.MoveNext())
									{
										IMsoDataLabel msoDataLabel = (IMsoDataLabel)enumerator3.Current;
										try
										{
											string text = msoDataLabel.Text;
											num3 += A(text);
											item = new NavigationItem(this, new IndexedObject(RuntimeHelpers.GetObjectValue(current.SlideOrLayout), current.Shape, msoDataLabel), ((BaseItem)this).NavItemDataTemplate, strHeader, text);
											item.IconPath = Icons.CHART;
											list.Add(item);
										}
										catch (Exception ex3)
										{
											ProjectData.SetProjectError(ex3);
											Exception ex4 = ex3;
											ProjectData.ClearProjectError();
										}
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_04e7;
										}
										continue;
										end_IL_04e7:
										break;
									}
								}
								finally
								{
									if (enumerator3 is IDisposable)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											(enumerator3 as IDisposable).Dispose();
											break;
										}
									}
								}
								break;
							}
						}
						else if (indexedObject.Child is Legend)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								Chart chart = (Chart)((Legend)indexedObject.Child).Parent;
								try
								{
									enumerator4 = ((IEnumerable)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
									while (enumerator4.MoveNext())
									{
										string text = ((IMsoSeries)enumerator4.Current).Name;
										num3 += A(text);
										list.Add(A(current, strHeader, text, Icons.CHART));
									}
								}
								finally
								{
									if (enumerator4 is IDisposable)
									{
										while (true)
										{
											switch (7)
											{
											case 0:
												continue;
											}
											(enumerator4 as IDisposable).Dispose();
											break;
										}
									}
								}
								chart = null;
								break;
							}
						}
						else if (indexedObject.Child is DataTable)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								Chart chart2 = (Chart)((DataTable)indexedObject.Child).Parent;
								try
								{
									enumerator5 = ((IEnumerable)chart2.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
									while (enumerator5.MoveNext())
									{
										IMsoSeries msoSeries = (IMsoSeries)enumerator5.Current;
										if (!Charts.ImplsPoints(msoSeries))
										{
											continue;
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											break;
										}
										string text = msoSeries.Name;
										num3 = Conversions.ToInteger(Operators.AddObject(num3, Operators.AddObject(NewLateBinding.LateGet(msoSeries.Points(RuntimeHelpers.GetObjectValue(Missing.Value)), null, AH.A(13955), new object[0], null, null, null), A(text))));
										list.Add(A(current, strHeader, text, Icons.CHART));
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_06e7;
										}
										continue;
										end_IL_06e7:
										break;
									}
								}
								finally
								{
									if (enumerator5 is IDisposable)
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											(enumerator5 as IDisposable).Dispose();
											break;
										}
									}
								}
								chart2 = null;
								break;
							}
						}
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
					indexedObject = null;
					if (num3 <= 0)
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
					num2 += num3;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_074a;
					}
					continue;
					end_IL_074a:
					break;
				}
			}
			num = num2;
			item = null;
			return new Tuple<List<NavigationItem>, int>(list, num);
		}
	}

	private bool A(Axis A)
	{
		bool result;
		try
		{
			if (A.CategoryNames != null)
			{
				result = false;
				goto IL_0074;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		try
		{
			double majorUnit = A.MajorUnit;
			result = !(majorUnit <= 0.0) && !((A.MaximumScale - A.MinimumScale) / majorUnit + 1.0 <= 0.0);
		}
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			result = false;
			ProjectData.ClearProjectError();
		}
		goto IL_0074;
		IL_0074:
		return result;
	}

	private NavigationItem A(IndexedObject A, string B, string C, string D)
	{
		return new NavigationItem(this, A, ((BaseItem)this).NavItemDataTemplate, B, C);
	}

	private int A(string A)
	{
		return Regex.Matches(A, AH.A(49231)).Count;
	}
}
