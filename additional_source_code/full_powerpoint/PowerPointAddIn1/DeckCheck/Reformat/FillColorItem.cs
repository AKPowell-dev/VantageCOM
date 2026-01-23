using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Colors;
using PowerPointAddIn1.DeckCheck.Fix;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class FillColorItem : ColorItem
{
	public FillColorItem(int intColor, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intIndex, int intTotal)
		: base(intColor, listObjects, intIndex, template, navItemTemplate, AH.A(49920), intTotal)
	{
	}

	public override void Reformat(int intNewColor, int intOldColor, ref List<string> listErrors)
	{
		NG.A.Application.StartNewUndoEntry();
		using List<NavigationItem>.Enumerator enumerator = base.Objects.GetEnumerator();
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		while (enumerator.MoveNext())
		{
			IndexedObject indexedObject = enumerator.Current.IndexedObject;
			try
			{
				if (indexedObject.Child is Microsoft.Office.Interop.PowerPoint.Shape)
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
						((Microsoft.Office.Interop.PowerPoint.Shape)indexedObject.Child).Fill.ForeColor.RGB = intNewColor;
						break;
					}
				}
				else if (indexedObject.Child is Cell)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						((Cell)indexedObject.Child).Shape.Fill.ForeColor.RGB = intNewColor;
						break;
					}
				}
				else if (indexedObject.Shape.HasChart == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						if (indexedObject.Child is ChartArea)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								A(((ChartArea)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is PlotArea)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								A(((PlotArea)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is Legend)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								A(((Legend)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is ChartTitle)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								A(((ChartTitle)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is DataTable)
						{
							A(((DataTable)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
						if (indexedObject.Child is AxisTitle)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								A(((AxisTitle)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is Axis)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								A(((Axis)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is Gridlines)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								A(((Gridlines)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is HiLoLines)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								A(((HiLoLines)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is DropLines)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								A(((DropLines)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is UpBars)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								A(((UpBars)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is DownBars)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								A(((DownBars)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is IMsoErrorBars)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								A(((IMsoErrorBars)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is IMsoLeaderLines)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								A(((IMsoLeaderLines)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (indexedObject.Child is IMsoTrendline)
						{
							A(((IMsoTrendline)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
						if (indexedObject.Child is IMsoSeries)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								try
								{
									Charts.A((IMsoSeries)indexedObject.Child, intNewColor, intOldColor, A, this.A);
								}
								catch (Exception projectError)
								{
									ProjectData.SetProjectError(projectError);
									ProjectData.ClearProjectError();
								}
								break;
							}
							break;
						}
						if (indexedObject.Child is ChartPoint)
						{
							if (indexedObject.IsMarker)
							{
								B((ChartPoint)indexedObject.Child, intNewColor, intOldColor);
							}
							else
							{
								A((ChartPoint)indexedObject.Child, intNewColor, intOldColor);
							}
							break;
						}
						if (indexedObject.Child is IMsoDataLabel)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								A(((IMsoDataLabel)indexedObject.Child).Format, intNewColor, intOldColor);
								break;
							}
							break;
						}
						if (!(indexedObject.Child is IMsoLegendKey))
						{
							break;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							A((IMsoLegendKey)indexedObject.Child, intNewColor, intOldColor, indexedObject.Shape, indexedObject.IsMarker);
							break;
						}
						break;
					}
				}
				else if (indexedObject.Child is Microsoft.Office.Core.Shape)
				{
					A(((Microsoft.Office.Core.Shape)indexedObject.Child).Fill, intNewColor, intOldColor);
				}
				else if (indexedObject.Child is SmartArt)
				{
					try
					{
						enumerator2 = indexedObject.Shape.SmartArt.AllNodes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							SmartArtNode smartArtNode = (SmartArtNode)enumerator2.Current;
							try
							{
								enumerator3 = smartArtNode.Shapes.GetEnumerator();
								while (enumerator3.MoveNext())
								{
									Microsoft.Office.Core.Shape shape = (Microsoft.Office.Core.Shape)enumerator3.Current;
									A(shape.Fill, intNewColor, intOldColor);
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_05dd;
									}
									continue;
									end_IL_05dd:
									break;
								}
							}
							finally
							{
								if (enumerator3 is IDisposable)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										(enumerator3 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0611;
							}
							continue;
							end_IL_0611:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				listErrors.Add(ex2.Message);
				ProjectData.ClearProjectError();
			}
			indexedObject = null;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void A(ChartFormat A, int B, int C)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor = A.Fill.ForeColor;
			if (foreColor.RGB == C)
			{
				foreColor.RGB = B;
			}
			foreColor = null;
			Microsoft.Office.Interop.PowerPoint.ColorFormat backColor = A.Fill.BackColor;
			if (backColor.RGB == C)
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
				backColor.RGB = B;
			}
			backColor = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(IMsoChartFormat A, int B, int C)
	{
		this.A(A.Fill, B, C);
	}

	private void A(IMsoLegendKey A, int B, int C, Microsoft.Office.Interop.PowerPoint.Shape D, bool E)
	{
		if (E)
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
					if (A.MarkerBackgroundColor == C)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								A.MarkerBackgroundColor = B;
								return;
							}
						}
					}
					return;
				}
			}
		}
		if (clsCharts.UsesLegendLinesForSeriesClrs(D.Chart))
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					if (A.Format.Line.ForeColor.RGB == C)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								A.Format.Line.ForeColor.RGB = B;
								return;
							}
						}
					}
					return;
				}
			}
		}
		if (!Operators.ConditionalCompareObjectEqual(A.Interior.Color, C, TextCompare: false))
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
			A.Interior.Color = B;
			return;
		}
	}

	private void A(ChartPoint A, int B, int C)
	{
		if (A.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
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
			if (A.MarkerBackgroundColor != Base.TRANSPARENT)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (A.MarkerBackgroundColor == C)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									A.MarkerBackgroundColor = B;
									return;
								}
							}
						}
						return;
					}
				}
			}
		}
		this.A(A.Format, B, C);
	}

	private void A(Microsoft.Office.Core.FillFormat A, int B, int C)
	{
		try
		{
			if (A.Visible != MsoTriState.msoTrue)
			{
				return;
			}
			Microsoft.Office.Core.ColorFormat foreColor = A.ForeColor;
			if (foreColor.RGB == C)
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
				foreColor.RGB = B;
			}
			foreColor = null;
			Microsoft.Office.Core.ColorFormat backColor = A.BackColor;
			if (backColor.RGB == C)
			{
				backColor.RGB = B;
			}
			backColor = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private bool A(Microsoft.Office.Core.FillFormat A, int B, int C)
	{
		try
		{
			if (A.Visible == MsoTriState.msoTrue)
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
					int rGB = A.ForeColor.RGB;
					if (rGB == C)
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
						if (!object.Equals(rGB, B))
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									return true;
								}
							}
						}
					}
					int rGB2 = A.BackColor.RGB;
					if (rGB2 != C)
					{
						break;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						if (object.Equals(rGB2, B))
						{
							break;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							return true;
						}
					}
					break;
				}
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		return false;
	}

	private void B(ChartPoint A, int B, int C)
	{
		if (A.MarkerBackgroundColor != C)
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
			A.MarkerBackgroundColor = B;
			return;
		}
	}
}
