using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class BorderDashItem : BaseItem
{
	[CompilerGenerated]
	private MsoLineDashStyle m_A;

	[CompilerGenerated]
	private DoubleCollection m_A;

	[CompilerGenerated]
	private PenLineCap m_A;

	[CompilerGenerated]
	private double m_A;

	public MsoLineDashStyle Style
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

	public DoubleCollection DashArray
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

	public PenLineCap Cap
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

	public double Thickness
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

	public BorderDashItem(int intStyle, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal, int intIndex)
		: base(intTotal, intIndex, listObjects, template, navItemTemplate, AH.A(49279))
	{
		Thickness = 3.0;
		Cap = PenLineCap.Flat;
		Style = (MsoLineDashStyle)intStyle;
		switch (Style)
		{
		case MsoLineDashStyle.msoLineSolid:
			DashArray = new DoubleCollection();
			break;
		case MsoLineDashStyle.msoLineRoundDot:
		case MsoLineDashStyle.msoLineSysDot:
			DashArray = new DoubleCollection(new double[2] { 0.0, 2.0 });
			Cap = PenLineCap.Round;
			Thickness = 4.0;
			break;
		case MsoLineDashStyle.msoLineSquareDot:
		case MsoLineDashStyle.msoLineSysDash:
			DashArray = new DoubleCollection(new double[1] { 1.0 });
			break;
		case MsoLineDashStyle.msoLineDash:
			DashArray = new DoubleCollection(new double[2] { 4.0, 2.0 });
			break;
		case MsoLineDashStyle.msoLineDashDot:
			DashArray = new DoubleCollection(new double[4] { 4.0, 2.0, 2.0, 2.0 });
			break;
		case MsoLineDashStyle.msoLineLongDash:
			DashArray = new DoubleCollection(new double[2] { 8.0, 2.0 });
			break;
		case MsoLineDashStyle.msoLineLongDashDot:
			DashArray = new DoubleCollection(new double[4] { 8.0, 2.0, 2.0, 2.0 });
			break;
		case MsoLineDashStyle.msoLineLongDashDotDot:
			DashArray = new DoubleCollection(new double[6] { 8.0, 2.0, 2.0, 2.0, 2.0, 2.0 });
			break;
		case MsoLineDashStyle.msoLineDashDotDot:
			break;
		}
	}

	public void Reformat(MsoLineDashStyle msoNewStyle, ref List<string> listErrors)
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
						switch (6)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						((Microsoft.Office.Interop.PowerPoint.Shape)indexedObject.Child).Line.DashStyle = msoNewStyle;
						break;
					}
				}
				else if (indexedObject.Shape.HasChart == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						if (indexedObject.Child is ChartArea)
						{
							A(((ChartArea)indexedObject.Child).Format, msoNewStyle);
							break;
						}
						if (indexedObject.Child is PlotArea)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								A(((PlotArea)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is Legend)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								A(((Legend)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is ChartTitle)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								A(((ChartTitle)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is DataTable)
						{
							A(((DataTable)indexedObject.Child).Format, msoNewStyle);
							break;
						}
						if (indexedObject.Child is AxisTitle)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								A(((AxisTitle)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is Axis)
						{
							A(((Axis)indexedObject.Child).Format, msoNewStyle);
							break;
						}
						if (indexedObject.Child is Gridlines)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								A(((Gridlines)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is HiLoLines)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								A(((HiLoLines)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is DropLines)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								A(((DropLines)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is UpBars)
						{
							A(((UpBars)indexedObject.Child).Format, msoNewStyle);
							break;
						}
						if (indexedObject.Child is DownBars)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								A(((DownBars)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is IMsoErrorBars)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								A(((IMsoErrorBars)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is IMsoLeaderLines)
						{
							A(((IMsoLeaderLines)indexedObject.Child).Format, msoNewStyle);
							break;
						}
						if (indexedObject.Child is IMsoTrendline)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								A(((IMsoTrendline)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is IMsoSeries)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								A(((IMsoSeries)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (indexedObject.Child is ChartPoint)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								A(((ChartPoint)indexedObject.Child).Format, msoNewStyle);
								break;
							}
							break;
						}
						if (!(indexedObject.Child is IMsoDataLabel))
						{
							break;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							A(((IMsoDataLabel)indexedObject.Child).Format, msoNewStyle);
							break;
						}
						break;
					}
				}
				else if (indexedObject.Child is SmartArt)
				{
					try
					{
						enumerator2 = indexedObject.Shape.SmartArt.AllNodes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							SmartArtNode smartArtNode = (SmartArtNode)enumerator2.Current;
							{
								enumerator3 = smartArtNode.Shapes.GetEnumerator();
								try
								{
									while (enumerator3.MoveNext())
									{
										Microsoft.Office.Core.Shape shape = (Microsoft.Office.Core.Shape)enumerator3.Current;
										A(shape.Line, msoNewStyle);
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_04ce;
										}
										continue;
										end_IL_04ce:
										break;
									}
								}
								finally
								{
									IDisposable disposable = enumerator3 as IDisposable;
									if (disposable != null)
									{
										disposable.Dispose();
									}
								}
							}
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_04f8;
							}
							continue;
							end_IL_04f8:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (4)
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
			switch (2)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void A(ChartFormat A, MsoLineDashStyle B)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat line = A.Line;
			if (line.Visible == MsoTriState.msoTrue)
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
				line.DashStyle = B;
			}
			line = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(IMsoChartFormat A, MsoLineDashStyle B)
	{
		this.A((Microsoft.Office.Core.LineFormat)A.Fill, B);
	}

	private void A(Microsoft.Office.Core.LineFormat A, MsoLineDashStyle B)
	{
		try
		{
			if (A.Visible != MsoTriState.msoTrue)
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
				A.DashStyle = B;
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
