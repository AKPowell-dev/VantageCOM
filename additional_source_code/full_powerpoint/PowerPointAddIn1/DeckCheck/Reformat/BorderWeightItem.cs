using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class BorderWeightItem : BaseItem
{
	[CompilerGenerated]
	private float m_A;

	[CompilerGenerated]
	private double m_A;

	public float Weight
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

	public BorderWeightItem(float sngWeight, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal, int intIndex)
		: base(intTotal, intIndex, listObjects, template, navItemTemplate, AH.A(49316))
	{
		Weight = sngWeight;
		if (sngWeight == 0.25f)
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
					Thickness = 1.0;
					return;
				}
			}
		}
		if (sngWeight == 0.5f)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					Thickness = 1.0;
					return;
				}
			}
		}
		if (sngWeight == 0.75f)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					Thickness = 1.0;
					return;
				}
			}
		}
		if (sngWeight == 1f)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					Thickness = 2.0;
					return;
				}
			}
		}
		if (sngWeight == 1.5f)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					Thickness = 2.0;
					return;
				}
			}
		}
		if (sngWeight == 2.25f)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					Thickness = 3.0;
					return;
				}
			}
		}
		if (sngWeight == 3f)
		{
			Thickness = 4.0;
			return;
		}
		if (sngWeight == 4.5f)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					Thickness = 5.0;
					return;
				}
			}
		}
		if (sngWeight != 6f)
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
			Thickness = 6.0;
			return;
		}
	}

	public void Reformat(float sngNewWeight, ref List<string> listErrors)
	{
		NG.A.Application.StartNewUndoEntry();
		using List<NavigationItem>.Enumerator enumerator = base.Objects.GetEnumerator();
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
						((Microsoft.Office.Interop.PowerPoint.Shape)indexedObject.Child).Line.Weight = sngNewWeight;
						break;
					}
				}
				else if (indexedObject.Shape.HasChart == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						if (indexedObject.Child is ChartArea)
						{
							A(((ChartArea)indexedObject.Child).Format, sngNewWeight);
							break;
						}
						if (indexedObject.Child is PlotArea)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								A(((PlotArea)indexedObject.Child).Format, sngNewWeight);
								break;
							}
							break;
						}
						if (indexedObject.Child is Legend)
						{
							A(((Legend)indexedObject.Child).Format, sngNewWeight);
							break;
						}
						if (indexedObject.Child is ChartTitle)
						{
							A(((ChartTitle)indexedObject.Child).Format, sngNewWeight);
							break;
						}
						if (indexedObject.Child is DataTable)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								A(((DataTable)indexedObject.Child).Format, sngNewWeight);
								break;
							}
							break;
						}
						if (indexedObject.Child is AxisTitle)
						{
							A(((AxisTitle)indexedObject.Child).Format, sngNewWeight);
							break;
						}
						if (indexedObject.Child is Axis)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								A(((Axis)indexedObject.Child).Format, sngNewWeight);
								break;
							}
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
								A(((Gridlines)indexedObject.Child).Format, sngNewWeight);
								break;
							}
							break;
						}
						if (indexedObject.Child is HiLoLines)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								A(((HiLoLines)indexedObject.Child).Format, sngNewWeight);
								break;
							}
							break;
						}
						if (indexedObject.Child is DropLines)
						{
							A(((DropLines)indexedObject.Child).Format, sngNewWeight);
							break;
						}
						if (indexedObject.Child is UpBars)
						{
							A(((UpBars)indexedObject.Child).Format, sngNewWeight);
							break;
						}
						if (indexedObject.Child is DownBars)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								A(((DownBars)indexedObject.Child).Format, sngNewWeight);
								break;
							}
							break;
						}
						if (indexedObject.Child is IMsoErrorBars)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								A(((IMsoErrorBars)indexedObject.Child).Format, sngNewWeight);
								break;
							}
							break;
						}
						if (indexedObject.Child is IMsoLeaderLines)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								A(((IMsoLeaderLines)indexedObject.Child).Format, sngNewWeight);
								break;
							}
							break;
						}
						if (indexedObject.Child is IMsoTrendline)
						{
							A(((IMsoTrendline)indexedObject.Child).Format, sngNewWeight);
							break;
						}
						if (indexedObject.Child is IMsoSeries)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								A(((IMsoSeries)indexedObject.Child).Format, sngNewWeight);
								break;
							}
							break;
						}
						if (indexedObject.Child is ChartPoint)
						{
							A(((ChartPoint)indexedObject.Child).Format, sngNewWeight);
							break;
						}
						if (!(indexedObject.Child is IMsoDataLabel))
						{
							break;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							A(((IMsoDataLabel)indexedObject.Child).Format, sngNewWeight);
							break;
						}
						break;
					}
				}
				else if (indexedObject.Child is SmartArt)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						foreach (SmartArtNode allNode in indexedObject.Shape.SmartArt.AllNodes)
						{
							try
							{
								enumerator3 = allNode.Shapes.GetEnumerator();
								while (enumerator3.MoveNext())
								{
									Microsoft.Office.Core.Shape shape = (Microsoft.Office.Core.Shape)enumerator3.Current;
									A(shape.Line, sngNewWeight);
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_04af;
									}
									continue;
									end_IL_04af:
									break;
								}
							}
							finally
							{
								if (enumerator3 is IDisposable)
								{
									while (true)
									{
										switch (6)
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
						break;
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
			switch (7)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void A(ChartFormat A, float B)
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat line = A.Line;
			if (line.Visible == MsoTriState.msoTrue)
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
				line.Weight = B;
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

	private void A(IMsoChartFormat A, float B)
	{
		this.A((Microsoft.Office.Core.LineFormat)A.Fill, B);
	}

	private void A(Microsoft.Office.Core.LineFormat A, float B)
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
				A.Weight = B;
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
