using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class BorderColorItem : ColorItem
{
	public BorderColorItem(int intColor, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intIndex, int intTotal)
		: base(intColor, listObjects, intIndex, template, navItemTemplate, AH.A(49238), intTotal)
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
					((Microsoft.Office.Interop.PowerPoint.Shape)indexedObject.Child).Line.ForeColor.RGB = intNewColor;
				}
				else if (indexedObject.Child is Cell)
				{
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
						((Cell)indexedObject.Child).Shape.Line.ForeColor.RGB = intNewColor;
						break;
					}
				}
				else if (indexedObject.Shape.HasChart == MsoTriState.msoTrue)
				{
					if (indexedObject.Child is ChartArea)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							A(((ChartArea)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
					else if (indexedObject.Child is PlotArea)
					{
						A(((PlotArea)indexedObject.Child).Format, intNewColor, intOldColor);
					}
					else if (indexedObject.Child is Legend)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							A(((Legend)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
					else if (indexedObject.Child is ChartTitle)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							A(((ChartTitle)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
					else if (indexedObject.Child is DataTable)
					{
						A(((DataTable)indexedObject.Child).Format, intNewColor, intOldColor);
					}
					else if (indexedObject.Child is AxisTitle)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							A(((AxisTitle)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
					else if (indexedObject.Child is Axis)
					{
						A(((Axis)indexedObject.Child).Format, intNewColor, intOldColor);
					}
					else if (indexedObject.Child is Gridlines)
					{
						A(((Gridlines)indexedObject.Child).Format, intNewColor, intOldColor);
					}
					else if (indexedObject.Child is HiLoLines)
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
					}
					else if (indexedObject.Child is DropLines)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							A(((DropLines)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
					else if (indexedObject.Child is UpBars)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							A(((UpBars)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
					else if (indexedObject.Child is DownBars)
					{
						A(((DownBars)indexedObject.Child).Format, intNewColor, intOldColor);
					}
					else if (indexedObject.Child is IMsoErrorBars)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							A(((IMsoErrorBars)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
					else if (indexedObject.Child is IMsoLeaderLines)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							A(((IMsoLeaderLines)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
					else if (indexedObject.Child is IMsoTrendline)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							A(((IMsoTrendline)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
					else if (indexedObject.Child is IMsoSeries)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							A(((IMsoSeries)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
					else if (indexedObject.Child is ChartPoint)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							if (indexedObject.IsMarker)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									A((ChartPoint)indexedObject.Child, intNewColor, intOldColor);
									break;
								}
							}
							else
							{
								A(((ChartPoint)indexedObject.Child).Format, intNewColor, intOldColor);
							}
							break;
						}
					}
					else if (indexedObject.Child is IMsoDataLabel)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							A(((IMsoDataLabel)indexedObject.Child).Format, intNewColor, intOldColor);
							break;
						}
					}
				}
				else if (indexedObject.Child is Microsoft.Office.Core.Shape)
				{
					A(((Microsoft.Office.Core.Shape)indexedObject.Child).Line, intNewColor, intOldColor);
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
						enumerator2 = indexedObject.Shape.SmartArt.AllNodes.GetEnumerator();
						try
						{
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
											A(shape.Line, intNewColor, intOldColor);
										}
										while (true)
										{
											switch (2)
											{
											case 0:
												break;
											default:
												goto end_IL_056a;
											}
											continue;
											end_IL_056a:
											break;
										}
									}
									finally
									{
										IDisposable disposable2 = enumerator3 as IDisposable;
										if (disposable2 != null)
										{
											disposable2.Dispose();
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
									goto end_IL_0596;
								}
								continue;
								end_IL_0596:
								break;
							}
						}
						finally
						{
							IDisposable disposable = enumerator2 as IDisposable;
							if (disposable != null)
							{
								disposable.Dispose();
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
			switch (4)
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
		Microsoft.Office.Interop.PowerPoint.LineFormat line = A.Line;
		if (line.Visible == MsoTriState.msoTrue)
		{
			Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor = line.ForeColor;
			if (foreColor.RGB == C)
			{
				foreColor.RGB = B;
			}
			foreColor = null;
			Microsoft.Office.Interop.PowerPoint.ColorFormat backColor = line.BackColor;
			if (backColor.RGB == C)
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
				backColor.RGB = B;
			}
			backColor = null;
		}
		line = null;
	}

	private void A(IMsoChartFormat A, int B, int C)
	{
		this.A(A.Line, B, C);
	}

	private void A(Microsoft.Office.Core.LineFormat A, int B, int C)
	{
		Microsoft.Office.Core.LineFormat lineFormat = A;
		if (lineFormat.Visible == MsoTriState.msoTrue)
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
			Microsoft.Office.Core.ColorFormat foreColor = lineFormat.ForeColor;
			if (foreColor.RGB == C)
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
				foreColor.RGB = B;
			}
			foreColor = null;
			Microsoft.Office.Core.ColorFormat backColor = lineFormat.BackColor;
			if (backColor.RGB == C)
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
				backColor.RGB = B;
			}
			backColor = null;
		}
		lineFormat = null;
	}

	private void A(ChartPoint A, int B, int C)
	{
		if (A.MarkerForegroundColor != C)
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
			A.MarkerForegroundColor = B;
			return;
		}
	}
}
