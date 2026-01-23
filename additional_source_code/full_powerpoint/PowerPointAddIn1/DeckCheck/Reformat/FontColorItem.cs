using System;
using System.Collections.Generic;
using System.Windows;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing.UI.Reformat;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class FontColorItem : ColorItem
{
	public FontColorItem(int intColor, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intIndex, int intTotal)
		: base(intColor, listObjects, intIndex, template, navItemTemplate, AH.A(49982), intTotal)
	{
		Tuple<List<NavigationItem>, int> tuple = GenerateChildren(listObjects, ((BaseItem)this).Header);
		base.Objects = tuple.Item1;
		((BaseItem)this).Quantity = tuple.Item2;
		tuple = null;
	}

	public override void Reformat(int intNewColor, int intOldColor, ref List<string> listErrors)
	{
		NG.A.Application.StartNewUndoEntry();
		using List<NavigationItem>.Enumerator enumerator = base.Objects.GetEnumerator();
		while (enumerator.MoveNext())
		{
			IndexedObject indexedObject = enumerator.Current.IndexedObject;
			try
			{
				if (indexedObject.Child is TextRange2)
				{
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
						Font2 font = ((TextRange2)indexedObject.Child).Font;
						if (font.Highlight.RGB == intOldColor)
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
							if (intOldColor == 0)
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
								if (font.Highlight.Type != MsoColorType.msoColorTypeRGB)
								{
									goto IL_010b;
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
							float size = font.Size;
							string name = font.Name;
							int rGB = font.Fill.ForeColor.RGB;
							font.Highlight.RGB = intNewColor;
							font.Size = size;
							font.Name = name;
							font.Fill.ForeColor.RGB = rGB;
						}
						goto IL_010b;
						IL_010b:
						if (font.UnderlineStyle != MsoTextUnderlineType.msoNoUnderline)
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
							if (font.UnderlineColor.RGB == intOldColor)
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
								float size2 = font.Size;
								string name2 = font.Name;
								int rGB2 = font.Fill.ForeColor.RGB;
								font.UnderlineColor.RGB = intNewColor;
								font.Size = size2;
								font.Name = name2;
								font.Fill.ForeColor.RGB = rGB2;
							}
						}
						Microsoft.Office.Core.LineFormat line = font.Line;
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
							if (line.ForeColor.RGB == intOldColor)
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
								line.ForeColor.RGB = intNewColor;
							}
						}
						line = null;
						Microsoft.Office.Core.ColorFormat foreColor = font.Fill.ForeColor;
						if (foreColor.RGB == intOldColor)
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
							foreColor.RGB = intNewColor;
						}
						foreColor = null;
						Microsoft.Office.Core.ColorFormat backColor = font.Fill.BackColor;
						if (backColor.RGB == intOldColor)
						{
							backColor.RGB = intNewColor;
						}
						backColor = null;
						font = null;
						break;
					}
				}
				else if (indexedObject.Child is BulletFormat2)
				{
					Microsoft.Office.Core.FillFormat fill = ((BulletFormat2)indexedObject.Child).Font.Fill;
					if (fill.ForeColor.RGB == intOldColor)
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
						fill.ForeColor.RGB = intNewColor;
					}
					if (fill.BackColor.RGB == intOldColor)
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
						fill.BackColor.RGB = intNewColor;
					}
					fill = null;
				}
				else if (indexedObject.Child is IMsoDataLabels)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						Microsoft.Office.Core.ChartFont font2 = ((IMsoDataLabels)indexedObject.Child).Font;
						if (Operators.ConditionalCompareObjectEqual(font2.Color, intOldColor, TextCompare: false))
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
							font2.Color = intNewColor;
						}
						font2 = null;
						break;
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
						Microsoft.Office.Core.ChartFont font3 = ((IMsoDataLabel)indexedObject.Child).Font;
						if (Operators.ConditionalCompareObjectEqual(font3.Color, intOldColor, TextCompare: false))
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
							font3.Color = intNewColor;
						}
						font3 = null;
						break;
					}
				}
				else if (indexedObject.Child is ChartTitle)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						ChartTitle chartTitle = (ChartTitle)indexedObject.Child;
						if (clsCharts.ImplsFont(chartTitle))
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
							Microsoft.Office.Interop.PowerPoint.ChartFont font4 = chartTitle.Font;
							if (Operators.ConditionalCompareObjectEqual(font4.Color, intOldColor, TextCompare: false))
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
								font4.Color = intNewColor;
							}
							font4 = null;
						}
						chartTitle = null;
						break;
					}
				}
				else if (indexedObject.Child is AxisTitle)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						Microsoft.Office.Interop.PowerPoint.ChartFont font5 = ((AxisTitle)indexedObject.Child).Font;
						if (Operators.ConditionalCompareObjectEqual(font5.Color, intOldColor, TextCompare: false))
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
							font5.Color = intNewColor;
						}
						font5 = null;
						break;
					}
				}
				else if (indexedObject.Child is TickLabels)
				{
					Microsoft.Office.Interop.PowerPoint.ChartFont font6 = ((TickLabels)indexedObject.Child).Font;
					if (Operators.ConditionalCompareObjectEqual(font6.Color, intOldColor, TextCompare: false))
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
						font6.Color = intNewColor;
					}
					font6 = null;
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
						Microsoft.Office.Interop.PowerPoint.ChartFont font7 = ((Legend)indexedObject.Child).Font;
						if (Operators.ConditionalCompareObjectEqual(font7.Color, intOldColor, TextCompare: false))
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
							font7.Color = intNewColor;
						}
						font7 = null;
						break;
					}
				}
				else if (indexedObject.Child is DataTable)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						Microsoft.Office.Interop.PowerPoint.ChartFont font8 = ((DataTable)indexedObject.Child).Font;
						if (Operators.ConditionalCompareObjectEqual(font8.Color, intOldColor, TextCompare: false))
						{
							font8.Color = intNewColor;
						}
						font8 = null;
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
}
