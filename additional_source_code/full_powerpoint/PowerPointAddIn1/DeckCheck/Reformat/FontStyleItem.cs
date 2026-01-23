using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing.UI.Reformat;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class FontStyleItem : BaseItem
{
	[CompilerGenerated]
	private FontStyle A;

	public FontStyle Style
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public FontStyleItem(FontStyle sty, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal)
		: base(intTotal, 0, listObjects, template, navItemTemplate, AH.A(50032))
	{
		((BaseItem)this).Label = sty.Family + AH.A(14625) + sty.Size + AH.A(17909);
		Style = sty;
		Tuple<List<NavigationItem>, int> tuple = GenerateChildren(listObjects, ((BaseItem)this).Header);
		base.Objects = tuple.Item1;
		((BaseItem)this).Quantity = tuple.Item2;
		tuple = null;
	}

	public void Reformat(string strNewFont, float sngNewSize, ref List<string> listErrors)
	{
		NG.A.Application.StartNewUndoEntry();
		foreach (NavigationItem @object in base.Objects)
		{
			IndexedObject indexedObject = @object.IndexedObject;
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
						font.Name = strNewFont;
						font.Size = sngNewSize;
						_ = null;
						break;
					}
				}
				else if (indexedObject.Child is BulletFormat2)
				{
					Font2 font2 = ((BulletFormat2)indexedObject.Child).Font;
					font2.Name = strNewFont;
					font2.Size = sngNewSize;
					_ = null;
				}
				else if (indexedObject.Child is IMsoDataLabels)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						Microsoft.Office.Core.ChartFont font3 = ((IMsoDataLabels)indexedObject.Child).Font;
						font3.Name = strNewFont;
						font3.Size = sngNewSize;
						_ = null;
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
						Microsoft.Office.Core.ChartFont font4 = ((IMsoDataLabel)indexedObject.Child).Font;
						font4.Name = strNewFont;
						font4.Size = sngNewSize;
						_ = null;
						break;
					}
				}
				else if (indexedObject.Child is ChartTitle)
				{
					while (true)
					{
						switch (7)
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
							Microsoft.Office.Interop.PowerPoint.ChartFont font5 = chartTitle.Font;
							font5.Name = strNewFont;
							font5.Size = sngNewSize;
							_ = null;
						}
						chartTitle = null;
						break;
					}
				}
				else if (indexedObject.Child is AxisTitle)
				{
					Microsoft.Office.Interop.PowerPoint.ChartFont font6 = ((AxisTitle)indexedObject.Child).Font;
					font6.Name = strNewFont;
					font6.Size = sngNewSize;
					_ = null;
				}
				else if (indexedObject.Child is TickLabels)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						Microsoft.Office.Interop.PowerPoint.ChartFont font7 = ((TickLabels)indexedObject.Child).Font;
						font7.Name = strNewFont;
						font7.Size = sngNewSize;
						_ = null;
						break;
					}
				}
				else if (indexedObject.Child is Legend)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						Microsoft.Office.Interop.PowerPoint.ChartFont font8 = ((Legend)indexedObject.Child).Font;
						font8.Name = strNewFont;
						font8.Size = sngNewSize;
						_ = null;
						break;
					}
				}
				else if (indexedObject.Child is DataTable)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						Microsoft.Office.Interop.PowerPoint.ChartFont font9 = ((DataTable)indexedObject.Child).Font;
						font9.Name = strNewFont;
						font9.Size = sngNewSize;
						_ = null;
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
	}
}
