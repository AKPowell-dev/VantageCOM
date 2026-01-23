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

public sealed class FontFamilyItem : BaseItem
{
	[CompilerGenerated]
	private string A;

	public string Family
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

	public FontFamilyItem(string strFamily, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal)
		: base(intTotal, 0, listObjects, template, navItemTemplate, AH.A(50005))
	{
		((BaseItem)this).Label = strFamily;
		Family = strFamily;
		Tuple<List<NavigationItem>, int> tuple = GenerateChildren(listObjects, ((BaseItem)this).Header);
		base.Objects = tuple.Item1;
		((BaseItem)this).Quantity = tuple.Item2;
		tuple = null;
	}

	public void Reformat(string strNewFont, ref List<string> listErrors)
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
						((TextRange2)indexedObject.Child).Font.Name = strNewFont;
						break;
					}
				}
				else if (indexedObject.Child is BulletFormat2)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						((BulletFormat2)indexedObject.Child).Font.Name = strNewFont;
						_ = null;
						break;
					}
				}
				else if (indexedObject.Child is IMsoDataLabels)
				{
					((IMsoDataLabels)indexedObject.Child).Font.Name = strNewFont;
					_ = null;
				}
				else if (indexedObject.Child is IMsoDataLabel)
				{
					((IMsoDataLabel)indexedObject.Child).Font.Name = strNewFont;
					_ = null;
				}
				else if (indexedObject.Child is ChartTitle)
				{
					while (true)
					{
						switch (6)
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
							chartTitle.Font.Name = strNewFont;
							_ = null;
						}
						chartTitle = null;
						break;
					}
				}
				else if (indexedObject.Child is AxisTitle)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						((AxisTitle)indexedObject.Child).Font.Name = strNewFont;
						_ = null;
						break;
					}
				}
				else if (indexedObject.Child is TickLabels)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						((TickLabels)indexedObject.Child).Font.Name = strNewFont;
						_ = null;
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
						((Legend)indexedObject.Child).Font.Name = strNewFont;
						_ = null;
						break;
					}
				}
				else if (indexedObject.Child is DataTable)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						((DataTable)indexedObject.Child).Font.Name = strNewFont;
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
		while (true)
		{
			switch (1)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}
}
