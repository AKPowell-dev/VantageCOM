using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing.UI.Reformat;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Reformat;

public sealed class TextDecorationItem : BaseItem
{
	[CompilerGenerated]
	private TextDecoration m_A;

	public TextDecoration Decoration
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

	public TextDecorationItem(TextDecoration decor, List<IndexedObject> listObjects, DataTemplate template, DataTemplate navItemTemplate, int intTotal)
		: base(intTotal, 0, listObjects, template, navItemTemplate, AH.A(50897))
	{
		Tuple<List<NavigationItem>, int> tuple = GenerateChildren(listObjects, ((BaseItem)this).Header);
		base.Objects = tuple.Item1;
		((BaseItem)this).Quantity = tuple.Item2;
		tuple = null;
		Decoration = decor;
		string text = "";
		if (decor.Bold)
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
			if (decor.Italic)
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
				text = AH.A(50930);
				goto IL_00d2;
			}
		}
		if (decor.Bold)
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
			text = AH.A(50957);
		}
		else if (decor.Italic)
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
			text = AH.A(50966);
		}
		goto IL_00d2;
		IL_00d2:
		if (decor.UnderlineStyle != MsoTextUnderlineType.msoNoUnderline)
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
			if (text.Length > 0)
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
				text += AH.A(50979);
			}
			text = decor.UnderlineStyle switch
			{
				MsoTextUnderlineType.msoUnderlineSingleLine => text + AH.A(50986), 
				MsoTextUnderlineType.msoUnderlineDashHeavyLine => text + AH.A(51005), 
				MsoTextUnderlineType.msoUnderlineDashLine => text + AH.A(51036), 
				MsoTextUnderlineType.msoUnderlineDashLongHeavyLine => text + AH.A(51055), 
				MsoTextUnderlineType.msoUnderlineDashLongLine => text + AH.A(51096), 
				MsoTextUnderlineType.msoUnderlineDotDashHeavyLine => text + AH.A(51125), 
				MsoTextUnderlineType.msoUnderlineDotDashLine => text + AH.A(51164), 
				MsoTextUnderlineType.msoUnderlineDotDotDashHeavyLine => text + AH.A(51191), 
				MsoTextUnderlineType.msoUnderlineDotDotDashLine => text + AH.A(51238), 
				MsoTextUnderlineType.msoUnderlineDottedHeavyLine => text + AH.A(51273), 
				MsoTextUnderlineType.msoUnderlineDottedLine => text + AH.A(51308), 
				MsoTextUnderlineType.msoUnderlineDoubleLine => text + AH.A(51331), 
				MsoTextUnderlineType.msoUnderlineHeavyLine => text + AH.A(51354), 
				MsoTextUnderlineType.msoUnderlineWavyDoubleLine => text + AH.A(51375), 
				MsoTextUnderlineType.msoUnderlineWavyHeavyLine => text + AH.A(51408), 
				MsoTextUnderlineType.msoUnderlineWavyLine => text + AH.A(51439), 
				_ => text + AH.A(50986), 
			};
		}
		((BaseItem)this).Label = text;
	}

	public void Reformat(TextDecorationOption opt, ref List<string> listErrors)
	{
		TextDecoration decoration = opt.Decoration;
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
						switch (7)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						Font2 font = ((TextRange2)indexedObject.Child).Font;
						if (decoration.Bold)
						{
							font.Bold = MsoTriState.msoTrue;
						}
						else
						{
							font.Bold = MsoTriState.msoFalse;
						}
						if (decoration.Italic)
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
							font.Italic = MsoTriState.msoTrue;
						}
						else
						{
							font.Italic = MsoTriState.msoFalse;
						}
						font.UnderlineStyle = decoration.UnderlineStyle;
						font = null;
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
								switch (5)
								{
								case 0:
									continue;
								}
								break;
							}
							A(chartTitle.Font, decoration);
						}
						chartTitle = null;
						break;
					}
				}
				else if (indexedObject.Child is AxisTitle)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						A(((AxisTitle)indexedObject.Child).Font, decoration);
						break;
					}
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
						A(((TickLabels)indexedObject.Child).Font, decoration);
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
						A(((Legend)indexedObject.Child).Font, decoration);
						break;
					}
				}
				else if (indexedObject.Child is DataTable)
				{
					A(((DataTable)indexedObject.Child).Font, decoration);
				}
				else if (indexedObject.Child is IMsoDataLabel)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						A(((IMsoDataLabel)indexedObject.Child).Font, decoration);
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
			switch (3)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.ChartFont A, TextDecoration B)
	{
		Microsoft.Office.Interop.PowerPoint.ChartFont chartFont = A;
		chartFont.Bold = B.Bold;
		chartFont.Italic = B.Italic;
		switch (B.UnderlineStyle)
		{
		case MsoTextUnderlineType.msoNoUnderline:
			chartFont.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
			break;
		case MsoTextUnderlineType.msoUnderlineSingleLine:
			chartFont.Underline = XlUnderlineStyle.xlUnderlineStyleSingle;
			break;
		case MsoTextUnderlineType.msoUnderlineDoubleLine:
			chartFont.Underline = XlUnderlineStyle.xlUnderlineStyleDouble;
			break;
		}
		chartFont = null;
	}

	private void A(Microsoft.Office.Core.ChartFont A, TextDecoration B)
	{
		Microsoft.Office.Core.ChartFont chartFont = A;
		chartFont.Bold = B.Bold;
		chartFont.Italic = B.Italic;
		switch (B.UnderlineStyle)
		{
		case MsoTextUnderlineType.msoNoUnderline:
			chartFont.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
			break;
		case MsoTextUnderlineType.msoUnderlineSingleLine:
			chartFont.Underline = XlUnderlineStyle.xlUnderlineStyleSingle;
			break;
		case MsoTextUnderlineType.msoUnderlineDoubleLine:
			chartFont.Underline = XlUnderlineStyle.xlUnderlineStyleDouble;
			break;
		}
		chartFont = null;
	}
}
