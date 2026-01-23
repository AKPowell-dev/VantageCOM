using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck;

public class BaseError : BaseError
{
	[CompilerGenerated]
	private Slide m_A;

	[CompilerGenerated]
	private Microsoft.Office.Interop.PowerPoint.Shape m_A;

	[CompilerGenerated]
	private List<Microsoft.Office.Interop.PowerPoint.Shape> m_A;

	[CompilerGenerated]
	private PlotArea m_A;

	[CompilerGenerated]
	private Axis m_A;

	[CompilerGenerated]
	private AxisTitle m_A;

	[CompilerGenerated]
	private Legend m_A;

	[CompilerGenerated]
	private ChartTitle m_A;

	[CompilerGenerated]
	private DataTable m_A;

	[CompilerGenerated]
	private ChartGroup m_A;

	[CompilerGenerated]
	private ErrorType m_A;

	internal Slide Slide
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

	internal Microsoft.Office.Interop.PowerPoint.Shape Shape
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

	internal List<Microsoft.Office.Interop.PowerPoint.Shape> Shapes
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

	internal PlotArea PlotArea
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

	internal Axis Axis
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

	internal AxisTitle AxisTitle
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

	internal Legend Legend
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

	internal ChartTitle ChartTitle
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

	internal DataTable DataTable
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

	internal ChartGroup ChartGroup
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

	internal ErrorType Type
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

	public BaseError(ErrorType errType, Severity sev, Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, bool blnHasFix, bool blnCanFixMultiple = false)
		: base(sev, blnHasFix, blnCanFixMultiple)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		Slide = null;
		Shape = null;
		Shapes = null;
		PlotArea = null;
		Axis = null;
		AxisTitle = null;
		Legend = null;
		ChartTitle = null;
		DataTable = null;
		ChartGroup = null;
		Type = errType;
		Slide = sld;
		Shape = shp;
		ErrorType type = Type;
		if (type <= ErrorType.ProofingLanguage)
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
			switch (type)
			{
			default:
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				goto IL_067e;
			case ErrorType.ColorPaletteFill:
			case ErrorType.ColorPaletteFont:
			case ErrorType.ColorPaletteBorder:
			case ErrorType.PlaceholderFillMismatch:
			case ErrorType.PlaceholderFontColorMismatch:
				break;
			case ErrorType.Text:
				goto IL_034d;
			case ErrorType.FillTransparency:
				((BaseError)this).IconData = Icons.OPACITY;
				((BaseError)this).IconPadding = new Thickness(1.0);
				return;
			case ErrorType.FillGradient:
				((BaseError)this).IconData = Icons.GRADIENT;
				((BaseError)this).IconPadding = new Thickness(3.0);
				return;
			case ErrorType.MultipleFontFamilies:
			case ErrorType.ShrinkTextOnOverflow:
			case ErrorType.PlaceholderFontStyleMismatch:
			case ErrorType.IllegalFont:
				goto IL_03cd;
			case ErrorType.LineSpacing:
				((BaseError)this).IconData = Icons.LINE_SPACING;
				((BaseError)this).IconPadding = new Thickness(3.0);
				return;
			case ErrorType.BulletPunctuation:
			case ErrorType.BulletSize:
			case ErrorType.BulletFontFamily:
			case ErrorType.BulletIndent:
			case ErrorType.PlaceholderBulletMismatch:
				((BaseError)this).IconData = Icons.BULLETS;
				((BaseError)this).IconPadding = new Thickness(1.0);
				return;
			case ErrorType.PlaceholderIndentMismatch:
				((BaseError)this).IconData = Icons.INDENT;
				((BaseError)this).IconPadding = new Thickness(3.0);
				return;
			case ErrorType.ShapeOutOfBounds:
			case ErrorType.MasterShapePosition:
			case ErrorType.MisalignedShape:
			case ErrorType.RotatedShape:
			case ErrorType.CrookedLine:
				goto IL_0468;
			case ErrorType.PlaceholderLayoutMismatch:
			case ErrorType.PlaceholderMarginsMismatch:
				((BaseError)this).IconData = Icons.PLACEHOLDER;
				((BaseError)this).IconPadding = new Thickness(3.0);
				return;
			case ErrorType.OverlappingText:
				((BaseError)this).IconData = Icons.OVERLAP;
				((BaseError)this).IconPadding = new Thickness(4.0);
				return;
			case ErrorType.FootnoteMissing:
			case ErrorType.FootnotesSequence:
				((BaseError)this).IconData = Icons.FOOTNOTE;
				((BaseError)this).IconPadding = new Thickness(3.0, 4.0, 3.0, 2.0);
				return;
			case ErrorType.TableCellMargins:
				((BaseError)this).IconData = Icons.TABLE;
				((BaseError)this).IconPadding = new Thickness(3.0);
				return;
			case ErrorType.ImageCropping:
			case ErrorType.LinkedPicture:
			case ErrorType.AirplaneMode:
			case ErrorType.ImageDistortion:
				((BaseError)this).IconData = Icons.IMAGE;
				((BaseError)this).IconPadding = new Thickness(3.0);
				return;
			case ErrorType.LinkBroken:
			case ErrorType.LinkNewerVersionAvailable:
				((BaseError)this).IconData = Icons.LINK;
				((BaseError)this).IconPadding = new Thickness(2.0);
				return;
			case ErrorType.ProofingLanguage:
				((BaseError)this).IconData = Icons.LANGUAGE;
				((BaseError)this).IconPadding = new Thickness(4.0);
				return;
			case ErrorType.HiddenSlide:
			case ErrorType.HiddenShape:
				((BaseError)this).IconData = Icons.HIDDEN;
				((BaseError)this).IconPadding = new Thickness(3.0);
				return;
			case ErrorType.Hyperlinks:
				((BaseError)this).IconData = Icons.HYPERLINK;
				((BaseError)this).IconPadding = new Thickness(1.0);
				return;
			case (ErrorType)1:
			case (ErrorType)2:
			case (ErrorType)3:
			case (ErrorType)4:
			case (ErrorType)5:
			case (ErrorType)6:
			case (ErrorType)7:
			case (ErrorType)8:
			case (ErrorType)9:
			case (ErrorType)10:
			case (ErrorType)11:
			case (ErrorType)12:
			case (ErrorType)15:
			case (ErrorType)18:
			case ErrorType.TitleMissing:
			case (ErrorType)24:
			case (ErrorType)32:
			case (ErrorType)33:
			case (ErrorType)34:
			case (ErrorType)35:
			case (ErrorType)36:
			case (ErrorType)37:
			case (ErrorType)38:
			case (ErrorType)39:
			case (ErrorType)40:
			case (ErrorType)41:
			case (ErrorType)42:
			case (ErrorType)43:
			case (ErrorType)47:
			case (ErrorType)48:
			case (ErrorType)49:
			case (ErrorType)50:
			case ErrorType.AgendaNotUpdated:
			case ErrorType.AgendaMissing:
			case (ErrorType)67:
			case (ErrorType)68:
			case ErrorType.MissingSlideNumber:
			case ErrorType.ExcessSlideNumber:
				goto IL_067e;
			}
		}
		else
		{
			switch (type)
			{
			case ErrorType.SlideWordCount:
			case ErrorType.BulletWordCount:
				goto IL_034d;
			case ErrorType.CorporateDictionary:
				((BaseError)this).IconData = Icons.SPELLCHECK;
				((BaseError)this).IconPadding = new Thickness(3.0);
				return;
			case ErrorType.MaxMinFontSize:
			case ErrorType.FractionalFontSize:
			case ErrorType.StrikethroughFont:
				goto IL_03cd;
			case ErrorType.ShapeEffects:
			case ErrorType.TextEffects:
				((BaseError)this).IconData = Icons.EFFECTS;
				((BaseError)this).IconPadding = new Thickness(1.0);
				return;
			case ErrorType.Animation:
				((BaseError)this).IconData = Icons.TIMER;
				((BaseError)this).IconPadding = new Thickness(0.0);
				return;
			case ErrorType.Ink:
				((BaseError)this).IconData = Icons.MARKER;
				((BaseError)this).IconPadding = new Thickness(4.0);
				return;
			case ErrorType.SlideCount:
				((BaseError)this).IconData = Icons.CLOCK;
				((BaseError)this).IconPadding = new Thickness(1.0);
				return;
			case (ErrorType)111:
			case (ErrorType)112:
			case (ErrorType)113:
			case (ErrorType)114:
			case (ErrorType)115:
			case (ErrorType)116:
			case (ErrorType)117:
			case (ErrorType)118:
			case (ErrorType)119:
				goto IL_067e;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
			switch (type)
			{
			case ErrorType.ColorPaletteChartSeries:
			case ErrorType.ColorPaletteChartDataLabels:
			case ErrorType.ColorPaletteChartAreaFill:
			case ErrorType.ColorPaletteChartAreaBorder:
			case ErrorType.ColorPaletteChartPlotFill:
			case ErrorType.ColorPaletteChartPlotBorder:
			case ErrorType.ColorPaletteChartGridlines:
			case ErrorType.ColorPaletteChartLegendFont:
			case ErrorType.ColorPaletteChartLegendFill:
			case ErrorType.ColorPaletteChartLegendBorder:
			case ErrorType.ColorPaletteChartTitleFont:
			case ErrorType.ColorPaletteChartTitleFill:
			case ErrorType.ColorPaletteChartTitleBorder:
			case ErrorType.ColorPaletteChartAxisFont:
			case ErrorType.ColorPaletteChartAxisFill:
			case ErrorType.ColorPaletteChartAxisBorder:
			case ErrorType.ColorPaletteChartAxisTitleFont:
			case ErrorType.ColorPaletteChartAxisTitleFill:
			case ErrorType.ColorPaletteChartAxisTitleBorder:
			case ErrorType.ColorPaletteChartDataTableFont:
			case ErrorType.ColorPaletteChartDataTableFill:
			case ErrorType.ColorPaletteChartDataTableBorder:
			case ErrorType.ColorPaletteChartHiLoLines:
			case ErrorType.ColorPaletteChartDropLines:
				break;
			case ErrorType.ShapeOutsideMargins:
				goto IL_0468;
			case ErrorType.ChartLegendEntryMissing:
			case ErrorType.ChartDataLabelMissing:
			case ErrorType.ChartDataLabelsInconsistent:
			case ErrorType.ChartDataLabelNumberFormats:
				((BaseError)this).IconData = Icons.CHART;
				((BaseError)this).IconPadding = new Thickness(3.0, 2.0, 3.0, 4.0);
				return;
			default:
				goto IL_067e;
			}
		}
		((BaseError)this).IconData = Icons.PALETTE;
		((BaseError)this).IconPadding = new Thickness(0.0);
		return;
		IL_034d:
		((BaseError)this).IconData = Icons.TEXT_LINES;
		((BaseError)this).IconPadding = new Thickness(2.0);
		return;
		IL_067e:
		((BaseError)this).IconData = Icons.QUESTION_MARK;
		((BaseError)this).IconPadding = new Thickness(1.0);
		return;
		IL_0468:
		((BaseError)this).IconData = Icons.MOVE;
		((BaseError)this).IconPadding = new Thickness(1.0);
		return;
		IL_03cd:
		((BaseError)this).IconData = Icons.FONT;
		((BaseError)this).IconPadding = new Thickness(2.0, 3.0, 2.0, 1.0);
	}

	internal string A(List<TextRange2> A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		string text = "";
		checked
		{
			if (A != null)
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
				if (A.Count == 1)
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
					try
					{
						TextRange2 textRange = A[0];
						TextRange2 textRange2;
						if (textRange.Parent is TextRange2)
						{
							textRange2 = (TextRange2)textRange.Parent;
						}
						else if (!(textRange.Parent is Microsoft.Office.Interop.PowerPoint.TextFrame2))
						{
							textRange2 = ((!(textRange.Parent is Microsoft.Office.Core.TextFrame2)) ? B.TextFrame2.TextRange : ((Microsoft.Office.Core.TextFrame2)textRange.Parent).TextRange);
						}
						else
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
							textRange2 = ((Microsoft.Office.Interop.PowerPoint.TextFrame2)textRange.Parent).TextRange;
						}
						int length = textRange.Length;
						int num = textRange.Start - 1;
						int num2 = 1;
						int length2 = textRange.Length;
						textRange = null;
						if (length < textRange2.Text.Length)
						{
							MatchCollection matchCollection = new Regex(AH.A(47306)).Matches(textRange2.Text);
							{
								IEnumerator enumerator = matchCollection.GetEnumerator();
								try
								{
									while (true)
									{
										if (enumerator.MoveNext())
										{
											Match match = (Match)enumerator.Current;
											int index = match.Groups[1].Index;
											if (index < num)
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
												num2 = index + 1;
											}
											else
											{
												if (index + match.Groups[1].Length <= num + length)
												{
													continue;
												}
												while (true)
												{
													switch (4)
													{
													case 0:
														continue;
													}
													length2 = index + match.Groups[1].Length - num2 + 1;
													break;
												}
												break;
											}
											continue;
										}
										while (true)
										{
											switch (1)
											{
											case 0:
												break;
											default:
												goto end_IL_01cd;
											}
											continue;
											end_IL_01cd:
											break;
										}
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
							matchCollection = null;
						}
						text = AH.A(43409) + textRange2.get_Characters(num2, length2).Text;
						textRange2 = null;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
			}
			return text.Replace(AH.A(7894), AH.A(14625)).Replace(AH.A(47331), AH.A(14625)).Replace(AH.A(47334), AH.A(14625));
		}
	}
}
