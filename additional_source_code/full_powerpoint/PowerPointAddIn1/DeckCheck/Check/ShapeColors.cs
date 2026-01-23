using System;
using System.Collections.Generic;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.Check;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class ShapeColors
{
	public static void FillColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev, bool blnCheckPlaceholderFillColor)
	{
		//IL_02d5: Unknown result type (might be due to invalid IL or missing references)
		//IL_0226: Unknown result type (might be due to invalid IL or missing references)
		//IL_016e: Unknown result type (might be due to invalid IL or missing references)
		checked
		{
			try
			{
				if (shp.HasTable == MsoTriState.msoTrue)
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
						Dictionary<int, List<Microsoft.Office.Interop.PowerPoint.Shape>> dictionary = new Dictionary<int, List<Microsoft.Office.Interop.PowerPoint.Shape>>();
						Table table = shp.Table;
						int count = table.Rows.Count;
						int count2 = table.Columns.Count;
						int num = count;
						for (int i = 1; i <= num; i++)
						{
							int num2 = count2;
							for (int j = 1; j <= num2; j++)
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = table.Cell(i, j).Shape;
								Microsoft.Office.Interop.PowerPoint.FillFormat fill = shape.Fill;
								if (fill.Visible == MsoTriState.msoTrue)
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
									int rGB = fill.ForeColor.RGB;
									if (Color.ColorNotInPalette(rGB, listColors))
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
										if (dictionary.TryGetValue(rGB, out var value))
										{
											value.Add(shape);
											dictionary[rGB] = value;
										}
										else
										{
											value = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
											value.Add(shape);
											dictionary.Add(rGB, value);
										}
										value = null;
									}
								}
								fill = null;
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							if (dictionary.Count > 0)
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
								using Dictionary<int, List<Microsoft.Office.Interop.PowerPoint.Shape>>.Enumerator enumerator = dictionary.GetEnumerator();
								while (enumerator.MoveNext())
								{
									KeyValuePair<int, List<Microsoft.Office.Interop.PowerPoint.Shape>> current = enumerator.Current;
									Main.Analysis.Errors.Add(new TableCellFillColor(sld, shp, current.Key, current.Value, sev));
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_0182;
									}
									continue;
									end_IL_0182:
									break;
								}
							}
							table = null;
							dictionary = null;
							Microsoft.Office.Interop.PowerPoint.Shape shape = null;
							return;
						}
					}
				}
				if (shp.HasSmartArt == MsoTriState.msoTrue)
				{
					IEnumerator<KeyValuePair<int, IList<Microsoft.Office.Core.Shape>>> enumerator2 = default(IEnumerator<KeyValuePair<int, IList<Microsoft.Office.Core.Shape>>>);
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						IDictionary<int, IList<Microsoft.Office.Core.Shape>> dictionary2 = Color.SmartArtFill(shp.SmartArt, listColors);
						if (dictionary2.Count > 0)
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
							try
							{
								enumerator2 = dictionary2.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									KeyValuePair<int, IList<Microsoft.Office.Core.Shape>> current2 = enumerator2.Current;
									Main.Analysis.Errors.Add(new SmartArtFillColor(sld, shp, current2.Key, (List<Microsoft.Office.Core.Shape>)current2.Value, sev));
								}
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										goto end_IL_023c;
									}
									continue;
									end_IL_023c:
									break;
								}
							}
							finally
							{
								if (enumerator2 != null)
								{
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										enumerator2.Dispose();
										break;
									}
								}
							}
						}
						dictionary2 = null;
						return;
					}
				}
				Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = shp.Fill;
				if (fill2.Visible == MsoTriState.msoTrue)
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
					if (!blnCheckPlaceholderFillColor)
					{
						if (shp.Type == MsoShapeType.msoPlaceholder)
						{
							goto IL_02e0;
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
					int rGB = fill2.ForeColor.RGB;
					if (Color.ColorNotInPalette(rGB, listColors))
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
						Main.Analysis.Errors.Add(new FillColor(sld, shp, rGB, sev));
					}
				}
				goto IL_02e0;
				IL_02e0:
				fill2 = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
	}

	public static void BorderColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev)
	{
		//IL_00fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		try
		{
			IEnumerator<KeyValuePair<int, IList<Microsoft.Office.Core.Shape>>> enumerator = default(IEnumerator<KeyValuePair<int, IList<Microsoft.Office.Core.Shape>>>);
			if (shp.HasSmartArt == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						IDictionary<int, IList<Microsoft.Office.Core.Shape>> dictionary = Color.SmartArtBorder(shp.SmartArt, listColors);
						if (dictionary.Count > 0)
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
							try
							{
								enumerator = dictionary.GetEnumerator();
								while (enumerator.MoveNext())
								{
									KeyValuePair<int, IList<Microsoft.Office.Core.Shape>> current = enumerator.Current;
									Main.Analysis.Errors.Add(new SmartArtBorderColor(sld, shp, current.Key, (List<Microsoft.Office.Core.Shape>)current.Value, sev));
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_008b;
									}
									continue;
									end_IL_008b:
									break;
								}
							}
							finally
							{
								if (enumerator != null)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											enumerator.Dispose();
											goto end_IL_009a;
										}
										continue;
										end_IL_009a:
										break;
									}
								}
							}
						}
						dictionary = null;
						return;
					}
					}
				}
			}
			Microsoft.Office.Interop.PowerPoint.LineFormat line = shp.Line;
			if (line.Visible == MsoTriState.msoTrue)
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
				int rGB = line.ForeColor.RGB;
				if (Color.ColorNotInPalette(rGB, listColors))
				{
					Main.Analysis.Errors.Add(new BorderColor(sld, shp, rGB, sev));
				}
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

	public static void TextColor(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<int> listColors, Severity sev, bool blnCheckPlaceholderFontColor)
	{
		//IL_016d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0172: Unknown result type (might be due to invalid IL or missing references)
		//IL_0174: Unknown result type (might be due to invalid IL or missing references)
		//IL_0178: Unknown result type (might be due to invalid IL or missing references)
		//IL_0179: Unknown result type (might be due to invalid IL or missing references)
		//IL_0180: Unknown result type (might be due to invalid IL or missing references)
		//IL_0187: Unknown result type (might be due to invalid IL or missing references)
		//IL_018e: Unknown result type (might be due to invalid IL or missing references)
		//IL_019c: Unknown result type (might be due to invalid IL or missing references)
		//IL_013a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0264: Unknown result type (might be due to invalid IL or missing references)
		checked
		{
			try
			{
				if (shp.HasTable == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
						{
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							IDictionary<int, IList<TextRange2>> d = new Dictionary<int, IList<TextRange2>>();
							IDictionary<int, IList<TextRange2>> e = new Dictionary<int, IList<TextRange2>>();
							IDictionary<int, IList<TextRange2>> f = new Dictionary<int, IList<TextRange2>>();
							IDictionary<int, IList<TextRange2>> g = new Dictionary<int, IList<TextRange2>>();
							Table table = shp.Table;
							int count = table.Rows.Count;
							int count2 = table.Columns.Count;
							int num = count;
							for (int i = 1; i <= num; i++)
							{
								int num2 = count2;
								for (int j = 1; j <= num2; j++)
								{
									Microsoft.Office.Interop.PowerPoint.Shape shape = table.Cell(i, j).Shape;
									if (shape.HasTextFrame == MsoTriState.msoTrue)
									{
										Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shape.TextFrame2;
										if (textFrame.HasText == MsoTriState.msoTrue)
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
											Color.TextFont(textFrame.TextRange, listColors, ref d);
											Color.TextUnderline(textFrame.TextRange, listColors, ref e);
											Color.TextHighlight(textFrame.TextRange, listColors, ref f);
											Color.TextOutline(textFrame.TextRange, listColors, ref g);
										}
										textFrame = null;
									}
									shape = null;
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_0112;
									}
									continue;
									end_IL_0112:
									break;
								}
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									table = null;
									A(sld, shp, sev, d, e, f, g);
									d = null;
									e = null;
									f = null;
									g = null;
									return;
								}
							}
						}
						}
					}
				}
				if (shp.HasSmartArt == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
						{
							SmartArtTextColors val = Color.SmartArtText(shp.SmartArt, listColors);
							A(sld, shp, sev, val.FontColors, val.UnderlineColors, val.HighlightColors, val.OutlineColors);
							val = default(SmartArtTextColors);
							return;
						}
						}
					}
				}
				if (shp.HasTextFrame != MsoTriState.msoTrue)
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
					Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame2 = shp.TextFrame2;
					IDictionary<int, IList<TextRange2>> d2;
					IDictionary<int, IList<TextRange2>> e2;
					IDictionary<int, IList<TextRange2>> f2;
					IDictionary<int, IList<TextRange2>> g2;
					if (textFrame2.HasText == MsoTriState.msoTrue)
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
						d2 = new Dictionary<int, IList<TextRange2>>();
						e2 = new Dictionary<int, IList<TextRange2>>();
						f2 = new Dictionary<int, IList<TextRange2>>();
						g2 = new Dictionary<int, IList<TextRange2>>();
						if (!blnCheckPlaceholderFontColor)
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
							if (shp.Type == MsoShapeType.msoPlaceholder)
							{
								goto IL_022f;
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
						}
						Color.TextFont(textFrame2.TextRange, listColors, ref d2);
						goto IL_022f;
					}
					goto IL_027e;
					IL_027e:
					textFrame2 = null;
					return;
					IL_022f:
					Color.TextUnderline(textFrame2.TextRange, listColors, ref e2);
					Color.TextHighlight(textFrame2.TextRange, listColors, ref f2);
					Color.TextOutline(textFrame2.TextRange, listColors, ref g2);
					A(sld, shp, sev, d2, e2, f2, g2);
					d2 = null;
					e2 = null;
					f2 = null;
					g2 = null;
					goto IL_027e;
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

	private static void A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, Severity C, IDictionary<int, IList<TextRange2>> D, IDictionary<int, IList<TextRange2>> E, IDictionary<int, IList<TextRange2>> F, IDictionary<int, IList<TextRange2>> G)
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_009a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0106: Unknown result type (might be due to invalid IL or missing references)
		//IL_0175: Unknown result type (might be due to invalid IL or missing references)
		using (IEnumerator<KeyValuePair<int, IList<TextRange2>>> enumerator = D.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				KeyValuePair<int, IList<TextRange2>> current = enumerator.Current;
				Main.Analysis.Errors.Add(new FontColor(A, B, current.Key, current.Value, C));
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
		IEnumerator<KeyValuePair<int, IList<TextRange2>>> enumerator2 = default(IEnumerator<KeyValuePair<int, IList<TextRange2>>>);
		try
		{
			enumerator2 = E.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				KeyValuePair<int, IList<TextRange2>> current2 = enumerator2.Current;
				Main.Analysis.Errors.Add(new TextUnderlineColor(A, B, current2.Key, current2.Value, C));
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_00af;
				}
				continue;
				end_IL_00af:
				break;
			}
		}
		finally
		{
			if (enumerator2 != null)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					enumerator2.Dispose();
					break;
				}
			}
		}
		IEnumerator<KeyValuePair<int, IList<TextRange2>>> enumerator3 = default(IEnumerator<KeyValuePair<int, IList<TextRange2>>>);
		try
		{
			enumerator3 = F.GetEnumerator();
			while (enumerator3.MoveNext())
			{
				KeyValuePair<int, IList<TextRange2>> current3 = enumerator3.Current;
				Main.Analysis.Errors.Add(new TextHighlightColor(A, B, current3.Key, current3.Value, C));
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_011c;
				}
				continue;
				end_IL_011c:
				break;
			}
		}
		finally
		{
			if (enumerator3 != null)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					enumerator3.Dispose();
					break;
				}
			}
		}
		foreach (KeyValuePair<int, IList<TextRange2>> item in G)
		{
			Main.Analysis.Errors.Add(new TextOutlineColor(A, B, item.Key, item.Value, C));
		}
	}

	public static void FillTransparency(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		if (shp.Fill.Visible != MsoTriState.msoTrue)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (shp.Fill.Transparency > 0f)
			{
				Main.Analysis.Errors.Add(new FillTransparency(sld, shp));
			}
			return;
		}
	}

	public static void FillGradient(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		if (shp.Fill.Visible != MsoTriState.msoTrue)
		{
			return;
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
			if (shp.Fill.Type != MsoFillType.msoFillGradient)
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				Main.Analysis.Errors.Add(new FillGradient(sld, shp));
				return;
			}
		}
	}
}
