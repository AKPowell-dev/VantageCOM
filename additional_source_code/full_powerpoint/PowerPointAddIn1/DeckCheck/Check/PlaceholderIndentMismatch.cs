using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class PlaceholderIndentMismatch
{
	public void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, Microsoft.Office.Interop.PowerPoint.Shape placeholder)
	{
		try
		{
			if (placeholder.HasTextFrame != MsoTriState.msoTrue)
			{
				return;
			}
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator enumerator3 = default(IEnumerator);
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
					Dictionary<int, Tuple<float, float, float>> dictionary = new Dictionary<int, Tuple<float, float, float>>();
					enumerator = placeholder.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							ParagraphFormat2 paragraphFormat = ((TextRange2)enumerator.Current).ParagraphFormat;
							if (!dictionary.ContainsKey(paragraphFormat.IndentLevel))
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
								dictionary.Add(paragraphFormat.IndentLevel, new Tuple<float, float, float>(paragraphFormat.FirstLineIndent, paragraphFormat.LeftIndent, paragraphFormat.RightIndent));
							}
							paragraphFormat = null;
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_00c1;
							}
							continue;
							end_IL_00c1:
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
					List<TextRange2> list;
					using (Dictionary<int, Tuple<float, float, float>>.Enumerator enumerator2 = dictionary.GetEnumerator())
					{
						while (enumerator2.MoveNext())
						{
							KeyValuePair<int, Tuple<float, float, float>> current = enumerator2.Current;
							list = new List<TextRange2>();
							{
								enumerator3 = shp.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
								try
								{
									ParagraphFormat2 paragraphFormat2;
									for (; enumerator3.MoveNext(); paragraphFormat2 = null)
									{
										TextRange2 textRange = (TextRange2)enumerator3.Current;
										paragraphFormat2 = textRange.ParagraphFormat;
										if (paragraphFormat2.IndentLevel != current.Key)
										{
											continue;
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												continue;
											}
											break;
										}
										if (paragraphFormat2.FirstLineIndent == current.Value.Item1)
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
											if (paragraphFormat2.LeftIndent == current.Value.Item2)
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
												if (paragraphFormat2.RightIndent == current.Value.Item3)
												{
													continue;
												}
												while (true)
												{
													switch (7)
													{
													case 0:
														continue;
													}
													break;
												}
											}
										}
										list.Add(textRange);
									}
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											goto end_IL_01e0;
										}
										continue;
										end_IL_01e0:
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
							if (list.Count <= 0)
							{
								continue;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								break;
							}
							Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.PlaceholderIndentMismatch(sld, shp, list, current.Key, current.Value.Item1, current.Value.Item2, current.Value.Item3));
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0275;
							}
							continue;
							end_IL_0275:
							break;
						}
					}
					dictionary = null;
					list = null;
					return;
				}
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
