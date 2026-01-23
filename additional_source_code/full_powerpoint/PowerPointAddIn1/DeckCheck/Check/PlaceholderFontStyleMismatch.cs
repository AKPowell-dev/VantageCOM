using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class PlaceholderFontStyleMismatch
{
	public void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, Microsoft.Office.Interop.PowerPoint.Shape placeholder)
	{
		if (placeholder.HasTextFrame != MsoTriState.msoTrue)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
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
			if (shp.HasTextFrame != MsoTriState.msoTrue)
			{
				return;
			}
			Dictionary<int, Font2> dictionary = new Dictionary<int, Font2>();
			try
			{
				enumerator = placeholder.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
				while (enumerator.MoveNext())
				{
					TextRange2 textRange = (TextRange2)enumerator.Current;
					ParagraphFormat2 paragraphFormat = textRange.ParagraphFormat;
					if (!dictionary.ContainsKey(paragraphFormat.IndentLevel))
					{
						dictionary.Add(paragraphFormat.IndentLevel, textRange.Font);
					}
					paragraphFormat = null;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			List<TextRange2> list;
			using (Dictionary<int, Font2>.Enumerator enumerator2 = dictionary.GetEnumerator())
			{
				while (enumerator2.MoveNext())
				{
					KeyValuePair<int, Font2> current = enumerator2.Current;
					list = new List<TextRange2>();
					try
					{
						enumerator3 = shp.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
						while (enumerator3.MoveNext())
						{
							TextRange2 textRange2 = (TextRange2)enumerator3.Current;
							TextRange2 textRange3 = textRange2;
							if (textRange3.ParagraphFormat.IndentLevel == current.Key)
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
								if (Operators.CompareString(textRange3.Font.Name, current.Value.Name, TextCompare: false) != 0)
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
									list.Add(textRange2);
								}
								else
								{
									{
										enumerator4 = textRange2.get_Runs(-1, -1).GetEnumerator();
										try
										{
											while (true)
											{
												if (enumerator4.MoveNext())
												{
													if (((TextRange2)enumerator4.Current).Font.Size == current.Value.Size)
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
														list.Add(textRange2);
														break;
													}
													break;
												}
												while (true)
												{
													switch (7)
													{
													case 0:
														break;
													default:
														goto end_IL_01dc;
													}
													continue;
													end_IL_01dc:
													break;
												}
												break;
											}
										}
										finally
										{
											IDisposable disposable = enumerator4 as IDisposable;
											if (disposable != null)
											{
												disposable.Dispose();
											}
										}
									}
								}
							}
							textRange3 = null;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_020f;
							}
							continue;
							end_IL_020f:
							break;
						}
					}
					finally
					{
						if (enumerator3 is IDisposable)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								(enumerator3 as IDisposable).Dispose();
								break;
							}
						}
					}
					if (list.Count <= 0)
					{
						continue;
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
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.PlaceholderFontStyleMismatch(sld, shp, list, current.Key, current.Value));
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_028b;
					}
					continue;
					end_IL_028b:
					break;
				}
			}
			dictionary = null;
			list = null;
			return;
		}
	}
}
