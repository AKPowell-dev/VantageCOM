using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class PlaceholderBulletMismatch
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
				switch (7)
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
					Dictionary<int, BulletFormat2> dictionary = new Dictionary<int, BulletFormat2>();
					enumerator = placeholder.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							ParagraphFormat2 paragraphFormat = ((TextRange2)enumerator.Current).ParagraphFormat;
							if (A(paragraphFormat.Bullet))
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
								if (!dictionary.ContainsKey(paragraphFormat.IndentLevel))
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
									dictionary.Add(paragraphFormat.IndentLevel, paragraphFormat.Bullet);
								}
							}
							paragraphFormat = null;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_00d2;
							}
							continue;
							end_IL_00d2:
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
					using (Dictionary<int, BulletFormat2>.Enumerator enumerator2 = dictionary.GetEnumerator())
					{
						while (enumerator2.MoveNext())
						{
							KeyValuePair<int, BulletFormat2> current = enumerator2.Current;
							list = new List<TextRange2>();
							try
							{
								ParagraphFormat2 paragraphFormat2;
								for (enumerator3 = shp.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator(); enumerator3.MoveNext(); paragraphFormat2 = null)
								{
									TextRange2 textRange = (TextRange2)enumerator3.Current;
									paragraphFormat2 = textRange.ParagraphFormat;
									if (paragraphFormat2.IndentLevel != current.Key || !A(paragraphFormat2.Bullet))
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
										break;
									}
									BulletFormat2 value = current.Value;
									BulletFormat2 bullet = paragraphFormat2.Bullet;
									if (bullet.Type == MsoBulletType.msoBulletUnnumbered)
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
										if (value.Type == MsoBulletType.msoBulletUnnumbered)
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
											if (bullet.Character != value.Character)
											{
												goto IL_02f2;
											}
										}
									}
									if (bullet.Type == MsoBulletType.msoBulletUnnumbered && value.Type == MsoBulletType.msoBulletUnnumbered && bullet.UseTextFont == MsoTriState.msoFalse)
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
										if (Operators.CompareString(bullet.Font.Name, value.Font.Name, TextCompare: false) != 0)
										{
											goto IL_02f2;
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
									if (bullet.UseTextColor == MsoTriState.msoFalse)
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
										if (bullet.Font.Fill.ForeColor.RGB != value.Font.Fill.ForeColor.RGB)
										{
											goto IL_02f2;
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												continue;
											}
											break;
										}
									}
									if (bullet.RelativeSize != value.RelativeSize)
									{
										goto IL_02f2;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										break;
									}
									if (bullet.Type == MsoBulletType.msoBulletNumbered)
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
										if (value.Type == MsoBulletType.msoBulletNumbered)
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
											if (bullet.Style != value.Style)
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
												goto IL_02f2;
											}
										}
									}
									goto IL_02fa;
									IL_02f2:
									list.Add(textRange);
									goto IL_02fa;
									IL_02fa:
									bullet = null;
									value = null;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_030e;
									}
									continue;
									end_IL_030e:
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
								switch (6)
								{
								case 0:
									continue;
								}
								break;
							}
							Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.PlaceholderBulletMismatch(sld, shp, list, current.Key, current.Value));
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0384;
							}
							continue;
							end_IL_0384:
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

	private bool A(BulletFormat2 A)
	{
		if (A.Type != MsoBulletType.msoBulletNumbered)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return A.Type == MsoBulletType.msoBulletUnnumbered;
				}
			}
		}
		return true;
	}
}
