using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class FootnoteExplanationMissing
{
	[CompilerGenerated]
	private Regex A;

	[CompilerGenerated]
	private Dictionary<Slide, List<int>> A;

	private Regex RegexObj
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	private Dictionary<Slide, List<int>> FootnoteNumbers
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

	public FootnoteExplanationMissing(Dictionary<Slide, List<int>> dict)
	{
		RegexObj = new Regex(AH.A(15141));
		FootnoteNumbers = dict;
	}

	public void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, ref List<int> listFoundFootnotes)
	{
		List<int> value = null;
		if (!FootnoteNumbers.TryGetValue(sld, out value))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			if (value.Count > 0)
			{
				{
					enumerator = para.get_Runs(-1, -1).GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							TextRange2 textRange = (TextRange2)enumerator.Current;
							if (textRange.Font.Superscript != MsoTriState.msoTrue)
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
							try
							{
								enumerator2 = RegexObj.Matches(textRange.Text).GetEnumerator();
								while (enumerator2.MoveNext())
								{
									string value2 = ((Match)enumerator2.Current).Groups[1].Value;
									string[] array = ((!value2.Contains(AH.A(12717))) ? new string[1] { value2 } : value2.Split(','));
									string[] array2 = array;
									foreach (string text in array2)
									{
										listFoundFootnotes.Add(Conversions.ToInteger(text));
										if (value.Contains(Conversions.ToInteger(text)))
										{
											continue;
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
										Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.FootnoteExplanationMissing(sld, shp, textRange, text));
									}
								}
								while (true)
								{
									switch (5)
									{
									case 0:
										break;
									default:
										goto end_IL_0165;
									}
									continue;
									end_IL_0165:
									break;
								}
							}
							finally
							{
								if (enumerator2 is IDisposable)
								{
									while (true)
									{
										switch (6)
										{
										case 0:
											continue;
										}
										(enumerator2 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_019c;
							}
							continue;
							end_IL_019c:
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
			}
			value = null;
			return;
		}
	}
}
