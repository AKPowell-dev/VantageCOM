using System;
using System.Collections;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class SlideTitleCapitalization
{
	[CompilerGenerated]
	private Regex A;

	[CompilerGenerated]
	private bool A;

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

	private bool RequireTitleCase
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

	public SlideTitleCapitalization(Conventions conv)
	{
		RegexObj = new Regex(Constants.REGEX_WORDS);
		RequireTitleCase = conv.TitleCaseTitlesCount >= conv.SentenceCaseTitlesCount;
	}

	public void Check(Slide sld)
	{
		try
		{
			if (sld.Shapes.HasTitle != MsoTriState.msoTrue)
			{
				return;
			}
			IEnumerator enumerator = default(IEnumerator);
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
				Microsoft.Office.Interop.PowerPoint.Shape title = sld.Shapes.Title;
				if (title.HasTextFrame == MsoTriState.msoTrue)
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
					if (title.TextFrame2.HasText == MsoTriState.msoTrue)
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
						string text = Strings.Trim(title.TextFrame2.TextRange.Text);
						if (RegexObj.Matches(text).Count > 1)
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
							if (Text.CountCapsInTitle(text) == 0)
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
								if (RequireTitleCase)
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
									string text2 = new CultureInfo(CultureInfo.CurrentCulture.Name, useUserOverride: false).TextInfo.ToTitleCase(text);
									string[] array = new string[13]
									{
										AH.A(7902),
										AH.A(14823),
										AH.A(14828),
										AH.A(14835),
										AH.A(14842),
										AH.A(14849),
										AH.A(14854),
										AH.A(14861),
										AH.A(14868),
										AH.A(14873),
										AH.A(14878),
										AH.A(14883),
										AH.A(14892)
									};
									foreach (string text3 in array)
									{
										text2 = text2.Replace(AH.A(14625) + text3 + AH.A(14625), AH.A(14625) + text3.ToLower() + AH.A(14625));
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
									Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.SlideTitleCapitalization(sld, sld.Shapes.Title, title.TextFrame2.TextRange.get_Paragraphs(1, -1).get_Characters(1, text.Length), text2));
								}
							}
							else if (!RequireTitleCase)
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
								string text2 = text;
								MatchCollection matchCollection = Regex.Matches(text, AH.A(14897));
								{
									enumerator = matchCollection.GetEnumerator();
									try
									{
										while (enumerator.MoveNext())
										{
											Group obj = ((Match)enumerator.Current).Groups[1];
											text2 = Strings.Left(text2, obj.Index) + obj.Value.ToLower() + Strings.Right(text2, checked(text2.Length - obj.Index - obj.Length));
											obj = null;
										}
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												goto end_IL_0332;
											}
											continue;
											end_IL_0332:
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
								Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.SlideTitleCapitalization(sld, sld.Shapes.Title, title.TextFrame2.TextRange.get_Paragraphs(1, -1).get_Characters(1, text.Length), text2));
							}
						}
					}
					else
					{
						Main.Analysis.Errors.Add(new SlideTitleMissing(sld, sld.Shapes.Title));
					}
				}
				title = null;
				return;
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
