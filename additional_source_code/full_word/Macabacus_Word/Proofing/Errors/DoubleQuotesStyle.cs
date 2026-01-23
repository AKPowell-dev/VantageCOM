using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Errors;

public sealed class DoubleQuotesStyle : BaseTextError
{
	public DoubleQuotesStyle(Range rng, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).QuotesStyle, rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		A(strFix);
		((BaseError)this).Subtitle = GenerateSnippet(rng);
	}

	public DoubleQuotesStyle(object shp, TextRange2 rng, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).QuotesStyle, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		A(strFix);
		((BaseError)this).Subtitle = GenerateSnippet(rng);
	}

	private void A(string A)
	{
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.DoubleQuotesStyle(ref val, ((Settings)Main.Analysis.Options).QuotesStyleConvention);
	}

	public override void FixAction(int i)
	{
		string b = ((BaseError)this).ReplacementText[i];
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(26780));
		foreach (TextRange2 textRange in ((BaseError)this).TextRanges)
		{
			A(textRange, b);
		}
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}

	private void A(TextRange2 A, string B)
	{
		checked
		{
			Regex regex;
			if (Operators.CompareString(B, XC.A(24629), TextCompare: false) == 0)
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
				regex = new Regex(XC.A(6379) + Constants.DOUBLE_QUOTE_OPEN + Constants.DOUBLE_QUOTE_CLOSE + XC.A(6382));
				int count = A.get_Paragraphs(-1, -1).Count;
				IEnumerator enumerator = default(IEnumerator);
				for (int i = 1; i <= count; i++)
				{
					TextRange2 textRange = A.get_Paragraphs(i, -1);
					if (regex.Matches(textRange.Text).Count > 0)
					{
						try
						{
							enumerator = textRange.get_Runs(-1, -1).GetEnumerator();
							while (enumerator.MoveNext())
							{
								TextRange2 textRange2 = (TextRange2)enumerator.Current;
								textRange2.Text = regex.Replace(Text.PrintableText(textRange2.Text), B);
							}
						}
						finally
						{
							if (enumerator is IDisposable)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									(enumerator as IDisposable).Dispose();
									break;
								}
							}
						}
					}
					textRange = null;
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
			}
			else
			{
				regex = new Regex(XC.A(24629));
				int count2 = A.get_Paragraphs(-1, -1).Count;
				IEnumerator enumerator2 = default(IEnumerator);
				for (int j = 1; j <= count2; j++)
				{
					TextRange2 textRange3 = A.get_Paragraphs(j, -1);
					try
					{
						enumerator2 = regex.Matches(textRange3.Text).GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Match match = (Match)enumerator2.Current;
							if (match.Index == 0)
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
								textRange3.get_Characters(match.Index + 1, 1).Text = Constants.DOUBLE_QUOTE_OPEN;
							}
							else if (Operators.CompareString(textRange3.get_Characters(match.Index, 1).Text, XC.A(18458), TextCompare: false) == 0)
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
								textRange3.get_Characters(match.Index + 1, 1).Text = Constants.DOUBLE_QUOTE_OPEN;
							}
							else
							{
								textRange3.get_Characters(match.Index + 1, 1).Text = Constants.DOUBLE_QUOTE_CLOSE;
							}
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_0241;
							}
							continue;
							end_IL_0241:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
					textRange3 = null;
				}
			}
			regex = null;
		}
	}
}
