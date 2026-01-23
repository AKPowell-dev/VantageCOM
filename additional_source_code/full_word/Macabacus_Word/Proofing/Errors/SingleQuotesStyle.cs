using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing.Errors;

public sealed class SingleQuotesStyle : BaseTextError
{
	public SingleQuotesStyle(Range rng, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).QuotesStyle, rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		A(strFix);
		((BaseError)this).Subtitle = GenerateSnippet(rng);
	}

	public SingleQuotesStyle(object shp, TextRange2 rng, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).QuotesStyle, RuntimeHelpers.GetObjectValue(shp), rng, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		A(strFix);
		((BaseError)this).Subtitle = GenerateSnippet(rng);
	}

	private void A(string A)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.SingleQuotesStyle(ref val, ((Settings)Main.Analysis.Options).QuotesStyleConvention);
	}

	public override void FixAction(int i)
	{
		string b = ((BaseError)this).ReplacementText[i];
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(26780));
		using (IEnumerator<TextRange2> enumerator = ((BaseError)this).TextRanges.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				TextRange2 current = enumerator.Current;
				A(current, b);
			}
			while (true)
			{
				switch (1)
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
		undoRecord.EndCustomRecord();
		undoRecord = null;
	}

	private void A(TextRange2 A, string B)
	{
		checked
		{
			Regex regex;
			if (Operators.CompareString(B, XC.A(6376), TextCompare: false) == 0)
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
				regex = new Regex(XC.A(6379) + Constants.SINGLE_QUOTE_OPEN + Constants.SINGLE_QUOTE_CLOSE + XC.A(6382));
				int count = A.get_Paragraphs(-1, -1).Count;
				IEnumerator enumerator = default(IEnumerator);
				for (int i = 1; i <= count; i++)
				{
					TextRange2 textRange = A.get_Paragraphs(i, -1);
					if (regex.Matches(textRange.Text).Count > 0)
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
						{
							enumerator = textRange.get_Runs(-1, -1).GetEnumerator();
							try
							{
								while (enumerator.MoveNext())
								{
									TextRange2 textRange2 = (TextRange2)enumerator.Current;
									textRange2.Text = regex.Replace(Text.PrintableText(textRange2.Text), B);
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_00ef;
									}
									continue;
									end_IL_00ef:
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
					textRange = null;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					break;
				}
			}
			else
			{
				regex = new Regex(XC.A(6376));
				int count2 = A.get_Paragraphs(-1, -1).Count;
				for (int j = 1; j <= count2; j++)
				{
					TextRange2 textRange3 = A.get_Paragraphs(j, -1);
					foreach (Match item in regex.Matches(textRange3.Text))
					{
						if (item.Index == 0)
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
							textRange3.get_Characters(item.Index + 1, 1).Text = Constants.SINGLE_QUOTE_OPEN;
						}
						else if (Operators.CompareString(textRange3.get_Characters(item.Index, 1).Text, XC.A(18458), TextCompare: false) == 0)
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
							textRange3.get_Characters(item.Index + 1, 1).Text = Constants.SINGLE_QUOTE_OPEN;
						}
						else
						{
							textRange3.get_Characters(item.Index + 1, 1).Text = Constants.SINGLE_QUOTE_CLOSE;
						}
					}
					textRange3 = null;
				}
			}
			regex = null;
		}
	}
}
