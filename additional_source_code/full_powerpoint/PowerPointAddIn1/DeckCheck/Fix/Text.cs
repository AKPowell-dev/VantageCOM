using System.Globalization;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using MacabacusMacros.Proofing.CorporateDictionary;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Fix;

public sealed class Text
{
	public static void ReplaceText(BaseError err, int i)
	{
		NG.A.Application.StartNewUndoEntry();
		if ((object)((object)err).GetType() == typeof(DoubleQuotesStyle))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (Operators.CompareString(((BaseError)err).ReplacementText[i], AH.A(15132), TextCompare: false) != 0)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (((BaseError)err).TextRanges[0].Start == 1)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									((BaseError)err).TextRanges[0].Text = Constants.DOUBLE_QUOTE_OPEN;
									return;
								}
							}
						}
						if (Operators.CompareString(err.Shape.TextFrame2.TextRange.get_Characters(checked(((BaseError)err).TextRanges[0].Start - 1), 1).Text, AH.A(14625), TextCompare: false) == 0)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									((BaseError)err).TextRanges[0].Text = Constants.DOUBLE_QUOTE_OPEN;
									return;
								}
							}
						}
						((BaseError)err).TextRanges[0].Text = Constants.DOUBLE_QUOTE_CLOSE;
						return;
					}
				}
			}
		}
		((BaseError)err).TextRanges[0].Text = ((BaseError)err).ReplacementText[i];
	}

	internal static string A(string A, string B, Rule C)
	{
		//IL_0087: Unknown result type (might be due to invalid IL or missing references)
		//IL_008c: Unknown result type (might be due to invalid IL or missing references)
		//IL_008e: Unknown result type (might be due to invalid IL or missing references)
		//IL_008f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0091: Unknown result type (might be due to invalid IL or missing references)
		//IL_0093: Invalid comparison between Unknown and I4
		//IL_00a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a4: Invalid comparison between Unknown and I4
		if (C.IsRegex && C.ReplaceWith.Count == 1)
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
			if (Regex.IsMatch(C.ReplaceWith[0], AH.A(47510)))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						B = C.SearchRegex.Replace(A, C.ReplaceWith[0]);
						return B;
					}
				}
			}
		}
		RuleReason reason = C.Reason;
		if (reason - 2 > 1)
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
			if ((int)reason != 6 && !C.ReplaceMatchCase)
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
				checked
				{
					if (Operators.CompareString(A, A.ToLower(), TextCompare: false) == 0)
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
						B = B.ToLower();
					}
					else if (Operators.CompareString(A, A.ToUpper(), TextCompare: false) == 0)
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
						B = B.ToUpper();
					}
					else if (Operators.CompareString(Strings.Left(A, 1), Strings.Left(A, 1).ToUpper(), TextCompare: false) == 0)
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
						if (Operators.CompareString(A, CultureInfo.CurrentCulture.TextInfo.ToTitleCase(A), TextCompare: false) == 0 && Text.A(A) > 1)
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
							B = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(B);
						}
						else if (Operators.CompareString(Strings.Right(A, A.Length - 1), Strings.Right(A, A.Length - 1).ToLower(), TextCompare: false) == 0)
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
							B = Strings.Left(B, 1).ToUpper() + Strings.Right(B, B.Length - 1).ToLower();
						}
					}
				}
			}
		}
		return B;
	}

	private static int A(string A)
	{
		return new Regex(AH.A(47519)).Matches(A).Count;
	}
}
