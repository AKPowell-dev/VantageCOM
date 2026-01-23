using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SlashSpacingUnbalanced : BaseTextError
{
	public SlashSpacingUnbalanced(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).PunctuationSpacingInconsistent, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		int count = listRanges.Count;
		string text;
		if (count == 1)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			text = A((List<TextRange2>)((BaseError)this).TextRanges, shp);
		}
		else
		{
			text = AH.A(42825) + count + AH.A(46658);
		}
		BaseError val = (BaseError)(object)this;
		Errors.SlashSpacingUnbalanced(ref val, text);
	}

	public override void FixAction(int i)
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Invalid comparison between Unknown and I4
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		if ((int)((Settings)Main.Analysis.Options).SlashSpacingConvention == 1)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					try
					{
						enumerator = ((BaseError)this).TextRanges.GetEnumerator();
						while (enumerator.MoveNext())
						{
							enumerator.Current.Text = AH.A(14622);
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								return;
							}
						}
					}
					finally
					{
						if (enumerator != null)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									enumerator.Dispose();
									goto end_IL_0074;
								}
								continue;
								end_IL_0074:
								break;
							}
						}
					}
				}
			}
		}
		IEnumerator<TextRange2> enumerator2 = default(IEnumerator<TextRange2>);
		try
		{
			enumerator2 = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				enumerator2.Current.Text = AH.A(17773);
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator2 != null)
			{
				while (true)
				{
					switch (6)
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
}
