using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class DashSpacingUnbalanced : BaseTextError
{
	public DashSpacingUnbalanced(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).PunctuationSpacingInconsistent, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		int count = listRanges.Count;
		string text = ((count != 1) ? (AH.A(42825) + count + AH.A(42838)) : A((List<TextRange2>)((BaseError)this).TextRanges, shp));
		BaseError val = (BaseError)(object)this;
		Errors.DashSpacingUnbalanced(ref val, text);
	}

	public override void FixAction(int i)
	{
		//IL_001f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Invalid comparison between Unknown and I4
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		if ((int)((Settings)Main.Analysis.Options).DashSpacingConvention == 1)
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
							enumerator.Current.Text = AH.A(15092);
						}
						while (true)
						{
							switch (7)
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
								switch (7)
								{
								case 0:
									break;
								default:
									enumerator.Dispose();
									goto end_IL_007c;
								}
								continue;
								end_IL_007c:
								break;
							}
						}
					}
				}
			}
		}
		foreach (TextRange2 textRange in ((BaseError)this).TextRanges)
		{
			textRange.Text = AH.A(15125);
		}
	}
}
