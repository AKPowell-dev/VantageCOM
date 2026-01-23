using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class SlashSpacingInconsistent : BaseTextError
{
	public SlashSpacingInconsistent(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).PunctuationSpacingInconsistent, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_0031: Unknown result type (might be due to invalid IL or missing references)
		//IL_0036: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.SlashSpacingInconsistent(ref val, ((Settings)Main.Analysis.Options).SlashSpacingConvention);
	}

	public override void FixAction(int i)
	{
		//IL_0021: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Invalid comparison between Unknown and I4
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
								switch (4)
								{
								case 0:
									break;
								default:
									enumerator.Dispose();
									goto end_IL_007e;
								}
								continue;
								end_IL_007e:
								break;
							}
						}
					}
				}
			}
		}
		using IEnumerator<TextRange2> enumerator2 = ((BaseError)this).TextRanges.GetEnumerator();
		while (enumerator2.MoveNext())
		{
			enumerator2.Current.Text = AH.A(17773);
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
}
