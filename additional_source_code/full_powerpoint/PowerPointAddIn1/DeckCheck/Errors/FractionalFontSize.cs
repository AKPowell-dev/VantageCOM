using System;
using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class FractionalFontSize : BaseTextError
{
	public FractionalFontSize(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.FractionalFontSize, Main.Analysis.Options.FractionalFontSize, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(31163);
		((BaseError)this).Subtitle = AH.A(31204);
		((BaseError)this).Tooltip = AH.A(31271);
		((BaseError)this).DisplayText = new List<string>(new string[3]
		{
			AH.A(31513),
			AH.A(31562),
			AH.A(31623)
		});
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		if (i == 0)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					using IEnumerator<TextRange2> enumerator = ((BaseError)this).TextRanges.GetEnumerator();
					while (enumerator.MoveNext())
					{
						TextRange2 current = enumerator.Current;
						current.Font.Size = (float)Math.Round(current.Font.Size);
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				}
			}
		}
		if (i == 1)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
				{
					foreach (TextRange2 textRange in ((BaseError)this).TextRanges)
					{
						textRange.Font.Size = (float)Math.Floor(textRange.Font.Size);
					}
					return;
				}
				}
			}
		}
		foreach (TextRange2 textRange2 in ((BaseError)this).TextRanges)
		{
			textRange2.Font.Size = (float)Math.Ceiling(textRange2.Font.Size);
		}
	}
}
