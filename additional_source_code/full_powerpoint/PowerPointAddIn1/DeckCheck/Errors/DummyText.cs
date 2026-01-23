using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class DummyText : BaseTextError
{
	public DummyText(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).DummyText, sld, shp, listRanges, blnHasFix: false)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		string text;
		if (listRanges.Count == 1)
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
			text = A((List<TextRange2>)((BaseError)this).TextRanges, shp);
		}
		else
		{
			text = AH.A(43409);
			using (List<TextRange2>.Enumerator enumerator = listRanges.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					TextRange2 current = enumerator.Current;
					text = text + current.Text + AH.A(14258);
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_00a0;
					}
					continue;
					end_IL_00a0:
					break;
				}
			}
			text = Strings.Left(text, checked(text.Length - 2));
		}
		BaseError val = (BaseError)(object)this;
		Errors.DummyText(ref val, text);
	}
}
