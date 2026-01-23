using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class HyphenatedAdverb : BaseTextError
{
	public HyphenatedAdverb(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.Text, (Severity)3, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		int count = listRanges.Count;
		string text;
		if (count == 1)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			text = A((List<TextRange2>)((BaseError)this).TextRanges, shp);
		}
		else
		{
			text = AH.A(42825) + count + AH.A(44789);
		}
		BaseError val = (BaseError)(object)this;
		Errors.HyphenAdverb(ref val, text);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		using IEnumerator<TextRange2> enumerator = ((BaseError)this).TextRanges.GetEnumerator();
		while (enumerator.MoveNext())
		{
			TextRange2 current = enumerator.Current;
			current.Text = Regex.Replace(current.Text, AH.A(44769), AH.A(44782), RegexOptions.IgnoreCase);
			_ = null;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			return;
		}
	}
}
