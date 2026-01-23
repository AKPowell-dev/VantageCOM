using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class HyphenMissing : BaseTextError
{
	public HyphenMissing(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.Text, (Severity)3, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		int count = listRanges.Count;
		string text = ((count != 1) ? (AH.A(42825) + count + AH.A(44938)) : A((List<TextRange2>)((BaseError)this).TextRanges, shp));
		BaseError val = (BaseError)(object)this;
		Errors.HyphenMissing(ref val, text);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				TextRange2 current = enumerator.Current;
				current.Text = Regex.Replace(current.Text, AH.A(44906), AH.A(44927), RegexOptions.IgnoreCase);
				_ = null;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}
}
