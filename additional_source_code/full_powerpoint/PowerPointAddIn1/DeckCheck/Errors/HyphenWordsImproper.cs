using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class HyphenWordsImproper : BaseTextError
{
	public HyphenWordsImproper(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).HyphenWordsImproper, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		int count = listRanges.Count;
		string text = ((count != 1) ? (AH.A(42825) + count + AH.A(45090)) : A((List<TextRange2>)((BaseError)this).TextRanges, shp));
		BaseError val = (BaseError)(object)this;
		Errors.HyphenWordsImproper(ref val, text);
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
				current.Text = current.Text.Replace(AH.A(17319), AH.A(14625));
				_ = null;
			}
			while (true)
			{
				switch (3)
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
					enumerator.Dispose();
					break;
				}
			}
		}
	}
}
