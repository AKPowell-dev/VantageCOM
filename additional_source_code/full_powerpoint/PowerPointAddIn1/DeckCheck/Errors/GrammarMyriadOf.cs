using System.Collections.Generic;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class GrammarMyriadOf : BaseTextError
{
	public GrammarMyriadOf(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).GrammarMyriadOf, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		string text = ((listRanges.Count != 1) ? (AH.A(42825) + listRanges.Count + AH.A(44714)) : A((List<TextRange2>)((BaseError)this).TextRanges, shp));
		BaseError val = (BaseError)(object)this;
		Errors.GrammarMyriadOf(ref val, text);
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
				current.Text = Regex.Replace(current.Text, AH.A(44689), AH.A(44617), RegexOptions.IgnoreCase);
				_ = null;
			}
			while (true)
			{
				switch (7)
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
					switch (6)
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
