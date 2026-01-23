using System.Collections.Generic;
using System.Globalization;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class HyphenWordsInconsistent : BaseTextError
{
	public HyphenWordsInconsistent(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 rng, List<string> listLabels, List<string> listFixes)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).HyphenWordsInconsistent, sld, shp, rng, blnHasFix: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.HyphenWordsInconsistent(ref val, A((List<TextRange2>)((BaseError)this).TextRanges, shp), listLabels, listFixes);
	}

	public override void FixAction(int i)
	{
		string text = ((BaseError)this).ReplacementText[i];
		NG.A.Application.StartNewUndoEntry();
		IEnumerator<TextRange2> enumerator = default(IEnumerator<TextRange2>);
		try
		{
			enumerator = ((BaseError)this).TextRanges.GetEnumerator();
			while (enumerator.MoveNext())
			{
				TextRange2 current = enumerator.Current;
				if (Operators.CompareString(Strings.Left(current.Text, 1), Strings.Left(current.Text.ToUpper(), 1), TextCompare: false) == 0)
				{
					while (true)
					{
						switch (4)
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
					text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(text);
				}
				current.Text = text;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (3)
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
