using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ColonSpacing : BaseTextError
{
	public ColonSpacing(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, string strFix)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).ColonSpacing, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.ColonSpacing(ref val, strFix);
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		foreach (TextRange2 textRange in ((BaseError)this).TextRanges)
		{
			textRange.Text = ((BaseError)this).ReplacementText[0];
		}
	}
}
