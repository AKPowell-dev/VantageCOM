using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public abstract class MinMaxFontSize : BaseTextError
{
	internal new int A;

	public MinMaxFontSize(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, int limit)
		: base(ErrorType.MaxMinFontSize, Main.Analysis.Options.MinMaxFontSize, sld, shp, listRanges, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		A = limit;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		using IEnumerator<TextRange2> enumerator = ((BaseError)this).TextRanges.GetEnumerator();
		while (enumerator.MoveNext())
		{
			enumerator.Current.Font.Size = A;
		}
		while (true)
		{
			switch (4)
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
