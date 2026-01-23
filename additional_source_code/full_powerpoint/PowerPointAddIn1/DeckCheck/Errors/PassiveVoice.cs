using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class PassiveVoice : BaseTextError
{
	public PassiveVoice(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, List<TextRange2> listRanges, string strSentence)
		: base(ErrorType.Text, ((Settings)Main.Analysis.Options).PassiveVoice, sld, shp, listRanges, blnHasFix: false)
	{
		//IL_0010: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.PassiveVoice(ref val, AH.A(43409) + Text.PrintableText(strSentence));
	}
}
