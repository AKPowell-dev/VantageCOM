using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class BulletWordCount : BaseTextError
{
	public BulletWordCount(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, int intCount, int intMax)
		: base(ErrorType.BulletWordCount, Main.Analysis.Options.BulletWordCount, sld, shp, para, blnHasFix: false)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(42723);
		((BaseError)this).Subtitle = AH.A(42758) + intCount + AH.A(37464) + intMax;
	}
}
