using System.Drawing;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck;

public class BaseColorError : BaseError
{
	public BaseColorError(ErrorType errType, Severity sev, Slide sld, Shape shp, int intColor)
		: base(errType, sev, sld, shp, blnHasFix: false)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).NonconformingColor = ColorTranslator.FromOle(intColor);
		((BaseError)this).HasColorFix = true;
	}
}
