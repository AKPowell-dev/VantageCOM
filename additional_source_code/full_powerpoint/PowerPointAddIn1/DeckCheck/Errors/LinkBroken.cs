using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class LinkBroken : BaseError
{
	public LinkBroken(Slide sld, Shape shp, string strSubtitle)
		: base(ErrorType.LinkBroken, ((Settings)Main.Analysis.Options).CheckLinks, sld, shp, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(35490);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(35513);
	}
}
