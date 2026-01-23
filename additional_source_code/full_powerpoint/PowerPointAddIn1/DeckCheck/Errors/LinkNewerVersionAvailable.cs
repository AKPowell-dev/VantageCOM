using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class LinkNewerVersionAvailable : BaseError
{
	public LinkNewerVersionAvailable(Slide sld, Shape shp, string strSubtitle)
		: base(ErrorType.LinkNewerVersionAvailable, ((Settings)Main.Analysis.Options).CheckLinks, sld, shp, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(35665);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(35720);
	}
}
