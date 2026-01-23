using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class LinkedPicture : BaseError
{
	public LinkedPicture(Slide sld, Shape shp)
		: base(ErrorType.LinkedPicture, ((Settings)Main.Analysis.Options).LinkedPictures, sld, shp, blnHasFix: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.LinkedPictures(ref val);
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		LinkFormat linkFormat = base.Shape.LinkFormat;
		linkFormat.SourceFullName = "";
		linkFormat.BreakLink();
		_ = null;
	}
}
