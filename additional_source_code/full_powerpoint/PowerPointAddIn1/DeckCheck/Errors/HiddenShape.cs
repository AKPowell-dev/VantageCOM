using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class HiddenShape : BaseError
{
	public HiddenShape(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
		: base(ErrorType.HiddenShape, Main.Analysis.Options.HiddenShapes, sld, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(33531);
		((BaseError)this).Subtitle = AH.A(33556) + shp.Name + AH.A(33571);
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		base.Shape.Visible = MsoTriState.msoTrue;
	}
}
