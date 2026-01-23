using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class Animation : BaseError
{
	public Animation(Slide sld, Shape shp)
		: base(ErrorType.Animation, Main.Analysis.Options.Animation, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(30372);
		((BaseError)this).Subtitle = AH.A(30391);
		((BaseError)this).Tooltip = AH.A(30450);
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		PowerPointAddIn1.Shapes.Animation.Remove(base.Shape);
	}
}
