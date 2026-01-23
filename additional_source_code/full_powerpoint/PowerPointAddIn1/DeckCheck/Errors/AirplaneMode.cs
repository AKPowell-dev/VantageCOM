using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class AirplaneMode : BaseError
{
	public AirplaneMode(Slide sld, Shape shp)
		: base(ErrorType.AirplaneMode, Main.Analysis.Options.AirplaneMode, sld, shp, blnHasFix: false)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(32975);
		((BaseError)this).Subtitle = AH.A(33002);
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		PowerPointAddIn1.Shapes.AirplaneMode.Show(base.Shape);
	}
}
