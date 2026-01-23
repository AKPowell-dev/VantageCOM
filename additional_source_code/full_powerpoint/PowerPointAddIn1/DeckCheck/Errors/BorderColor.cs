using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class BorderColor : BaseColorError
{
	public BorderColor(Slide sld, Shape shp, int intColor, Severity sev)
		: base(ErrorType.ColorPaletteBorder, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(24858);
		((BaseError)this).Subtitle = AH.A(24883);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.Shape.Line.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
