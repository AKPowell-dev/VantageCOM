using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class FillColor : BaseColorError
{
	public FillColor(Slide sld, Shape shp, int intColor, Severity sev)
		: base(ErrorType.ColorPaletteFill, sev, sld, shp, intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(26041);
		((BaseError)this).Subtitle = AH.A(26062);
	}

	public override void FixAction(Color color)
	{
		NG.A.Application.StartNewUndoEntry();
		base.Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(color);
	}
}
