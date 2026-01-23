using System.Drawing;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.Fix;

namespace Macabacus_Word.Proofing.Errors;

public sealed class NonconformingFontColor : BaseColorError
{
	public NonconformingFontColor(object shp, int intColor, Severity sev)
		: base(ErrorType.ColorPaletteFont, sev, RuntimeHelpers.GetObjectValue(shp), intColor)
	{
		//IL_0003: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = XC.A(32519);
		((BaseError)this).Subtitle = XC.A(32570);
	}

	public override void FixAction(Color color)
	{
		Macabacus_Word.Proofing.Fix.Colors.ReplaceColor(this, color);
	}
}
