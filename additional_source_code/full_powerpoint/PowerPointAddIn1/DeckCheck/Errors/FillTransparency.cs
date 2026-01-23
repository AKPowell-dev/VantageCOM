using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.UI;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class FillTransparency : BaseError
{
	public FillTransparency(Slide sld, Shape shp)
		: base(ErrorType.FillTransparency, ((Settings)Main.Analysis.Options).FillTransparency, sld, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.FillTransparency(ref val, shp.Fill.Transparency);
	}

	public override void FixAction(int i)
	{
		Callout.DoNotClose = true;
		FillFormat fill = base.Shape.Fill;
		Color color = Fixes.ConvertToOpaqueColor(fill.ForeColor.RGB, fill.Transparency);
		fill = null;
		Callout.DoNotClose = false;
		PowerPointAddIn1.DeckCheck.UI.Pane.F();
		if (i == 1)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			color = Fixes.FindNearestColor(color);
		}
		NG.A.Application.StartNewUndoEntry();
		FillFormat fill2 = base.Shape.Fill;
		fill2.Transparency = 0f;
		fill2.ForeColor.RGB = ColorTranslator.ToOle(color);
		_ = null;
	}
}
