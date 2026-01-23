using System.Drawing;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Proofing.UI;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Proofing.Errors;

public sealed class FillTransparency : BaseError
{
	public FillTransparency(Shape shp)
		: base(ErrorType.FillTransparency, ((Settings)Main.Analysis.Options).FillTransparency, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		BaseError val = (BaseError)(object)this;
		Errors.FillTransparency(ref val, shp.Fill.Transparency);
	}

	public override void FixAction(int i)
	{
		FillFormat fill = default(FillFormat);
		if (base.Shape != null)
		{
			while (true)
			{
				switch (3)
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
			fill = base.Shape.Fill;
		}
		else if (base.InlineShape != null)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			fill = base.InlineShape.Fill;
		}
		Callout.DoNotClose = true;
		Color color = Fixes.ConvertToOpaqueColor(fill.ForeColor.RGB, fill.Transparency);
		Callout.DoNotClose = false;
		Pane.RefocusActiveListBoxItem();
		if (i == 1)
		{
			color = Fixes.FindNearestColor(color);
		}
		UndoRecord undoRecord = PC.A.Application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(27710));
		FillFormat fillFormat = fill;
		fillFormat.Transparency = 0f;
		fillFormat.ForeColor.RGB = ColorTranslator.ToOle(color);
		_ = null;
		undoRecord.EndCustomRecord();
	}
}
