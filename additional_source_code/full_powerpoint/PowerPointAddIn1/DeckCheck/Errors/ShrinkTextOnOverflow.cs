using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ShrinkTextOnOverflow : BaseError
{
	public ShrinkTextOnOverflow(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
		: base(ErrorType.ShrinkTextOnOverflow, Main.Analysis.Options.ShrinkTextOnOverflow, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(32388);
		((BaseError)this).Subtitle = AH.A(32435);
		((BaseError)this).Tooltip = AH.A(32601);
		((BaseError)this).DisplayText = new List<string>(new string[2]
		{
			AH.A(32897),
			AH.A(32926)
		});
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = base.Shape.TextFrame2;
		if (i == 0)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			textFrame.AutoSize = MsoAutoSize.msoAutoSizeNone;
		}
		else
		{
			textFrame.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
			textFrame.WordWrap = MsoTriState.msoTrue;
		}
		textFrame = null;
	}
}
