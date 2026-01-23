using System.Collections.Generic;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class Ink : BaseError
{
	public Ink(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
		: base(ErrorType.Ink, ((Settings)Main.Analysis.Options).Ink, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		string title;
		if (shp.Type != MsoShapeType.msoInkComment)
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
			title = AH.A(30592);
		}
		else
		{
			title = AH.A(30599);
		}
		((BaseError)this).Title = title;
		((BaseError)this).Subtitle = AH.A(30622);
		((BaseError)this).DisplayText = new List<string>(new string[2]
		{
			AH.A(30794),
			AH.A(30815)
		});
	}

	public override void FixAction(int i)
	{
		NG.A.Application.StartNewUndoEntry();
		if (i == 0)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					base.Shape.Delete();
					return;
				}
			}
		}
		base.Shape.Visible = MsoTriState.msoFalse;
	}
}
