using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class CrookedLine : BaseError
{
	[CompilerGenerated]
	private new bool A;

	private bool Horizontal
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public CrookedLine(Slide sld, Shape shp, bool blnHorizontal)
		: base(ErrorType.CrookedLine, Main.Analysis.Options.RotatedShapes, sld, shp, blnHasFix: true, blnCanFixMultiple: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(33787);
		if (blnHorizontal)
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
			((BaseError)this).Subtitle = AH.A(33828);
		}
		else
		{
			((BaseError)this).Subtitle = AH.A(33998);
		}
		Horizontal = blnHorizontal;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		if (Horizontal)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					base.Shape.Height = 0f;
					return;
				}
			}
		}
		base.Shape.Width = 0f;
	}
}
