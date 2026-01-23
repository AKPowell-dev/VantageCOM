using System;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class RotatedShape : BaseError
{
	[CompilerGenerated]
	private new float A;

	private float NewRotation
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

	public RotatedShape(Slide sld, Shape shp, float sngOldRotation, float sngNewRotation)
		: base(ErrorType.RotatedShape, Main.Analysis.Options.RotatedShapes, sld, shp, blnHasFix: true)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(34164);
		((BaseError)this).Subtitle = AH.A(33556) + shp.Name + AH.A(34191) + Math.Round(sngOldRotation, 1) + AH.A(34290);
		NewRotation = sngNewRotation;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		base.Shape.Rotation = NewRotation;
	}
}
