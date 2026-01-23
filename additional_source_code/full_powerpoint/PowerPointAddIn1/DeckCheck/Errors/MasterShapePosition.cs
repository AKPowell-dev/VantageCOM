using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class MasterShapePosition : BaseError
{
	[CompilerGenerated]
	private new float A;

	[CompilerGenerated]
	private float B;

	private float Top
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

	private float Left
	{
		[CompilerGenerated]
		get
		{
			return B;
		}
		[CompilerGenerated]
		set
		{
			B = value;
		}
	}

	public MasterShapePosition(Slide sld, Shape shp, float sngTop, float sngLeft)
		: base(ErrorType.MasterShapePosition, Main.Analysis.Options.MasterShapePosition, sld, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(33594);
		((BaseError)this).Subtitle = AH.A(33649);
		Top = sngTop;
		Left = sngLeft;
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		Shape shape = base.Shape;
		shape.Top = Top;
		shape.Left = Left;
		_ = null;
	}
}
