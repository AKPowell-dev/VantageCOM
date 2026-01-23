using System.Windows.Forms;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class PlaceholderLayoutMismatch : BaseError
{
	public PlaceholderLayoutMismatch(Slide sld, Shape shp)
		: base(ErrorType.PlaceholderLayoutMismatch, Main.Analysis.Options.CheckPlaceholderLayoutMismatch, sld, shp, blnHasFix: true)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(40500);
		((BaseError)this).Subtitle = AH.A(40527);
	}

	public override void FixAction()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		application.StartNewUndoEntry();
		base.Shape.Delete();
		Shape shape = base.Slide.Shapes[base.Slide.Shapes.Count];
		float top = shape.Top;
		float left = shape.Left;
		float height = shape.Height;
		float width = shape.Width;
		_ = null;
		application.CommandBars.ExecuteMso(AH.A(40491));
		System.Windows.Forms.Application.DoEvents();
		application.StartNewUndoEntry();
		Shape shape2 = base.Shape;
		shape2.Top = top;
		shape2.Left = left;
		shape2.Height = height;
		shape2.Width = width;
		_ = null;
	}
}
