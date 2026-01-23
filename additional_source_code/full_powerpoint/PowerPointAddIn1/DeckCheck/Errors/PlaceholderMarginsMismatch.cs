using System;
using System.Collections;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class PlaceholderMarginsMismatch : BaseError
{
	public PlaceholderMarginsMismatch(Slide sld, Shape shp, string strSubtitle)
		: base(ErrorType.PlaceholderMarginsMismatch, Main.Analysis.Options.CheckPlaceholderMarginMismatch, sld, shp, blnHasFix: true)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(40689);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(40758);
	}

	public override void FixAction()
	{
		IEnumerator enumerator = base.Slide.CustomLayout.Shapes.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				Shape shape = (Shape)enumerator.Current;
				if (!Helpers.IsShapeMatch(shape, base.Shape))
				{
					continue;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					NG.A.Application.StartNewUndoEntry();
					TextFrame2 textFrame = shape.TextFrame2;
					TextFrame2 textFrame2 = base.Shape.TextFrame2;
					textFrame2.MarginTop = textFrame.MarginTop;
					textFrame2.MarginRight = textFrame.MarginRight;
					textFrame2.MarginBottom = textFrame.MarginBottom;
					textFrame2.MarginLeft = textFrame.MarginLeft;
					_ = null;
					textFrame = null;
					return;
				}
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}
}
