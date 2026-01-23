using System;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ShapeOutsideMargins : BaseError
{
	public ShapeOutsideMargins(Slide sld, Shape shp)
		: base(ErrorType.ShapeOutsideMargins, ((Settings)Main.Analysis.Options).ShapeOutsideMargins, sld, shp, blnHasFix: true)
	{
		//IL_0014: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(34724);
		((BaseError)this).Subtitle = AH.A(34767);
		((BaseError)this).Tooltip = AH.A(34858);
	}

	public override void FixAction()
	{
		NG.A.Application.StartNewUndoEntry();
		CustomLayout customLayout;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation;
		try
		{
			customLayout = base.Slide.CustomLayout;
			presentation = (Microsoft.Office.Interop.PowerPoint.Presentation)base.Slide.Parent;
			PowerPointAddIn1.Template.Settings settings = new PowerPointAddIn1.Template.Settings(presentation);
			double num;
			double num2;
			if (settings.SlideMargins.HasValue)
			{
				while (true)
				{
					switch (2)
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
				PowerPointAddIn1.Template.Settings.Margins value = settings.SlideMargins.Value;
				num = Math.Round(value.Left, 4);
				num2 = Math.Round(presentation.PageSetup.SlideWidth - value.Right, 4);
			}
			else
			{
				Shape bodyPlaceholder = Helpers.GetBodyPlaceholder(presentation);
				num = Math.Round(bodyPlaceholder.Left, 4);
				num2 = Math.Round(bodyPlaceholder.Left + bodyPlaceholder.Width, 4);
				bodyPlaceholder = null;
			}
			settings = null;
			Shape shape = base.Shape;
			if (Math.Round(shape.Left, 4) < num && shape.Left >= 0f)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				shape.Left = (float)num;
			}
			else if (Math.Round(shape.Left + shape.Width, 4) > num2)
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
				if (Math.Round(shape.Left + shape.Width, 4) <= Math.Round(customLayout.Width, 4))
				{
					shape.Left = (float)(num2 - (double)shape.Width);
				}
			}
			shape = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		customLayout = null;
		presentation = null;
	}
}
