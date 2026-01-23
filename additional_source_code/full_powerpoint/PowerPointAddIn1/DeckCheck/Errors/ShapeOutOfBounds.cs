using System;
using A;
using MacabacusMacros.Proofing;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck.Errors;

public sealed class ShapeOutOfBounds : BaseError
{
	public ShapeOutOfBounds(Slide sld, Shape shp, string strSubtitle)
		: base(ErrorType.ShapeOutOfBounds, ((Settings)Main.Analysis.Options).ShapeOutOfBounds, sld, shp, blnHasFix: true)
	{
		//IL_0011: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		((BaseError)this).Title = AH.A(34566);
		((BaseError)this).Subtitle = strSubtitle;
		((BaseError)this).Tooltip = AH.A(34605);
	}

	public override void FixAction()
	{
		CustomLayout customLayout = base.Slide.CustomLayout;
		NG.A.Application.StartNewUndoEntry();
		try
		{
			double num = Math.Round(customLayout.Height, 4);
			double num2 = Math.Round(customLayout.Width, 4);
			Shape shape = base.Shape;
			if (shape.Top < 0f && shape.Top + shape.Height > 0f)
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
				shape.Top = 0f;
			}
			else if (Math.Round(shape.Top + shape.Height, 4) > num)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				if (Math.Round(shape.Top, 4) < num)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
					shape.Top = customLayout.Height - shape.Height;
				}
			}
			if (shape.Left < 0f)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				if (shape.Left + shape.Width > 0f)
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
					shape.Left = 0f;
					goto IL_018b;
				}
			}
			if (Math.Round(shape.Left + shape.Width, 4) > num2)
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
				if (Math.Round(shape.Left, 4) < num2)
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
					shape.Left = customLayout.Width - shape.Width;
				}
			}
			goto IL_018b;
			IL_018b:
			shape = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		customLayout = null;
	}
}
