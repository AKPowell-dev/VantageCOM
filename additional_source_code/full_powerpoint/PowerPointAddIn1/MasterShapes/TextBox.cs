using System;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.MasterShapes;

public sealed class TextBox
{
	public static readonly string TEXTBOX_NAME = AH.A(151028);

	public static void Insert()
	{
		Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		float num = 250f;
		float num2 = 250f;
		if (Base.A(application, B: true))
		{
			application = null;
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
		try
		{
			activePresentation = application.ActivePresentation;
			shape = Shape(activePresentation);
			if (shape == null)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				shape = Helpers.GetBodyPlaceholder(activePresentation);
			}
			if (shape != null)
			{
				try
				{
					Selection selection = application.ActiveWindow.Selection;
					if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape2 = selection.ShapeRange[1];
						num = shape2.Width;
						num2 = shape2.Height;
						_ = null;
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				PageSetup pageSetup = activePresentation.PageSetup;
				float top = pageSetup.SlideHeight / 2f - num2 / 2f;
				float left = pageSetup.SlideWidth / 2f - num / 2f;
				_ = null;
				application.StartNewUndoEntry();
				Microsoft.Office.Interop.PowerPoint.Shape shape3 = application.ActiveWindow.Selection.SlideRange[1].Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, num, num2);
				shape3.Select();
				Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shape3.TextFrame2;
				textFrame.AutoSize = MsoAutoSize.msoAutoSizeNone;
				textFrame.TextRange.InsertAfter(AH.A(150895));
				textFrame.TextRange.Select();
				_ = null;
				shape.PickUp();
				shape3.Apply();
				Base.A(AH.A(150912));
			}
			else
			{
				Forms.WarningMessage(AH.A(150943));
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		application = null;
		activePresentation = null;
		shape = null;
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Shape(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		Microsoft.Office.Interop.PowerPoint.Shape result = null;
		try
		{
			result = pres.Designs[1].SlideMaster.Shapes[TEXTBOX_NAME];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
