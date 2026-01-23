using System.Collections.Generic;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.TurboShapes;

public sealed class CheckBox
{
	private enum EG
	{
		A = 1,
		B,
		C
	}

	public enum CheckStyle
	{
		Square = 1,
		Circle
	}

	private static readonly int m_A = 36;

	private static readonly string m_A = AH.A(158653);

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, 1, 1);
		shape.Height = 18f;
		shape.Top = (float)((double)(B.SlideHeight / 2f) - (double)CheckBox.m_A / 2.0);
		shape.Left = (float)((double)(B.SlideWidth / 2f) - (double)CheckBox.m_A / 2.0);
		shape.Fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Primary);
		shape.Fill.BackColor.RGB = shape.Fill.ForeColor.RGB;
		shape = null;
		Base.LogActivity(AH.A(158595));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, int state, int style)
	{
		List<string> list = new List<string>();
		Microsoft.Office.Interop.PowerPoint.Shape shape = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeFrame, 0f, 0f, CheckBox.m_A, CheckBox.m_A);
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
		shape2.Fill.Visible = MsoTriState.msoTrue;
		shape2.Line.Visible = MsoTriState.msoFalse;
		shape2.Adjustments[1] = 0.08f;
		list.Add(shape2.Name);
		shape2 = null;
		if (state == 1)
		{
			FreeformBuilder freeformBuilder = sld.Shapes.BuildFreeform(MsoEditingType.msoEditingCorner, 26f, 7.3f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 26f, 7.3f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 29.2f, 9.2f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 16.5f, 29f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 7f, 22f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 9.2f, 19f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 15.5f, 24f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 26f, 7.3f);
			Microsoft.Office.Interop.PowerPoint.Shape shape3 = freeformBuilder.ConvertToShape();
			_ = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape4 = shape3;
			shape4.Line.Visible = MsoTriState.msoFalse;
			shape4.ZOrder(MsoZOrderCmd.msoSendToBack);
			list.Add(shape4.Name);
			shape4 = null;
		}
		else if (state == 2)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			Microsoft.Office.Interop.PowerPoint.Shape shape5 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeMathMultiply, 0f, 0f, CheckBox.m_A, CheckBox.m_A);
			shape5.Fill.Visible = MsoTriState.msoTrue;
			shape5.Line.Visible = MsoTriState.msoFalse;
			shape5.Adjustments[1] = 0.1f;
			list.Add(shape5.Name);
			shape5 = null;
		}
		else
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape6 = shape.Duplicate()[1];
			shape6.Top = shape.Top;
			shape6.Left = shape.Left;
			list.Add(shape6.Name);
			shape6 = null;
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape7 = Base.MergeShapes(sld, list);
		Base.FinalizeShape(shape7, Base.TurboShapeType.CheckBox, state, AH.A(158636));
		shape7.Tags.Add(CheckBox.m_A, style.ToString());
		list = null;
		return shape7;
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, int val)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		wpfCycleState wpfCycleState2 = new wpfCycleState();
		wpfCycleState2.EditedShape = shpEdit;
		wpfCycleState2.TurboShapeType = Base.TurboShapeType.CheckBox;
		wpfCycleState2.ShapeStates = new List<int>();
		wpfCycleState2.ShapeStates.Add(1);
		wpfCycleState2.ShapeStates.Add(2);
		wpfCycleState2.ShapeStates.Add(3);
		wpfCycleState2.CurrentState = val;
		wpfCycleState2.Top = unitY - wpfCycleState2.Height;
		wpfCycleState2.Left = unitX;
		wpfCycleState2.ShowActivated = false;
		wpfCycleState2.Show();
		wpfCycleState2 = null;
	}
}
