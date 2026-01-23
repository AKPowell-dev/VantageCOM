using System.Collections.Generic;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.TurboShapes;

public sealed class NoticeIcon
{
	private enum FG
	{
		A = 1,
		B,
		C,
		D
	}

	private static readonly int m_A = 18;

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, 1);
		shape.Top = (float)((double)(B.SlideHeight / 2f) - (double)NoticeIcon.m_A / 2.0);
		shape.Left = (float)((double)(B.SlideWidth / 2f) - (double)NoticeIcon.m_A / 2.0);
		_ = null;
		Base.LogActivity(AH.A(161036));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, int state)
	{
		List<string> list = new List<string>();
		Microsoft.Office.Interop.PowerPoint.Shape shape = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 0f, NoticeIcon.m_A, NoticeIcon.m_A);
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
		shape2.Fill.Visible = MsoTriState.msoTrue;
		shape2.Line.Visible = MsoTriState.msoFalse;
		list.Add(shape2.Name);
		shape2 = null;
		int color;
		switch ((FG)state)
		{
		case FG.A:
		{
			color = Base.GetColor(Base.TurboShapeColor.Green);
			FreeformBuilder freeformBuilder = sld.Shapes.BuildFreeform(MsoEditingType.msoEditingCorner, 13.5f, 4.5f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 13.5f, 4.5f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 14.9f, 5.8f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 7.4f, 14.3f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 2.5f, 9.6f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 3.88f, 8.05f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 7.4f, 11.43f);
			freeformBuilder.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingCorner, 13.5f, 4.5f);
			Microsoft.Office.Interop.PowerPoint.Shape shape8 = freeformBuilder.ConvertToShape();
			_ = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape9 = shape8;
			shape9.Line.Visible = MsoTriState.msoFalse;
			list.Add(shape9.Name);
			shape9 = null;
			break;
		}
		case FG.B:
		{
			color = Base.GetColor(Base.TurboShapeColor.Yellow);
			Microsoft.Office.Interop.PowerPoint.Shape shape6 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0f, 3.3f, 2.4f, 7.3f);
			shape6.Fill.Visible = MsoTriState.msoTrue;
			shape6.Line.Visible = MsoTriState.msoFalse;
			shape6.Left = shape.Left + (shape.Width - shape6.Width) / 2f;
			list.Add(shape6.Name);
			shape6 = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape7 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 12.5f, 2.4f, 2.4f);
			shape7.Fill.Visible = MsoTriState.msoTrue;
			shape7.Line.Visible = MsoTriState.msoFalse;
			shape7.Left = shape.Left + (shape.Width - shape7.Width) / 2f;
			list.Add(shape7.Name);
			shape7 = null;
			break;
		}
		case FG.C:
		{
			color = Base.GetColor(Base.TurboShapeColor.Red);
			Microsoft.Office.Interop.PowerPoint.Shape shape5 = checked(sld.Shapes.AddShape(MsoAutoShapeType.msoShapeMathMultiply, 0f, 0f, NoticeIcon.m_A - 2, NoticeIcon.m_A - 2));
			shape5.Fill.Visible = MsoTriState.msoTrue;
			shape5.Line.Visible = MsoTriState.msoFalse;
			shape5.Adjustments[1] = 0.12f;
			shape5.Left = shape.Left + (shape.Width - shape5.Width) / 2f;
			shape5.Top = shape.Top + (shape.Height - shape5.Height) / 2f;
			list.Add(shape5.Name);
			shape5 = null;
			break;
		}
		default:
		{
			color = Base.GetColor(Base.TurboShapeColor.Blue);
			Microsoft.Office.Interop.PowerPoint.Shape shape3 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 3.3f, 2.4f, 2.4f);
			shape3.Fill.Visible = MsoTriState.msoTrue;
			shape3.Line.Visible = MsoTriState.msoFalse;
			shape3.Left = shape.Left + (shape.Width - shape3.Width) / 2f;
			list.Add(shape3.Name);
			shape3 = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape4 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0f, 7.6f, 2.4f, 7.3f);
			shape4.Fill.Visible = MsoTriState.msoTrue;
			shape4.Line.Visible = MsoTriState.msoFalse;
			shape4.Left = shape.Left + (shape.Width - shape4.Width) / 2f;
			list.Add(shape4.Name);
			shape4 = null;
			break;
		}
		}
		sld.Shapes.Range(list.ToArray()).MergeShapes(MsoMergeCmd.msoMergeCombine);
		Microsoft.Office.Interop.PowerPoint.Shape shape10 = sld.Shapes[sld.Shapes.Count];
		Base.FinalizeShape(shape10, Base.TurboShapeType.NoticeIcon, state, AH.A(161081));
		list = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape11 = shape10;
		shape11.Fill.ForeColor.RGB = color;
		shape11.Fill.BackColor.RGB = shape11.Fill.ForeColor.RGB;
		shape11 = null;
		return shape10;
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, int val)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		wpfCycleState wpfCycleState2 = new wpfCycleState();
		wpfCycleState2.EditedShape = shpEdit;
		wpfCycleState2.TurboShapeType = Base.TurboShapeType.NoticeIcon;
		wpfCycleState2.ShapeStates = new List<int>();
		wpfCycleState2.ShapeStates.Add(1);
		wpfCycleState2.ShapeStates.Add(2);
		wpfCycleState2.ShapeStates.Add(3);
		wpfCycleState2.ShapeStates.Add(4);
		wpfCycleState2.CurrentState = val;
		wpfCycleState2.Top = unitY - wpfCycleState2.Height;
		wpfCycleState2.Left = unitX;
		wpfCycleState2.ShowActivated = false;
		wpfCycleState2.Show();
		wpfCycleState2 = null;
	}
}
