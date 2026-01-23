using System.Collections.Generic;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TurboShapes;

public sealed class Thermometer
{
	public enum MeterStyle
	{
		OneColor = 1,
		TwoColor
	}

	public static readonly string TAG_STYLE = AH.A(161945);

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, 25, MeterStyle.OneColor);
		shape.Height = 36f;
		shape.Top = B.SlideHeight / 2f - shape.Height / 2f;
		shape.Left = B.SlideWidth / 2f - shape.Width / 2f;
		shape = null;
		Base.LogActivity(AH.A(161877));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, int val, MeterStyle style)
	{
		List<string> list = new List<string>();
		float left = 8f;
		Microsoft.Office.Interop.PowerPoint.Shape shape7;
		Microsoft.Office.Interop.PowerPoint.Shape shape3;
		Microsoft.Office.Interop.PowerPoint.Shape shape6;
		if (style == MeterStyle.TwoColor)
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
			Microsoft.Office.Interop.PowerPoint.Shape shape = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 80f, 32f, 32f);
			shape.Fill.Visible = MsoTriState.msoTrue;
			shape.Line.Visible = MsoTriState.msoFalse;
			list.Add(shape.Name);
			shape = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape2 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, 0f, 16f, 95f);
			shape2.Fill.Visible = MsoTriState.msoTrue;
			shape2.Line.Visible = MsoTriState.msoFalse;
			list.Add(shape2.Name);
			shape2 = null;
			shape3 = Base.MergeShapes(sld, list);
			shape3.Fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Secondary);
			shape3.Fill.BackColor.RGB = shape3.Fill.ForeColor.RGB;
			list = new List<string>();
			Microsoft.Office.Interop.PowerPoint.Shape shape4 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 80f, 32f, 32f);
			shape4.Fill.Visible = MsoTriState.msoTrue;
			shape4.Line.Visible = MsoTriState.msoFalse;
			list.Add(shape4.Name);
			shape4 = null;
			float num = 15f + 80f * (float)val / 100f;
			float top = 95f - num;
			Microsoft.Office.Interop.PowerPoint.Shape shape5 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, 16f, num);
			shape5.Fill.Visible = MsoTriState.msoTrue;
			shape5.Line.Visible = MsoTriState.msoFalse;
			list.Add(shape5.Name);
			shape5 = null;
			shape6 = Base.MergeShapes(sld, list);
			shape6.Fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Red);
			shape6.Fill.BackColor.RGB = shape6.Fill.ForeColor.RGB;
			list = new List<string>();
			list.Add(shape3.Name);
			list.Add(shape6.Name);
			shape7 = sld.Shapes.Range(list.ToArray()).Group();
		}
		else
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape8 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeFrame, left, 0f, 16f, 95f);
			shape8.Fill.Visible = MsoTriState.msoTrue;
			shape8.Line.Visible = MsoTriState.msoFalse;
			shape8.Adjustments[1] = 0.1f;
			list.Add(shape8.Name);
			shape8 = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape9 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 80f, 32f, 32f);
			shape9.Fill.Visible = MsoTriState.msoTrue;
			shape9.Line.Visible = MsoTriState.msoFalse;
			list.Add(shape9.Name);
			shape9 = null;
			float num = 15f + 80f * (float)val / 100f;
			float top = 95f - num;
			Microsoft.Office.Interop.PowerPoint.Shape shape10 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, 16f, num);
			shape10.Fill.Visible = MsoTriState.msoTrue;
			shape10.Line.Visible = MsoTriState.msoFalse;
			list.Add(shape10.Name);
			shape10 = null;
			shape7 = Base.MergeShapes(sld, list);
			shape7.Fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Red);
			shape7.Fill.BackColor.RGB = shape7.Fill.ForeColor.RGB;
		}
		Base.FinalizeShape(shape7, Base.TurboShapeType.Thermometer, val, AH.A(161922));
		Tags tags = shape7.Tags;
		string tAG_STYLE = TAG_STYLE;
		int num2 = (int)style;
		tags.Add(tAG_STYLE, num2.ToString());
		list = null;
		shape3 = null;
		shape6 = null;
		return shape7;
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, int val)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		wpfThermometer wpfThermometer2 = new wpfThermometer();
		wpfThermometer2.EditedShape = shpEdit;
		wpfThermometer2.MeterStyles = new List<MeterStyle>();
		wpfThermometer2.MeterStyles.Add(MeterStyle.OneColor);
		wpfThermometer2.MeterStyles.Add(MeterStyle.TwoColor);
		wpfThermometer2.CurrentStyle = (MeterStyle)Conversions.ToInteger(shpEdit.Tags[TAG_STYLE].ToString());
		wpfThermometer2.Slider.Value = val;
		wpfThermometer2.numValue.Value = val;
		wpfThermometer2.Top = unitY - wpfThermometer2.Height;
		wpfThermometer2.Left = unitX;
		wpfThermometer2.ShowActivated = false;
		wpfThermometer2.Show();
		wpfThermometer2 = null;
	}
}
