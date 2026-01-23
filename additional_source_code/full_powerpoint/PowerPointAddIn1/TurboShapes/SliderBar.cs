using System.Collections.Generic;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TurboShapes;

public sealed class SliderBar
{
	public enum BarStyle
	{
		FrameOneColor = 1,
		FrameTwoColor,
		SolidTwoColor
	}

	public enum SliderStyle
	{
		Pentagon = 1,
		Rectangle,
		Circle
	}

	public static readonly string TAG_BAR_STYLE = AH.A(161176);

	public static readonly string TAG_SLIDER_STYLE = AH.A(161417);

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, 50, BarStyle.SolidTwoColor, SliderStyle.Circle);
		shape.Height = 18f;
		shape.Top = B.SlideHeight / 2f - shape.Height / 2f;
		shape.Left = B.SlideWidth / 2f - shape.Width / 2f;
		shape = null;
		Base.LogActivity(AH.A(161353));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, int val, BarStyle bar, SliderStyle slider)
	{
		List<string> list = new List<string>();
		Microsoft.Office.Interop.PowerPoint.Shape shape;
		if (bar == BarStyle.SolidTwoColor)
		{
			while (true)
			{
				switch (1)
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
			shape = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0f, 30f, 500f, 30f);
		}
		else
		{
			shape = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeFrame, 0f, 30f, 500f, 30f);
			shape.Adjustments[1] = 0.21f;
		}
		int rGB = ((bar != BarStyle.FrameOneColor) ? Base.GetColor(Base.TurboShapeColor.Secondary) : Base.GetColor(Base.TurboShapeColor.Primary));
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
		shape2.Fill.Visible = MsoTriState.msoTrue;
		shape2.Line.Visible = MsoTriState.msoFalse;
		shape2.LockAspectRatio = MsoTriState.msoTrue;
		shape2.Fill.ForeColor.RGB = rGB;
		shape2.Fill.ForeColor.RGB = rGB;
		list.Add(shape2.Name);
		shape2 = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape3;
		if (slider != SliderStyle.Pentagon)
		{
			if (slider != SliderStyle.Rectangle)
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
				float num = 90f;
				float left = shape.Left + (float)val * shape.Width / 100f - num / 2f;
				shape3 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, left, 0f, num, num);
			}
			else
			{
				float num = 30f;
				float left = shape.Left + (float)val * shape.Width / 100f - num / 2f;
				shape3 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, 0f, num, 90f);
			}
		}
		else
		{
			float num = 30f;
			float left = shape.Left + (float)val * shape.Width / 100f - num / 2f;
			shape3 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeFlowchartOffpageConnector, left, 0f, num, 90f);
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape4 = shape3;
		shape4.Fill.Visible = MsoTriState.msoTrue;
		shape4.Line.Visible = MsoTriState.msoFalse;
		shape4.Fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Primary);
		shape4.Fill.BackColor.RGB = shape4.Fill.ForeColor.RGB;
		list.Add(shape4.Name);
		shape4 = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape5;
		if (bar == BarStyle.FrameOneColor)
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
			shape5 = Base.MergeShapes(sld, list);
		}
		else
		{
			shape5 = sld.Shapes.Range(list.ToArray()).Group();
		}
		list = new List<string>();
		Base.FinalizeShape(shape5, Base.TurboShapeType.SliderBar, val, AH.A(161396));
		Tags tags = shape5.Tags;
		string tAG_BAR_STYLE = TAG_BAR_STYLE;
		int num2 = (int)bar;
		tags.Add(tAG_BAR_STYLE, num2.ToString());
		Tags tags2 = shape5.Tags;
		string tAG_SLIDER_STYLE = TAG_SLIDER_STYLE;
		num2 = (int)slider;
		tags2.Add(tAG_SLIDER_STYLE, num2.ToString());
		list = null;
		shape = null;
		shape3 = null;
		return shape5;
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, int val)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopLeft, ref unitX, ref unitY);
		wpfSliderBar wpfSliderBar2 = new wpfSliderBar();
		wpfSliderBar2.EditedShape = shpEdit;
		wpfSliderBar2.BarStyles = new List<BarStyle>();
		List<BarStyle> barStyles = wpfSliderBar2.BarStyles;
		barStyles.Add(BarStyle.FrameOneColor);
		barStyles.Add(BarStyle.FrameTwoColor);
		barStyles.Add(BarStyle.SolidTwoColor);
		_ = null;
		wpfSliderBar2.CurrentBarStyle = (BarStyle)Conversions.ToInteger(shpEdit.Tags[TAG_BAR_STYLE].ToString());
		wpfSliderBar2.CurrentSliderStyle = (SliderStyle)Conversions.ToInteger(shpEdit.Tags[TAG_SLIDER_STYLE].ToString());
		SliderStyle currentSliderStyle = wpfSliderBar2.CurrentSliderStyle;
		if (currentSliderStyle != SliderStyle.Pentagon)
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
			if (currentSliderStyle == SliderStyle.Circle)
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
				wpfSliderBar2.cbxSlider.SelectedIndex = 0;
			}
			else
			{
				wpfSliderBar2.cbxSlider.SelectedIndex = 2;
			}
		}
		else
		{
			wpfSliderBar2.cbxSlider.SelectedIndex = 1;
		}
		wpfSliderBar2.Slider.Value = val;
		wpfSliderBar2.numValue.Value = val;
		wpfSliderBar2.Top = unitY - wpfSliderBar2.Height;
		wpfSliderBar2.Left = unitX;
		wpfSliderBar2.ShowActivated = false;
		wpfSliderBar2.Show();
		wpfSliderBar2 = null;
	}
}
