using System.Collections.Generic;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TurboShapes;

public sealed class ToggleSwitch
{
	public enum SwitchState
	{
		OnGreen = 1,
		OffLeft,
		OffMiddle,
		OffRed
	}

	public enum SwitchStyle
	{
		Solid = 1,
		Frame
	}

	public static readonly string TAG_STYLE = AH.A(161715);

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, SwitchState.OnGreen, SwitchStyle.Solid);
		shape.Width = 36f;
		shape.Top = B.SlideHeight / 2f - shape.Height / 2f;
		shape.Left = B.SlideWidth / 2f - shape.Width / 2f;
		shape = null;
		Base.LogActivity(AH.A(161639));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, SwitchState state, SwitchStyle style)
	{
		List<string> list = new List<string>();
		int color;
		if (state != SwitchState.OnGreen)
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
			if ((uint)(state - 2) > 1u)
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
				color = Base.GetColor(Base.TurboShapeColor.Red);
			}
			else
			{
				color = Base.GetColor(Base.TurboShapeColor.Secondary);
			}
		}
		else
		{
			color = Base.GetColor(Base.TurboShapeColor.Green);
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape3;
		if (style == SwitchStyle.Solid)
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
			float left;
			if (state != SwitchState.OnGreen)
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
				if (state != SwitchState.OffMiddle)
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
					left = 2f;
				}
				else
				{
					left = 10f;
				}
			}
			else
			{
				left = 18f;
			}
			Microsoft.Office.Interop.PowerPoint.Shape shape = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 0f, 0f, 32f, 16f);
			shape.Fill.Visible = MsoTriState.msoTrue;
			shape.Line.Visible = MsoTriState.msoFalse;
			shape.Adjustments[1] = 1f;
			list.Add(shape.Name);
			shape = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape2 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, left, 2f, 12f, 12f);
			shape2.Fill.Visible = MsoTriState.msoTrue;
			shape2.Line.Visible = MsoTriState.msoFalse;
			list.Add(shape2.Name);
			shape2 = null;
			shape3 = Base.CombineShapes(sld, list);
			Microsoft.Office.Interop.PowerPoint.FillFormat fill = shape3.Fill;
			fill.ForeColor.RGB = color;
			fill.BackColor.RGB = color;
			_ = null;
		}
		else
		{
			float left;
			if (state != SwitchState.OnGreen)
			{
				if (state != SwitchState.OffMiddle)
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
					left = 3f;
				}
				else
				{
					left = 11f;
				}
			}
			else
			{
				left = 18.5f;
			}
			Microsoft.Office.Interop.PowerPoint.Shape shape4 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 0f, 0f, 32f, 16f);
			shape4.Fill.Visible = MsoTriState.msoTrue;
			shape4.Adjustments[1] = 1f;
			list.Add(shape4.Name);
			shape4 = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape5 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 1.5f, 1.5f, 29f, 13f);
			shape5.Fill.Visible = MsoTriState.msoTrue;
			shape5.Adjustments[1] = 1f;
			list.Add(shape5.Name);
			shape5 = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape6 = Base.CombineShapes(sld, list);
			shape6.Line.Visible = MsoTriState.msoFalse;
			list = new List<string>();
			list.Add(shape6.Name);
			shape6 = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape7 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, left, 3f, 10f, 10f);
			shape7.Fill.Visible = MsoTriState.msoTrue;
			shape7.Line.Visible = MsoTriState.msoFalse;
			shape7.Adjustments[1] = 1f;
			list.Add(shape7.Name);
			shape7 = null;
			shape3 = Base.MergeShapes(sld, list);
			Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = shape3.Fill;
			fill2.ForeColor.RGB = color;
			fill2.BackColor.RGB = color;
			_ = null;
		}
		Base.FinalizeShape(shape3, Base.TurboShapeType.ToggleSwitch, (float)state, AH.A(161688));
		Tags tags = shape3.Tags;
		string tAG_STYLE = TAG_STYLE;
		int num = (int)style;
		tags.Add(tAG_STYLE, num.ToString());
		list = null;
		return shape3;
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, int val)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		wpfToggleSwitch wpfToggleSwitch2 = new wpfToggleSwitch();
		wpfToggleSwitch2.EditedShape = shpEdit;
		wpfToggleSwitch2.CurrentState = (SwitchState)val;
		wpfToggleSwitch2.SwitchStates = new List<SwitchState>();
		wpfToggleSwitch2.SwitchStates.Add(SwitchState.OnGreen);
		wpfToggleSwitch2.SwitchStates.Add(SwitchState.OffLeft);
		wpfToggleSwitch2.SwitchStates.Add(SwitchState.OffMiddle);
		wpfToggleSwitch2.SwitchStates.Add(SwitchState.OffRed);
		wpfToggleSwitch2.SwitchStyle = (SwitchStyle)Conversions.ToInteger(shpEdit.Tags[TAG_STYLE].ToString());
		wpfToggleSwitch2.Top = unitY - wpfToggleSwitch2.Height;
		wpfToggleSwitch2.Left = unitX;
		wpfToggleSwitch2.ShowActivated = false;
		wpfToggleSwitch2.Show();
		wpfToggleSwitch2 = null;
	}
}
