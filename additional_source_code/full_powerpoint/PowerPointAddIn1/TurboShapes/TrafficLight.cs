using System.Collections.Generic;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TurboShapes;

public sealed class TrafficLight
{
	public enum TrafficLightState
	{
		Red = 1,
		Yellow,
		Green,
		All,
		None
	}

	public enum Style
	{
		Solid = 1,
		Donut
	}

	public static readonly string TAG_LIGHT_STYLE = AH.A(161834);

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, 1, 1);
		shape.Height = 36f;
		shape.Top = B.SlideHeight / 2f - shape.Height / 2f;
		shape.Left = B.SlideWidth / 2f - shape.Width / 2f;
		shape = null;
		Base.LogActivity(AH.A(161758));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, int state, int sty)
	{
		List<string> list = new List<string>();
		int num;
		if (state != 1)
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
			num = ((state == 4) ? 1 : 0);
		}
		else
		{
			num = 1;
		}
		bool flag = (byte)num != 0;
		int num2;
		if (state != 2)
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
			num2 = ((state == 4) ? 1 : 0);
		}
		else
		{
			num2 = 1;
		}
		bool flag2 = (byte)num2 != 0;
		int num3;
		if (state != 3)
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
			num3 = ((state == 4) ? 1 : 0);
		}
		else
		{
			num3 = 1;
		}
		bool flag3 = (byte)num3 != 0;
		if (sty == 1)
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
			Microsoft.Office.Interop.PowerPoint.Shape shape = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 0f, 16f, 16f);
			shape.Fill.Visible = MsoTriState.msoTrue;
			shape.Line.Visible = MsoTriState.msoFalse;
			shape.LockAspectRatio = MsoTriState.msoTrue;
			int color = Base.GetColor(flag ? Base.TurboShapeColor.Red : Base.TurboShapeColor.Secondary);
			shape.Fill.ForeColor.RGB = color;
			shape.Fill.BackColor.RGB = color;
			list.Add(shape.Name);
			shape = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape2 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 20f, 16f, 16f);
			shape2.Fill.Visible = MsoTriState.msoTrue;
			shape2.Line.Visible = MsoTriState.msoFalse;
			shape2.LockAspectRatio = MsoTriState.msoTrue;
			int clr;
			if (!flag2)
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
				clr = 2;
			}
			else
			{
				clr = 4;
			}
			color = Base.GetColor((Base.TurboShapeColor)clr);
			shape2.Fill.ForeColor.RGB = color;
			shape2.Fill.BackColor.RGB = color;
			list.Add(shape2.Name);
			shape2 = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape3 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 40f, 16f, 16f);
			shape3.Fill.Visible = MsoTriState.msoTrue;
			shape3.Line.Visible = MsoTriState.msoFalse;
			shape3.LockAspectRatio = MsoTriState.msoTrue;
			color = Base.GetColor(flag3 ? Base.TurboShapeColor.Green : Base.TurboShapeColor.Secondary);
			shape3.Fill.ForeColor.RGB = color;
			shape3.Fill.BackColor.RGB = color;
			list.Add(shape3.Name);
			shape3 = null;
		}
		else
		{
			int num4;
			if (!flag)
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
				num4 = 18;
			}
			else
			{
				num4 = 9;
			}
			MsoAutoShapeType type = (MsoAutoShapeType)num4;
			Microsoft.Office.Interop.PowerPoint.Shape shape4 = sld.Shapes.AddShape(type, 0f, 0f, 16f, 16f);
			shape4.Fill.Visible = MsoTriState.msoTrue;
			shape4.Line.Visible = MsoTriState.msoFalse;
			shape4.LockAspectRatio = MsoTriState.msoTrue;
			shape4.Fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Red);
			shape4.Fill.BackColor.RGB = shape4.Fill.ForeColor.RGB;
			if (!flag)
			{
				shape4.Adjustments[1] = 0.1f;
			}
			list.Add(shape4.Name);
			shape4 = null;
			int num5;
			if (!flag2)
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
				num5 = 18;
			}
			else
			{
				num5 = 9;
			}
			type = (MsoAutoShapeType)num5;
			Microsoft.Office.Interop.PowerPoint.Shape shape5 = sld.Shapes.AddShape(type, 0f, 20f, 16f, 16f);
			shape5.Fill.Visible = MsoTriState.msoTrue;
			shape5.Line.Visible = MsoTriState.msoFalse;
			shape5.LockAspectRatio = MsoTriState.msoTrue;
			shape5.Fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Yellow);
			shape5.Fill.BackColor.RGB = shape5.Fill.ForeColor.RGB;
			if (!flag2)
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
				shape5.Adjustments[1] = 0.1f;
			}
			list.Add(shape5.Name);
			shape5 = null;
			int num6;
			if (!flag3)
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
				num6 = 18;
			}
			else
			{
				num6 = 9;
			}
			type = (MsoAutoShapeType)num6;
			Microsoft.Office.Interop.PowerPoint.Shape shape6 = sld.Shapes.AddShape(type, 0f, 40f, 16f, 16f);
			shape6.Fill.Visible = MsoTriState.msoTrue;
			shape6.Line.Visible = MsoTriState.msoFalse;
			shape6.LockAspectRatio = MsoTriState.msoTrue;
			shape6.Fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Green);
			shape6.Fill.BackColor.RGB = shape6.Fill.ForeColor.RGB;
			if (!flag3)
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
				shape6.Adjustments[1] = 0.1f;
			}
			list.Add(shape6.Name);
			shape6 = null;
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape7 = sld.Shapes.Range(list.ToArray()).Group();
		Base.FinalizeShape(shape7, Base.TurboShapeType.TrafficLight, state, AH.A(161807));
		shape7.Tags.Add(TAG_LIGHT_STYLE, sty.ToString());
		list = null;
		return shape7;
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, int val)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		wpfTrafficLight wpfTrafficLight2 = new wpfTrafficLight();
		wpfTrafficLight2.EditedShape = shpEdit;
		wpfTrafficLight2.LightStyle = (Style)Conversions.ToInteger(shpEdit.Tags[TAG_LIGHT_STYLE].ToString());
		wpfTrafficLight2.ShapeStates = new List<TrafficLightState>();
		wpfTrafficLight2.ShapeStates.Add(TrafficLightState.Red);
		wpfTrafficLight2.ShapeStates.Add(TrafficLightState.Yellow);
		wpfTrafficLight2.ShapeStates.Add(TrafficLightState.Green);
		wpfTrafficLight2.ShapeStates.Add(TrafficLightState.All);
		wpfTrafficLight2.ShapeStates.Add(TrafficLightState.None);
		wpfTrafficLight2.CurrentState = (TrafficLightState)val;
		wpfTrafficLight2.Top = unitY - wpfTrafficLight2.Height;
		wpfTrafficLight2.Left = unitX;
		wpfTrafficLight2.ShowActivated = false;
		wpfTrafficLight2.Show();
		wpfTrafficLight2 = null;
	}
}
