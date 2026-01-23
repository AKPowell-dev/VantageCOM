using System.Collections.Generic;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TurboShapes;

public sealed class Arrow
{
	public enum ArrowState
	{
		Up = 1,
		Down,
		Right,
		FourFive,
		OneThreeFive
	}

	public enum ArrowStyle
	{
		Solid = 1,
		Frame,
		Plain
	}

	private static readonly int m_A = 18;

	public static readonly string TAG_STYLE = AH.A(158159);

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, ArrowState.Up, ArrowStyle.Solid);
		shape.Width = 18f;
		shape.Top = B.SlideHeight / 2f - shape.Height / 2f;
		shape.Left = B.SlideWidth / 2f - shape.Width / 2f;
		shape = null;
		Base.LogActivity(AH.A(158115));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, ArrowState state, ArrowStyle style)
	{
		int num = 0;
		List<string> list = new List<string>();
		bool flag = true;
		int rGB = state switch
		{
			ArrowState.Up => Base.GetColor(Base.TurboShapeColor.Green), 
			ArrowState.Down => Base.GetColor(Base.TurboShapeColor.Red), 
			_ => Base.GetColor(Base.TurboShapeColor.Yellow), 
		};
		MsoAutoShapeType type;
		if (state != ArrowState.Down)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (state == ArrowState.Right)
			{
				type = MsoAutoShapeType.msoShapeRightArrow;
			}
			else
			{
				type = MsoAutoShapeType.msoShapeUpArrow;
				if (state == ArrowState.FourFive)
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
					num = 45;
				}
				else if (state == ArrowState.OneThreeFive)
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
					num = 135;
				}
			}
		}
		else
		{
			type = MsoAutoShapeType.msoShapeDownArrow;
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape = sld.Shapes.AddShape(type, 3.5f, 3.5f, 11f, 11f);
		shape.Fill.Visible = MsoTriState.msoTrue;
		shape.Line.Visible = MsoTriState.msoFalse;
		shape.Rotation = num;
		shape.Adjustments[1] = 0.4f;
		_ = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape3;
		if (style != ArrowStyle.Solid)
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
			if (style != ArrowStyle.Frame)
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
				Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape.Duplicate()[1];
				shape2.Top = shape.Top;
				shape2.Left = shape.Left;
				list.Add(shape.Name);
				list.Add(shape2.Name);
				shape3 = Base.MergeShapes(sld, list);
			}
			else
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape4 = A(sld, flag);
				shape4.Fill.Visible = MsoTriState.msoTrue;
				shape4.Line.Visible = MsoTriState.msoFalse;
				if (!flag)
				{
					shape4.Adjustments[1] = 0.08f;
				}
				list.Add(shape4.Name);
				list.Add(shape.Name);
				shape3 = Base.MergeShapes(sld, list);
			}
		}
		else
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape4 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 0f, Arrow.m_A, Arrow.m_A);
			shape4.Fill.Visible = MsoTriState.msoTrue;
			shape4.Line.Visible = MsoTriState.msoFalse;
			list.Add(shape4.Name);
			list.Add(shape.Name);
			shape3 = Base.CombineShapes(sld, list);
		}
		Base.FinalizeShape(shape3, Base.TurboShapeType.Arrow, (float)state, AH.A(158148));
		Tags tags = shape3.Tags;
		string tAG_STYLE = TAG_STYLE;
		int num2 = (int)style;
		tags.Add(tAG_STYLE, num2.ToString());
		list = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape5 = shape3;
		shape5.Fill.ForeColor.RGB = rGB;
		shape5.Fill.BackColor.RGB = shape5.Fill.ForeColor.RGB;
		shape5 = null;
		return shape3;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, bool B)
	{
		if (B)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Microsoft.Office.Interop.PowerPoint.Shapes shapes = A.Shapes;
					Microsoft.Office.Interop.PowerPoint.Shape shape = shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 0f, Arrow.m_A, Arrow.m_A);
					Microsoft.Office.Interop.PowerPoint.Shape shape2 = checked(shapes.AddShape(MsoAutoShapeType.msoShapeOval, 1.5f, 1.5f, Arrow.m_A - 3, Arrow.m_A - 3));
					string[] index = new string[2] { shape.Name, shape2.Name };
					shapes.Range(index).MergeShapes(MsoMergeCmd.msoMergeSubtract);
					return shapes[shapes.Count];
				}
				}
			}
		}
		return A.Shapes.AddShape(MsoAutoShapeType.msoShapeDonut, 0f, 0f, Arrow.m_A, Arrow.m_A);
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, int val)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		wpfArrow wpfArrow2 = new wpfArrow();
		wpfArrow2.EditedShape = shpEdit;
		wpfArrow2.CurrentState = (ArrowState)val;
		wpfArrow2.ArrowStates = new List<ArrowState>();
		wpfArrow2.ArrowStates.Add(ArrowState.Up);
		wpfArrow2.ArrowStates.Add(ArrowState.Down);
		wpfArrow2.ArrowStates.Add(ArrowState.Right);
		wpfArrow2.ArrowStates.Add(ArrowState.FourFive);
		wpfArrow2.ArrowStates.Add(ArrowState.OneThreeFive);
		wpfArrow2.ArrowStyle = (ArrowStyle)Conversions.ToInteger(shpEdit.Tags[TAG_STYLE].ToString());
		wpfArrow2.Top = unitY - wpfArrow2.Height;
		wpfArrow2.Left = unitX;
		wpfArrow2.ShowActivated = false;
		wpfArrow2.Show();
		wpfArrow2 = null;
	}
}
