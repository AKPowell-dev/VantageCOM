using System.Collections.Generic;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TurboShapes;

public sealed class ProgressBar
{
	public enum BarStyle
	{
		OneColor = 1,
		TwoColors
	}

	private static readonly int m_A = 63;

	private static readonly int B = 11;

	private static readonly string m_A = AH.A(161176);

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, ProgressBar.m_A, ProgressBar.B, 50f, BarStyle.OneColor);
		shape.Top = (float)((double)(B.SlideHeight / 2f) - (double)ProgressBar.B / 2.0);
		shape.Left = (float)((double)(B.SlideWidth / 2f) - (double)ProgressBar.m_A / 2.0);
		_ = null;
		Base.LogActivity(AH.A(161104));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, float sngWidth, float sngHeight, float val, BarStyle style)
	{
		List<string> list = new List<string>();
		bool flag = false;
		float width = default(float);
		if (val == 0f)
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
			flag = true;
		}
		else
		{
			width = sngWidth * val / 100f;
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape5;
		if (style == BarStyle.OneColor)
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
			Microsoft.Office.Interop.PowerPoint.Shape shape = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeFrame, 0f, 0f, sngWidth, sngHeight);
			Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
			shape2.Fill.Visible = MsoTriState.msoTrue;
			shape2.Line.Visible = MsoTriState.msoFalse;
			shape2.Adjustments[1] = 0.09f;
			list.Add(shape2.Name);
			shape2 = null;
			if (!flag)
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
				Microsoft.Office.Interop.PowerPoint.Shape shape3 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0f, 0f, width, sngHeight);
				shape3.Fill.Visible = MsoTriState.msoTrue;
				shape3.Line.Visible = MsoTriState.msoFalse;
				list.Add(shape3.Name);
				shape3 = null;
			}
			else
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape4 = shape.Duplicate()[1];
				shape4.Top = shape.Top;
				shape4.Left = shape.Left;
				list.Add(shape4.Name);
			}
			shape5 = Base.MergeShapes(sld, list);
			list = null;
			Microsoft.Office.Interop.PowerPoint.FillFormat fill = shape5.Fill;
			fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Primary);
			fill.BackColor.RGB = fill.ForeColor.RGB;
			fill = null;
		}
		else
		{
			if (val < 100f)
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
				Microsoft.Office.Interop.PowerPoint.Shape shape6 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0f, 0f, sngWidth, sngHeight);
				shape6.Fill.Visible = MsoTriState.msoTrue;
				shape6.Line.Visible = MsoTriState.msoFalse;
				list.Add(shape6.Name);
				Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = shape6.Fill;
				fill2.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Secondary);
				fill2.BackColor.RGB = fill2.ForeColor.RGB;
				fill2 = null;
				shape6 = null;
			}
			if (!flag)
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
				Microsoft.Office.Interop.PowerPoint.Shape shape7 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0f, 0f, width, sngHeight);
				shape7.Fill.Visible = MsoTriState.msoTrue;
				shape7.Line.Visible = MsoTriState.msoFalse;
				list.Add(shape7.Name);
				Microsoft.Office.Interop.PowerPoint.FillFormat fill3 = shape7.Fill;
				fill3.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Primary);
				fill3.BackColor.RGB = fill3.ForeColor.RGB;
				fill3 = null;
				shape7 = null;
				if (list.Count == 1)
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
					shape5 = sld.Shapes[sld.Shapes.Count];
				}
				else
				{
					shape5 = sld.Shapes.Range(list.ToArray()).Group();
				}
			}
			else
			{
				shape5 = sld.Shapes[sld.Shapes.Count];
			}
		}
		Base.FinalizeShape(shape5, Base.TurboShapeType.ProgressBar, val, AH.A(161151));
		Tags tags = shape5.Tags;
		string a = ProgressBar.m_A;
		int num = (int)style;
		tags.Add(a, num.ToString());
		return shape5;
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, int val)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopLeft, ref unitX, ref unitY);
		wpfProgressBar wpfProgressBar2 = new wpfProgressBar();
		wpfProgressBar2.EditedShape = shpEdit;
		wpfProgressBar2.CurrentBarStyle = (BarStyle)Conversions.ToInteger(shpEdit.Tags[ProgressBar.m_A].ToString());
		wpfProgressBar2.BarStyles = new List<BarStyle>();
		wpfProgressBar2.BarStyles.Add(BarStyle.OneColor);
		wpfProgressBar2.BarStyles.Add(BarStyle.TwoColors);
		wpfProgressBar2.Slider.Value = val;
		wpfProgressBar2.numValue.Value = val;
		wpfProgressBar2.Top = unitY - wpfProgressBar2.Height;
		wpfProgressBar2.Left = unitX;
		wpfProgressBar2.ShowActivated = false;
		wpfProgressBar2.Show();
		wpfProgressBar2 = null;
	}
}
