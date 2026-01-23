using System;
using System.Collections.Generic;
using System.Drawing;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TurboShapes;

public sealed class RatingBar
{
	public enum BarStyle
	{
		OneColor = 1,
		TwoColors
	}

	public enum ShapeType
	{
		Circle = 1,
		Square,
		Rectangle,
		Star,
		Diamond
	}

	private static readonly int m_A = 11;

	private static readonly int B = 2;

	private static readonly string m_A = AH.A(161277);

	private static readonly string B = AH.A(161176);

	private static readonly string C = AH.A(161316);

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, 2.5f, 5, ShapeType.Circle, BarStyle.OneColor);
		shape.Top = (float)((double)(B.SlideHeight / 2f) - (double)RatingBar.m_A / 2.0);
		shape.Left = (float)((double)(B.SlideWidth / 2f) - (double)RatingBar.m_A / 2.0);
		_ = null;
		Base.LogActivity(AH.A(161213));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, float val, int intSegments, ShapeType type, BarStyle bar)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = null;
		List<string> list = new List<string>();
		float num = 0f;
		bool flag = true;
		MsoAutoShapeType type2 = default(MsoAutoShapeType);
		switch (type)
		{
		case ShapeType.Circle:
			type2 = MsoAutoShapeType.msoShapeOval;
			break;
		case ShapeType.Square:
		case ShapeType.Rectangle:
			type2 = MsoAutoShapeType.msoShapeRectangle;
			break;
		case ShapeType.Star:
			type2 = MsoAutoShapeType.msoShape5pointStar;
			break;
		case ShapeType.Diamond:
			type2 = MsoAutoShapeType.msoShapeDiamond;
			break;
		}
		checked
		{
			int num2 = (int)Math.Floor(val);
			int num3 = num2;
			for (int i = 1; i <= num3; i++)
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape3;
				if (type == ShapeType.Rectangle)
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
					shape3 = sld.Shapes.AddShape(type2, num, 0f, RatingBar.m_A, (float)((double)RatingBar.m_A / 2.0));
				}
				else
				{
					shape3 = sld.Shapes.AddShape(type2, num, 0f, RatingBar.m_A, RatingBar.m_A);
				}
				Microsoft.Office.Interop.PowerPoint.Shape shape4 = shape3;
				shape4.Fill.Visible = MsoTriState.msoTrue;
				shape4.Line.Visible = MsoTriState.msoFalse;
				list.Add(shape4.Name);
				shape4 = null;
				num += (float)(RatingBar.m_A + RatingBar.B);
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				Microsoft.Office.Interop.PowerPoint.Shape shape3 = null;
				if (val - (float)num2 > 0f)
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
					List<string> list2 = new List<string>();
					if (type == ShapeType.Rectangle)
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
						shape3 = sld.Shapes.AddShape(type2, num, 0f, RatingBar.m_A, (float)((double)RatingBar.m_A / 2.0));
					}
					else
					{
						shape3 = sld.Shapes.AddShape(type2, num, 0f, RatingBar.m_A, RatingBar.m_A);
					}
					shape3.Fill.Visible = MsoTriState.msoTrue;
					list2.Add(shape3.Name);
					Microsoft.Office.Interop.PowerPoint.Shape shape5 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, (float)((double)num + (double)RatingBar.m_A * ((double)val - Math.Floor(val))), 0f, RatingBar.m_A, RatingBar.m_A);
					shape5.Fill.Visible = MsoTriState.msoTrue;
					list2.Add(shape5.Name);
					shape5 = null;
					sld.Shapes.Range(list2.ToArray()).MergeShapes(MsoMergeCmd.msoMergeSubtract);
					list2 = null;
					shape3 = sld.Shapes[sld.Shapes.Count];
					shape3.Line.Visible = MsoTriState.msoFalse;
					list.Add(shape3.Name);
				}
				int num4 = num2 + 1;
				Microsoft.Office.Interop.PowerPoint.Shape shape15;
				if (bar == BarStyle.OneColor)
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
					switch (type)
					{
					case ShapeType.Circle:
					{
						int num6 = intSegments;
						for (int k = num4; k <= num6; k++)
						{
							shape3 = A(sld, num, flag);
							Microsoft.Office.Interop.PowerPoint.Shape shape10 = shape3;
							shape10.Fill.Visible = MsoTriState.msoTrue;
							shape10.Line.Visible = MsoTriState.msoFalse;
							if (!flag)
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
								shape10.Adjustments[1] = 0.09f;
							}
							list.Add(shape10.Name);
							shape10 = null;
							num += (float)(RatingBar.m_A + RatingBar.B);
						}
						shape3 = null;
						break;
					}
					case ShapeType.Square:
					{
						int num7 = intSegments;
						for (int l = num4; l <= num7; l++)
						{
							shape3 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeFrame, num, 0f, RatingBar.m_A, RatingBar.m_A);
							Microsoft.Office.Interop.PowerPoint.Shape shape11 = shape3;
							shape11.Fill.Visible = MsoTriState.msoTrue;
							shape11.Line.Visible = MsoTriState.msoFalse;
							shape11.Adjustments[1] = 0.09f;
							list.Add(shape11.Name);
							shape11 = null;
							num += (float)(RatingBar.m_A + RatingBar.B);
						}
						shape3 = null;
						break;
					}
					case ShapeType.Rectangle:
					{
						int num9 = intSegments;
						for (int n = num4; n <= num9; n++)
						{
							shape3 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeFrame, num, 0f, RatingBar.m_A, (float)((double)RatingBar.m_A / 2.0));
							Microsoft.Office.Interop.PowerPoint.Shape shape14 = shape3;
							shape14.Fill.Visible = MsoTriState.msoTrue;
							shape14.Line.Visible = MsoTriState.msoFalse;
							shape14.Adjustments[1] = 0.09f;
							list.Add(shape14.Name);
							shape14 = null;
							num += (float)(RatingBar.m_A + RatingBar.B);
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							break;
						}
						shape3 = null;
						break;
					}
					case ShapeType.Star:
					{
						int num8 = intSegments;
						for (int m = num4; m <= num8; m++)
						{
							List<string> list3 = new List<string>();
							Microsoft.Office.Interop.PowerPoint.Shape shape6 = sld.Shapes.AddShape(MsoAutoShapeType.msoShape5pointStar, num, 0f, 200f, 200f);
							Microsoft.Office.Interop.PowerPoint.Shape shape12 = shape6;
							shape12.Fill.Visible = MsoTriState.msoTrue;
							shape12.Line.Visible = MsoTriState.msoFalse;
							list3.Add(shape12.Name);
							int rGB = shape12.Fill.ForeColor.RGB;
							shape12 = null;
							Microsoft.Office.Interop.PowerPoint.Shape shape8 = sld.Shapes.AddShape(MsoAutoShapeType.msoShape5pointStar, num, 0f, 110f, 110f);
							Microsoft.Office.Interop.PowerPoint.Shape shape13 = shape8;
							shape13.Fill.Visible = MsoTriState.msoTrue;
							shape13.Line.Visible = MsoTriState.msoFalse;
							shape13.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
							shape13.Top = (float)((double)shape6.Top + (double)(shape6.Height - shape8.Height) / 1.81);
							shape13.Left = shape6.Left + (shape6.Width - shape8.Width) / 2f;
							shape13.ZOrder(MsoZOrderCmd.msoBringToFront);
							list3.Add(shape13.Name);
							shape13 = null;
							sld.Shapes.Range(list3.ToArray()).MergeShapes(MsoMergeCmd.msoMergeCombine);
							shape3 = sld.Shapes[sld.Shapes.Count];
							shape3.Fill.ForeColor.RGB = rGB;
							shape3.LockAspectRatio = MsoTriState.msoTrue;
							shape3.Height = RatingBar.m_A;
							list.Add(shape3.Name);
							num += (float)(RatingBar.m_A + RatingBar.B);
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							break;
						}
						shape3 = null;
						break;
					}
					case ShapeType.Diamond:
					{
						int num5 = intSegments;
						for (int j = num4; j <= num5; j++)
						{
							List<string> list3 = new List<string>();
							Microsoft.Office.Interop.PowerPoint.Shape shape6 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeDiamond, num, 0f, 200f, 200f);
							Microsoft.Office.Interop.PowerPoint.Shape shape7 = shape6;
							shape7.Fill.Visible = MsoTriState.msoTrue;
							shape7.Line.Visible = MsoTriState.msoFalse;
							list3.Add(shape7.Name);
							int rGB = shape7.Fill.ForeColor.RGB;
							shape7 = null;
							Microsoft.Office.Interop.PowerPoint.Shape shape8 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeDiamond, num, 0f, 146f, 146f);
							Microsoft.Office.Interop.PowerPoint.Shape shape9 = shape8;
							shape9.Fill.Visible = MsoTriState.msoTrue;
							shape9.Line.Visible = MsoTriState.msoFalse;
							shape9.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);
							shape9.Top = shape6.Top + (shape6.Height - shape8.Height) / 2f;
							shape9.Left = shape6.Left + (shape6.Width - shape8.Width) / 2f;
							shape9.ZOrder(MsoZOrderCmd.msoBringToFront);
							list3.Add(shape9.Name);
							shape9 = null;
							sld.Shapes.Range(list3.ToArray()).MergeShapes(MsoMergeCmd.msoMergeCombine);
							shape3 = sld.Shapes[sld.Shapes.Count];
							shape3.Fill.ForeColor.RGB = rGB;
							shape3.LockAspectRatio = MsoTriState.msoTrue;
							shape3.Height = RatingBar.m_A;
							list.Add(shape3.Name);
							num += (float)(RatingBar.m_A + RatingBar.B);
						}
						shape3 = null;
						break;
					}
					}
					shape15 = Base.MergeShapes(sld, list);
					Microsoft.Office.Interop.PowerPoint.FillFormat fill = shape15.Fill;
					fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Primary);
					fill.BackColor.RGB = fill.ForeColor.RGB;
					fill = null;
				}
				else
				{
					if (list.Count > 0)
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
						if (list.Count == 1)
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
							shape = sld.Shapes[sld.Shapes.Count];
						}
						else
						{
							shape = Base.MergeShapes(sld, list);
						}
						Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = shape.Fill;
						fill2.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Primary);
						fill2.BackColor.RGB = fill2.ForeColor.RGB;
						fill2 = null;
					}
					list = new List<string>();
					int num10 = intSegments;
					for (int num11 = num4; num11 <= num10; num11++)
					{
						if (type == ShapeType.Rectangle)
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
							shape3 = sld.Shapes.AddShape(type2, num, 0f, RatingBar.m_A, (float)((double)RatingBar.m_A / 2.0));
						}
						else
						{
							shape3 = sld.Shapes.AddShape(type2, num, 0f, RatingBar.m_A, RatingBar.m_A);
						}
						Microsoft.Office.Interop.PowerPoint.Shape shape16 = shape3;
						shape16.Fill.Visible = MsoTriState.msoTrue;
						shape16.Line.Visible = MsoTriState.msoFalse;
						list.Add(shape16.Name);
						shape16 = null;
						num += (float)(RatingBar.m_A + RatingBar.B);
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					shape3 = null;
					if (list.Count > 0)
					{
						if (list.Count == 1)
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
							shape2 = sld.Shapes[sld.Shapes.Count];
						}
						else
						{
							shape2 = Base.MergeShapes(sld, list);
						}
						Microsoft.Office.Interop.PowerPoint.FillFormat fill3 = shape2.Fill;
						fill3.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Secondary);
						fill3.BackColor.RGB = fill3.ForeColor.RGB;
						fill3 = null;
						if (shape != null)
						{
							shape.ZOrder(MsoZOrderCmd.msoBringToFront);
							list = new List<string>();
							list.Add(shape.Name);
							list.Add(shape2.Name);
							shape15 = sld.Shapes.Range(list.ToArray()).Group();
						}
						else
						{
							shape15 = shape2;
						}
					}
					else
					{
						shape15 = shape;
					}
				}
				Base.FinalizeShape(shape15, Base.TurboShapeType.RatingBar, val, AH.A(161256));
				Tags tags = shape15.Tags;
				string a = RatingBar.m_A;
				unchecked
				{
					int num12 = (int)type;
					tags.Add(a, num12.ToString());
					shape15.Tags.Add(C, intSegments.ToString());
					Tags tags2 = shape15.Tags;
					string b = B;
					num12 = (int)bar;
					tags2.Add(b, num12.ToString());
					list = null;
					shape = null;
					shape2 = null;
					return shape15;
				}
			}
		}
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, float B, bool C)
	{
		if (C)
		{
			while (true)
			{
				switch (5)
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
					Microsoft.Office.Interop.PowerPoint.Shape shape = shapes.AddShape(MsoAutoShapeType.msoShapeOval, B, 0f, RatingBar.m_A, RatingBar.m_A);
					Microsoft.Office.Interop.PowerPoint.Shape shape2 = checked(shapes.AddShape(MsoAutoShapeType.msoShapeOval, (float)((double)B + 1.5), 1.5f, RatingBar.m_A - 3, RatingBar.m_A - 3));
					string[] index = new string[2] { shape.Name, shape2.Name };
					shapes.Range(index).MergeShapes(MsoMergeCmd.msoMergeSubtract);
					return shapes[shapes.Count];
				}
				}
			}
		}
		return A.Shapes.AddShape(MsoAutoShapeType.msoShapeDonut, B, 0f, RatingBar.m_A, RatingBar.m_A);
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, float val)
	{
		int num = Conversions.ToInteger(shpEdit.Tags[C].ToString());
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopLeft, ref unitX, ref unitY);
		wpfRatingBar wpfRatingBar2 = new wpfRatingBar();
		wpfRatingBar2.EditedShape = shpEdit;
		wpfRatingBar2.CurrentSegments = num;
		wpfRatingBar2.CurrentBarStyle = (BarStyle)Conversions.ToInteger(shpEdit.Tags[B].ToString());
		wpfRatingBar2.BarStyles = new List<BarStyle>();
		wpfRatingBar2.BarStyles.Add(BarStyle.OneColor);
		wpfRatingBar2.BarStyles.Add(BarStyle.TwoColors);
		switch ((ShapeType)Conversions.ToInteger(shpEdit.Tags[RatingBar.m_A].ToString()))
		{
		case ShapeType.Circle:
			wpfRatingBar2.cbxShapes.SelectedIndex = 0;
			break;
		case ShapeType.Square:
			wpfRatingBar2.cbxShapes.SelectedIndex = 1;
			break;
		case ShapeType.Rectangle:
			wpfRatingBar2.cbxShapes.SelectedIndex = 2;
			break;
		case ShapeType.Star:
			wpfRatingBar2.cbxShapes.SelectedIndex = 3;
			break;
		case ShapeType.Diamond:
			wpfRatingBar2.cbxShapes.SelectedIndex = 4;
			break;
		}
		wpfRatingBar2.Slider.Maximum = num;
		wpfRatingBar2.Slider.Value = val;
		wpfRatingBar2.numSegments.Value = num;
		wpfRatingBar2.Top = unitY - wpfRatingBar2.Height;
		wpfRatingBar2.Left = unitX;
		wpfRatingBar2.ShowActivated = false;
		wpfRatingBar2.Show();
		wpfRatingBar2 = null;
	}
}
