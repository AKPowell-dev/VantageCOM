using System.Collections.Generic;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TurboShapes;

public sealed class HarveyBall
{
	public enum BallStyle
	{
		Traditional = 1,
		Modern,
		TraditionalTwoColor,
		Ring
	}

	private static readonly int m_A = 36;

	private static readonly string m_A = AH.A(160997);

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, 50f, BallStyle.Traditional);
		shape.Top = (float)((double)(B.SlideHeight / 2f) - (double)HarveyBall.m_A / 2.0);
		shape.Left = (float)((double)(B.SlideWidth / 2f) - (double)HarveyBall.m_A / 2.0);
		_ = null;
		Base.LogActivity(AH.A(160929));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, float val, BallStyle style)
	{
		List<string> list = new List<string>();
		bool flag = true;
		if (val == 60f)
		{
			val = 59.95f;
		}
		else if (val == 80f)
		{
			val = 79.95f;
		}
		float val2 = -90f;
		float num = val;
		float val3 = default(float);
		if (num != 0f)
		{
			if (num == 25f)
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
				val3 = 0f;
			}
			else if (num == 50f)
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
				val3 = 90f;
			}
			else if (num == 75f)
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
				val3 = 180f;
			}
			else if (num == 100f)
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
				val3 = -90f;
			}
			else if (!(val < 25f))
			{
				val3 = ((!(val < 75f)) ? (-90f - 90f * ((100f - val) / 25f)) : (90f * (val / 25f - 1f)));
			}
			else
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
				val3 = -90f * (1f - val / 25f);
			}
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape9;
		Tags tags;
		string a;
		Microsoft.Office.Interop.PowerPoint.Shape shape3 = default(Microsoft.Office.Interop.PowerPoint.Shape);
		Microsoft.Office.Interop.PowerPoint.Shape shape11;
		checked
		{
			switch (style)
			{
			case BallStyle.Traditional:
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
					shape3 = A(sld, flag);
					Microsoft.Office.Interop.PowerPoint.Shape shape14 = shape3;
					shape14.Fill.Visible = MsoTriState.msoTrue;
					if (!flag)
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
						shape14.Adjustments[1] = 0.07f;
					}
					list.Add(shape14.Name);
					shape14 = null;
					if (val == 0f)
					{
						shape11 = shape3.Duplicate()[1];
						shape11.Top = shape3.Top;
						shape11.Left = shape3.Left;
						list.Add(shape11.Name);
					}
				}
				if (val > 0f)
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
					if (val == 100f)
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
						Microsoft.Office.Interop.PowerPoint.Shape shape15 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 0f, HarveyBall.m_A, HarveyBall.m_A);
						shape15.Fill.Visible = MsoTriState.msoTrue;
						list.Add(shape15.Name);
						shape15 = null;
					}
					else
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape16 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapePie, 0f, 0f, HarveyBall.m_A, HarveyBall.m_A);
						shape16.Fill.Visible = MsoTriState.msoTrue;
						shape16.Adjustments[1] = val2;
						shape16.Adjustments[2] = val3;
						list.Add(shape16.Name);
						shape16 = null;
					}
				}
				if (list.Count > 1)
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
					shape9 = Base.MergeShapes(sld, list);
				}
				else
				{
					shape9 = sld.Shapes[sld.Shapes.Count];
				}
				shape9.Line.Visible = MsoTriState.msoFalse;
				Microsoft.Office.Interop.PowerPoint.FillFormat fill4 = shape9.Fill;
				fill4.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Primary);
				fill4.BackColor.RGB = fill4.ForeColor.RGB;
				fill4 = null;
				break;
			}
			case BallStyle.Modern:
			{
				shape3 = A(sld, flag);
				Microsoft.Office.Interop.PowerPoint.Shape shape10 = shape3;
				shape10.Fill.Visible = MsoTriState.msoTrue;
				if (!flag)
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
					shape10.Adjustments[1] = 0.06f;
				}
				list.Add(shape10.Name);
				shape10 = null;
				if (val == 0f)
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
					shape11 = shape3.Duplicate()[1];
					shape11.Top = shape3.Top;
					shape11.Left = shape3.Left;
					list.Add(shape11.Name);
				}
				else if (val == 100f)
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
					Microsoft.Office.Interop.PowerPoint.Shape shape12 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 5f, 5f, HarveyBall.m_A - 10, HarveyBall.m_A - 10);
					shape12.Fill.Visible = MsoTriState.msoTrue;
					list.Add(shape12.Name);
					shape12 = null;
				}
				else
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape13 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapePie, 5f, 5f, HarveyBall.m_A - 10, HarveyBall.m_A - 10);
					shape13.Fill.Visible = MsoTriState.msoTrue;
					shape13.Adjustments[1] = val2;
					shape13.Adjustments[2] = val3;
					list.Add(shape13.Name);
					shape13 = null;
				}
				shape9 = Base.MergeShapes(sld, list);
				shape9.Line.Visible = MsoTriState.msoFalse;
				Microsoft.Office.Interop.PowerPoint.FillFormat fill3 = shape9.Fill;
				fill3.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Primary);
				fill3.BackColor.RGB = fill3.ForeColor.RGB;
				fill3 = null;
				break;
			}
			case BallStyle.TraditionalTwoColor:
				if (val < 100f)
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
					shape3 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 0f, HarveyBall.m_A, HarveyBall.m_A);
					Microsoft.Office.Interop.PowerPoint.Shape shape17 = shape3;
					shape17.Fill.Visible = MsoTriState.msoTrue;
					shape17.Line.Visible = MsoTriState.msoFalse;
					Microsoft.Office.Interop.PowerPoint.FillFormat fill5 = shape17.Fill;
					fill5.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Secondary);
					fill5.BackColor.RGB = fill5.ForeColor.RGB;
					fill5 = null;
					list.Add(shape17.Name);
					shape17 = null;
				}
				if (val > 0f)
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
					Microsoft.Office.Interop.PowerPoint.Shape shape7 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapePie, 0f, 0f, HarveyBall.m_A, HarveyBall.m_A);
					Microsoft.Office.Interop.PowerPoint.Shape shape18 = shape7;
					shape18.Fill.Visible = MsoTriState.msoTrue;
					shape18.Line.Visible = MsoTriState.msoFalse;
					Microsoft.Office.Interop.PowerPoint.FillFormat fill6 = shape18.Fill;
					fill6.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Primary);
					fill6.BackColor.RGB = fill6.ForeColor.RGB;
					fill6 = null;
					shape18.Adjustments[1] = val2;
					shape18.Adjustments[2] = val3;
					_ = null;
					List<string> list2 = new List<string>();
					Microsoft.Office.Interop.PowerPoint.Shape shape19 = shape7.Duplicate()[1];
					shape19.Top = shape7.Top;
					shape19.Left = shape7.Left;
					list2.Add(shape7.Name);
					list2.Add(shape19.Name);
					Microsoft.Office.Interop.PowerPoint.Shape shape20 = Base.MergeShapes(sld, list2);
					list.Add(shape20.Name);
					shape20 = null;
					list2 = null;
				}
				if (list.Count > 1)
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
					shape9 = sld.Shapes.Range(list.ToArray()).Group();
				}
				else
				{
					shape9 = sld.Shapes[sld.Shapes.Count];
				}
				break;
			default:
			{
				if (val < 100f)
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 0f, HarveyBall.m_A, HarveyBall.m_A);
					shape.Fill.Visible = MsoTriState.msoTrue;
					list.Add(shape.Name);
					shape = null;
					Microsoft.Office.Interop.PowerPoint.Shape shape2 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 6f, 6f, HarveyBall.m_A - 12, HarveyBall.m_A - 12);
					shape2.Fill.Visible = MsoTriState.msoTrue;
					list.Add(shape2.Name);
					shape2 = null;
					sld.Shapes.Range(list.ToArray()).MergeShapes(MsoMergeCmd.msoMergeCombine);
					shape3 = sld.Shapes[sld.Shapes.Count];
					Microsoft.Office.Interop.PowerPoint.Shape shape4 = shape3;
					shape4.Line.Visible = MsoTriState.msoFalse;
					Microsoft.Office.Interop.PowerPoint.FillFormat fill = shape4.Fill;
					fill.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Secondary);
					fill.BackColor.RGB = fill.ForeColor.RGB;
					fill = null;
					_ = null;
				}
				Microsoft.Office.Interop.PowerPoint.Shape shape7 = default(Microsoft.Office.Interop.PowerPoint.Shape);
				if (val > 0f)
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
					list = new List<string>();
					Microsoft.Office.Interop.PowerPoint.Shape shape5 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapePie, 0f, 0f, HarveyBall.m_A, HarveyBall.m_A);
					shape5.Fill.Visible = MsoTriState.msoTrue;
					shape5.Adjustments[1] = val2;
					shape5.Adjustments[2] = val3;
					list.Add(shape5.Name);
					shape5 = null;
					Microsoft.Office.Interop.PowerPoint.Shape shape6 = sld.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, 6f, 6f, HarveyBall.m_A - 12, HarveyBall.m_A - 12);
					shape6.Fill.Visible = MsoTriState.msoTrue;
					list.Add(shape6.Name);
					shape6 = null;
					sld.Shapes.Range(list.ToArray()).MergeShapes(MsoMergeCmd.msoMergeSubtract);
					shape7 = sld.Shapes[sld.Shapes.Count];
					Microsoft.Office.Interop.PowerPoint.Shape shape8 = shape7;
					shape8.Line.Visible = MsoTriState.msoFalse;
					Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = shape8.Fill;
					fill2.ForeColor.RGB = Base.GetColor(Base.TurboShapeColor.Primary);
					fill2.BackColor.RGB = fill2.ForeColor.RGB;
					fill2 = null;
					_ = null;
				}
				if (shape7 != null)
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
					if (shape3 != null)
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
						list = new List<string>();
						list.Add(shape3.Name);
						list.Add(shape7.Name);
						shape9 = sld.Shapes.Range(list.ToArray()).Group();
						break;
					}
				}
				if (shape7 != null)
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
					shape9 = shape7;
				}
				else
				{
					shape9 = shape3;
				}
				break;
			}
			}
			Base.FinalizeShape(shape9, Base.TurboShapeType.HarveyBall, val, AH.A(160974));
			tags = shape9.Tags;
			a = HarveyBall.m_A;
		}
		int num2 = (int)style;
		tags.Add(a, num2.ToString());
		list = null;
		shape3 = null;
		shape11 = null;
		return shape9;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, bool B)
	{
		if (B)
		{
			Microsoft.Office.Interop.PowerPoint.Shapes shapes = A.Shapes;
			Microsoft.Office.Interop.PowerPoint.Shape shape = shapes.AddShape(MsoAutoShapeType.msoShapeOval, 0f, 0f, HarveyBall.m_A, HarveyBall.m_A);
			Microsoft.Office.Interop.PowerPoint.Shape shape2 = checked(shapes.AddShape(MsoAutoShapeType.msoShapeOval, 2.5f, 2.5f, HarveyBall.m_A - 5, HarveyBall.m_A - 5));
			string[] index = new string[2] { shape.Name, shape2.Name };
			shapes.Range(index).MergeShapes(MsoMergeCmd.msoMergeSubtract);
			return shapes[shapes.Count];
		}
		return A.Shapes.AddShape(MsoAutoShapeType.msoShapeDonut, 0f, 0f, HarveyBall.m_A, HarveyBall.m_A);
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, int val)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		wpfHarveyBall wpfHarveyBall2 = new wpfHarveyBall();
		wpfHarveyBall2.EditedShape = shpEdit;
		wpfHarveyBall2.CurrentBallStyle = (BallStyle)Conversions.ToInteger(shpEdit.Tags[HarveyBall.m_A].ToString());
		wpfHarveyBall2.BallStyles = new List<BallStyle>();
		wpfHarveyBall2.BallStyles.Add(BallStyle.Traditional);
		wpfHarveyBall2.BallStyles.Add(BallStyle.Modern);
		wpfHarveyBall2.BallStyles.Add(BallStyle.TraditionalTwoColor);
		wpfHarveyBall2.BallStyles.Add(BallStyle.Ring);
		wpfHarveyBall2.Slider.Value = val;
		wpfHarveyBall2.numValue.Value = val;
		wpfHarveyBall2.Top = unitY - wpfHarveyBall2.Height;
		wpfHarveyBall2.Left = unitX;
		wpfHarveyBall2.ShowActivated = false;
		wpfHarveyBall2.Show();
		wpfHarveyBall2 = null;
	}
}
