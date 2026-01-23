using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.TurboShapes;

public sealed class Tachometer
{
	public enum Style
	{
		TriColor270 = 1,
		TriColor180,
		BiColor270,
		BiColor270NoNeedle,
		Rail270
	}

	public delegate int ComponentSelector(Color color);

	public static readonly string TAG_TACH_STYLE = AH.A(161540);

	public static readonly string TAG_TACH_REVERSED = AH.A(161571);

	public static readonly string TAG_TACH_LABEL = AH.A(161608);

	private static ComponentSelector m_A = [SpecialName] (Color A) => A.R;

	private static ComponentSelector m_B = [SpecialName] (Color A) => A.G;

	private static ComponentSelector C = [SpecialName] (Color A) => A.B;

	public static void Add()
	{
		Base.AddTurboShape(A);
	}

	private static void A(Slide A, PageSetup B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = Create(A, 50f, 1, blnReverse: false, blnLabel: false);
		shape.Height = 36f;
		shape.Top = B.SlideHeight / 2f - shape.Height / 2f;
		shape.Left = B.SlideWidth / 2f - shape.Width / 2f;
		shape = null;
		Base.LogActivity(AH.A(161460));
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape Create(Slide sld, float val, int sty, bool blnReverse, bool blnLabel)
	{
		Style style = (Style)sty;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
		List<string> list2 = new List<string>();
		if (blnLabel)
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
			if (style != Style.TriColor180)
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
				Microsoft.Office.Interop.PowerPoint.Shape shape2 = null;
				try
				{
					shape2 = Helpers.GetBodyPlaceholder(sld.Application.ActivePresentation);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				int rGB = default(int);
				string name = default(string);
				if (shape2 != null)
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
					Font2 font = shape2.TextFrame2.TextRange.Font;
					rGB = font.Fill.ForeColor.RGB;
					name = font.Name;
					_ = null;
				}
				shape = sld.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0f, 0f, 100f, 50f);
				Microsoft.Office.Interop.PowerPoint.Shape shape3 = shape;
				Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shape3.TextFrame2;
				TextRange2 textRange = textFrame.TextRange;
				textRange.Text = checked((int)Math.Round(val)) + AH.A(161503);
				textRange.Font.Size = 7f;
				if (shape2 != null)
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
					if (rGB >= 0)
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
						textRange.Font.Fill.ForeColor.RGB = rGB;
					}
					if (Operators.CompareString(name, string.Empty, TextCompare: false) != 0)
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
						textRange.Font.Name = name;
					}
					shape2 = null;
				}
				textRange = null;
				textFrame.WordWrap = MsoTriState.msoFalse;
				textFrame.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
				textFrame.MarginTop = 0f;
				textFrame.MarginBottom = 0f;
				textFrame.MarginLeft = 0f;
				textFrame.MarginRight = 0f;
				_ = null;
				shape3.Left = (36f - shape3.Width) / 2f;
				switch (style)
				{
				case Style.TriColor270:
				case Style.BiColor270:
					shape3.Top = 36f - shape3.Height + 1f;
					break;
				case Style.BiColor270NoNeedle:
					shape3.Top = (36f - shape3.Height) / 2f;
					break;
				}
				shape3 = null;
			}
		}
		Base.TurboShapeColor A = default(Base.TurboShapeColor);
		Base.TurboShapeColor B = default(Base.TurboShapeColor);
		Microsoft.Office.Interop.PowerPoint.Shape item = default(Microsoft.Office.Interop.PowerPoint.Shape);
		Microsoft.Office.Interop.PowerPoint.Shape shape6;
		Microsoft.Office.Interop.PowerPoint.Shape item4;
		Microsoft.Office.Interop.PowerPoint.Shape item5;
		Microsoft.Office.Interop.PowerPoint.Shape item2;
		Microsoft.Office.Interop.PowerPoint.Shape item3;
		switch (style)
		{
		case Style.TriColor270:
		{
			Tachometer.A(ref A, ref B, blnReverse);
			Microsoft.Office.Interop.PowerPoint.Shape shape4 = Tachometer.A(sld, 36f, A);
			Microsoft.Office.Interop.PowerPoint.Shape shape5 = shape4;
			Adjustments adjustments = shape5.Adjustments;
			adjustments[1] = 135f;
			adjustments[2] = 225f;
			_ = null;
			list2.Add(shape5.Name);
			shape5 = null;
			shape6 = Tachometer.A(sld, 36f);
			list2.Add(shape6.Name);
			shape4 = Base.SubtractShapes(sld, list2);
			list2.Clear();
			Microsoft.Office.Interop.PowerPoint.Shape shape7 = Tachometer.A(sld, 36f, Base.TurboShapeColor.Yellow);
			Microsoft.Office.Interop.PowerPoint.Shape shape8 = shape7;
			Adjustments adjustments2 = shape8.Adjustments;
			adjustments2[1] = 225f;
			adjustments2[2] = 315.01f;
			_ = null;
			list2.Add(shape8.Name);
			shape8 = null;
			shape6 = Tachometer.A(sld, 36f);
			list2.Add(shape6.Name);
			shape7 = Base.SubtractShapes(sld, list2);
			list2.Clear();
			Microsoft.Office.Interop.PowerPoint.Shape shape9 = Tachometer.A(sld, 36f, B);
			Microsoft.Office.Interop.PowerPoint.Shape shape10 = shape9;
			Adjustments adjustments3 = shape10.Adjustments;
			adjustments3[1] = 315f;
			adjustments3[2] = 45f;
			_ = null;
			list2.Add(shape10.Name);
			shape10 = null;
			shape6 = Tachometer.A(sld, 36f);
			list2.Add(shape6.Name);
			shape9 = Base.SubtractShapes(sld, list2);
			item4 = Tachometer.A(sld, 36f, val, (Style)sty);
			List<Microsoft.Office.Interop.PowerPoint.Shape> list3 = list;
			list3.Add(shape4);
			list3.Add(shape7);
			list3.Add(shape9);
			list3.Add(item4);
			if (shape != null)
			{
				list3.Add(shape);
			}
			list3 = null;
			item = Tachometer.A(sld, list);
			shape4 = null;
			shape7 = null;
			shape9 = null;
			break;
		}
		case Style.TriColor180:
		{
			Tachometer.A(ref A, ref B, blnReverse);
			Microsoft.Office.Interop.PowerPoint.Shape shape11 = Tachometer.A(sld, 36f, A);
			Microsoft.Office.Interop.PowerPoint.Shape shape12 = shape11;
			Adjustments adjustments4 = shape12.Adjustments;
			adjustments4[1] = 180f;
			adjustments4[2] = 240f;
			_ = null;
			list2.Add(shape12.Name);
			shape12 = null;
			shape6 = Tachometer.A(sld, 36f);
			list2.Add(shape6.Name);
			shape11 = Base.SubtractShapes(sld, list2);
			list2.Clear();
			Microsoft.Office.Interop.PowerPoint.Shape shape13 = Tachometer.A(sld, 36f, Base.TurboShapeColor.Yellow);
			Microsoft.Office.Interop.PowerPoint.Shape shape14 = shape13;
			Adjustments adjustments5 = shape14.Adjustments;
			adjustments5[1] = 240f;
			adjustments5[2] = 300f;
			_ = null;
			list2.Add(shape14.Name);
			shape14 = null;
			shape6 = Tachometer.A(sld, 36f);
			list2.Add(shape6.Name);
			shape13 = Base.SubtractShapes(sld, list2);
			list2.Clear();
			Microsoft.Office.Interop.PowerPoint.Shape shape15 = Tachometer.A(sld, 36f, B);
			Microsoft.Office.Interop.PowerPoint.Shape shape16 = shape15;
			Adjustments adjustments6 = shape16.Adjustments;
			adjustments6[1] = 300f;
			adjustments6[2] = 0f;
			_ = null;
			list2.Add(shape16.Name);
			shape16 = null;
			shape6 = Tachometer.A(sld, 36f);
			list2.Add(shape6.Name);
			shape15 = Base.SubtractShapes(sld, list2);
			item4 = Tachometer.A(sld, 36f, val, (Style)sty);
			List<Microsoft.Office.Interop.PowerPoint.Shape> list4 = list;
			list4.Add(shape11);
			list4.Add(shape13);
			list4.Add(shape15);
			list4.Add(item4);
			if (shape != null)
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
				list4.Add(shape);
			}
			list4 = null;
			item = Tachometer.A(sld, list);
			shape11 = null;
			shape13 = null;
			shape15 = null;
			break;
		}
		case Style.BiColor270:
			if (val == 0f)
			{
				item5 = Tachometer.B(sld, 36f);
				item4 = Tachometer.A(sld, 36f, val, (Style)sty);
				list.Add(item5);
				list.Add(item4);
			}
			else if (val == 100f)
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
				item5 = Tachometer.A(sld, 36f, blnReverse);
				item4 = Tachometer.A(sld, 36f, val, (Style)sty);
				list.Add(item5);
				list.Add(item4);
			}
			else
			{
				item2 = Tachometer.B(sld, 36f);
				item3 = Tachometer.A(sld, 36f, val, blnReverse);
				item4 = Tachometer.A(sld, 36f, val, (Style)sty);
				list.Add(item2);
				list.Add(item3);
				list.Add(item4);
			}
			if (shape != null)
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
				list.Add(shape);
			}
			item = Tachometer.A(sld, list);
			break;
		case Style.BiColor270NoNeedle:
			if (val == 0f)
			{
				item = Tachometer.B(sld, 36f);
				list.Add(item);
			}
			else if (val == 100f)
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
				item = Tachometer.A(sld, 36f, blnReverse);
				list.Add(item);
			}
			else
			{
				item2 = Tachometer.B(sld, 36f);
				item3 = Tachometer.A(sld, 36f, val, blnReverse);
				list.Add(item2);
				list.Add(item3);
			}
			if (shape != null)
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
				list.Add(shape);
			}
			if (list.Count > 1)
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
				item = Tachometer.A(sld, list);
			}
			else
			{
				item = list[0];
			}
			break;
		}
		Base.FinalizeShape(item, Base.TurboShapeType.Tachometer, val, AH.A(161506));
		Tags tags = item.Tags;
		tags.Add(TAG_TACH_STYLE, sty.ToString());
		tags.Add(TAG_TACH_REVERSED, blnReverse.ToString());
		tags.Add(TAG_TACH_LABEL, blnLabel.ToString());
		_ = null;
		shape6 = null;
		item4 = null;
		item5 = null;
		item2 = null;
		item3 = null;
		shape = null;
		list = null;
		list2 = null;
		return item;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, float B, Base.TurboShapeColor C)
	{
		return Tachometer.A(A, B, Base.GetColor(C));
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, float B, int C)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shapes.AddShape(MsoAutoShapeType.msoShapePie, 0f, 0f, B, B);
		shape.Rotation = 0f;
		shape.Fill.Visible = MsoTriState.msoTrue;
		shape.Line.Visible = MsoTriState.msoFalse;
		shape.LockAspectRatio = MsoTriState.msoTrue;
		shape.Fill.ForeColor.RGB = C;
		shape.Fill.BackColor.RGB = C;
		_ = null;
		return shape;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, float B)
	{
		int num = 8;
		Microsoft.Office.Interop.PowerPoint.Shape shape = checked(A.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, num, num, B - (float)(num * 2), B - (float)(num * 2)));
		shape.Fill.Visible = MsoTriState.msoTrue;
		return shape;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, float B, float C, Style D)
	{
		List<string> list = new List<string>();
		Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, (B - 6f) / 2f, (B - 6f) / 2f, 6f, 6f);
		shape.Fill.Visible = MsoTriState.msoTrue;
		shape.Line.Visible = MsoTriState.msoFalse;
		shape.LockAspectRatio = MsoTriState.msoTrue;
		shape.Fill.ForeColor.RGB = 0;
		shape.Fill.BackColor.RGB = 0;
		list.Add(shape.Name);
		float c = shape.Top + shape.Height / 2f;
		shape = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = A.Shapes.AddShape(MsoAutoShapeType.msoShapeIsoscelesTriangle, (B - 4f) / 2f, 3f, 4f, 20f);
		Microsoft.Office.Interop.PowerPoint.Shape shape3 = shape2;
		shape3.Fill.Visible = MsoTriState.msoTrue;
		shape3.Line.Visible = MsoTriState.msoFalse;
		shape3.LockAspectRatio = MsoTriState.msoTrue;
		shape3.Fill.ForeColor.RGB = 0;
		shape3.Fill.BackColor.RGB = 0;
		list.Add(shape3.Name);
		shape3 = null;
		int num;
		if (C != 50f)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (D != Style.TriColor270)
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
				if (D != Style.BiColor270)
				{
					num = 90;
					goto IL_0185;
				}
			}
			num = 135;
			goto IL_0185;
		}
		goto IL_019e;
		IL_0185:
		float b = (float)num * (C - 50f) / 50f;
		Tachometer.A(shape2, b, c);
		goto IL_019e;
		IL_019e:
		shape2 = Base.MergeShapes(A, list);
		list = null;
		return shape2;
	}

	private static Base.TurboShapeColor A(float A, bool B)
	{
		if (!B)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if ((double)A < 33.3333)
					{
						return Base.TurboShapeColor.Green;
					}
					if ((double)A < 66.6667)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								return Base.TurboShapeColor.Yellow;
							}
						}
					}
					return Base.TurboShapeColor.Red;
				}
			}
		}
		if ((double)A < 33.3333)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return Base.TurboShapeColor.Red;
				}
			}
		}
		if ((double)A < 66.6667)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return Base.TurboShapeColor.Yellow;
				}
			}
		}
		return Base.TurboShapeColor.Green;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, float B, float C, bool D)
	{
		List<string> list = new List<string>();
		Base.TurboShapeColor c;
		if (!D)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if ((double)C < 33.3333)
			{
				c = Base.TurboShapeColor.Green;
			}
			else if ((double)C < 66.6667)
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
				c = Base.TurboShapeColor.Yellow;
			}
			else
			{
				c = Base.TurboShapeColor.Red;
			}
		}
		else if ((double)C < 33.3333)
		{
			c = Base.TurboShapeColor.Red;
		}
		else if ((double)C < 66.6667)
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
			c = Base.TurboShapeColor.Yellow;
		}
		else
		{
			c = Base.TurboShapeColor.Green;
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape = Tachometer.A(A, B, c);
		list.Add(shape.Name);
		float num = 270f * C / 100f;
		Adjustments adjustments = shape.Adjustments;
		adjustments[1] = 135f;
		adjustments[2] = 135f + num;
		_ = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = Tachometer.A(A, B);
		list.Add(shape2.Name);
		shape = Base.SubtractShapes(A, list);
		shape2 = null;
		return shape;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape B(Slide A, float B)
	{
		List<string> list = new List<string>();
		Microsoft.Office.Interop.PowerPoint.Shape shape = Tachometer.A(A, B, Base.TurboShapeColor.Secondary);
		Adjustments adjustments = shape.Adjustments;
		adjustments[1] = 135f;
		adjustments[2] = 45.01f;
		_ = null;
		list.Add(shape.Name);
		shape = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = Tachometer.A(A, B);
		list.Add(shape2.Name);
		Microsoft.Office.Interop.PowerPoint.Shape result = Base.SubtractShapes(A, list);
		list = null;
		return result;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, float B, bool C)
	{
		List<string> list = new List<string>();
		Base.TurboShapeColor c;
		if (!C)
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
			c = Base.TurboShapeColor.Red;
		}
		else
		{
			c = Base.TurboShapeColor.Green;
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape = Tachometer.A(A, B, c);
		Adjustments adjustments = shape.Adjustments;
		adjustments[1] = 135f;
		adjustments[2] = 45.01f;
		_ = null;
		list.Add(shape.Name);
		shape = null;
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = Tachometer.A(A, B);
		list.Add(shape2.Name);
		Microsoft.Office.Interop.PowerPoint.Shape result = Base.SubtractShapes(A, list);
		shape2 = null;
		list = null;
		return result;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, float B, float C)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		shape.Rotation = B;
		double num = Math.PI;
		float num2 = shape.Top + shape.Height / 2f - C;
		float num3 = (float)((double)(2f * num2) * Math.Sin(0.5 * num * (double)Math.Abs(B) / 180.0));
		float num4 = 90f - (180f - Math.Abs(B) / 2f);
		float num5 = 90f - num4;
		float num6 = (float)((double)num3 * Math.Cos(num * (double)num5 / 180.0));
		float num7 = (float)((double)num3 * Math.Sin(num * (double)num5 / 180.0));
		Microsoft.Office.Interop.PowerPoint.Shape shape3;
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = (shape3 = shape);
		float left = shape3.Left;
		int num8;
		if (!(B > 0f))
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
			num8 = -1;
		}
		else
		{
			num8 = 1;
		}
		shape2.Left = left + num6 * (float)num8;
		shape.Top -= num7;
		shape = null;
	}

	private static void A(ref Base.TurboShapeColor A, ref Base.TurboShapeColor B, bool C)
	{
		if (!C)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					A = Base.TurboShapeColor.Green;
					B = Base.TurboShapeColor.Red;
					return;
				}
			}
		}
		A = Base.TurboShapeColor.Red;
		B = Base.TurboShapeColor.Green;
	}

	private static Microsoft.Office.Interop.PowerPoint.Shape A(Slide A, List<Microsoft.Office.Interop.PowerPoint.Shape> B)
	{
		return A.Shapes.Range(B.Select([SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => shape.Name).ToArray()).Group();
	}

	public static void Edit(Microsoft.Office.Interop.PowerPoint.Shape shpEdit, int val)
	{
		double unitX = default(double);
		double unitY = default(double);
		Base.TransformFromShape(shpEdit, Base.CalloutPosition.TopCenter, ref unitX, ref unitY);
		wpfTachometer wpfTachometer2 = new wpfTachometer();
		wpfTachometer2.EditedShape = shpEdit;
		wpfTachometer2.CurrentStyle = (Style)Conversions.ToInteger(shpEdit.Tags[TAG_TACH_STYLE].ToString());
		wpfTachometer2.IsReversed = Conversions.ToBoolean(shpEdit.Tags[TAG_TACH_REVERSED].ToString());
		wpfTachometer2.ShowLabel = Conversions.ToBoolean(shpEdit.Tags[TAG_TACH_LABEL].ToString());
		wpfTachometer2.Styles = new List<Style>();
		wpfTachometer2.Styles.Add(Style.TriColor270);
		wpfTachometer2.Styles.Add(Style.TriColor180);
		wpfTachometer2.Styles.Add(Style.BiColor270);
		wpfTachometer2.Styles.Add(Style.BiColor270NoNeedle);
		wpfTachometer2.Slider.Value = val;
		wpfTachometer2.numValue.Value = val;
		wpfTachometer2.Top = unitY - wpfTachometer2.Height;
		wpfTachometer2.Left = unitX;
		wpfTachometer2.ShowActivated = false;
		wpfTachometer2.Show();
		wpfTachometer2 = null;
	}

	public static Color InterpolateBetween(Color endPoint1, Color endPoint2, double lambda)
	{
		if (!(lambda < 0.0))
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
			if (!(lambda > 1.0))
			{
				return Color.FromArgb(A(endPoint1, endPoint2, lambda, Tachometer.m_A), A(endPoint1, endPoint2, lambda, Tachometer.m_B), A(endPoint1, endPoint2, lambda, C));
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
		}
		throw new ArgumentOutOfRangeException(AH.A(161527));
	}

	private static int A(Color A, Color B, double C, ComponentSelector D)
	{
		return checked((int)Math.Round((double)D(A) + (double)(D(B) - D(A)) * C));
	}
}
