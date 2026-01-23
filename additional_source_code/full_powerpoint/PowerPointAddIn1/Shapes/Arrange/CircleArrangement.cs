using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Shapes.Arrange;

public sealed class CircleArrangement : Arrangement
{
	public CircleArrangement(int intSlides)
		: base(intSlides)
	{
	}

	internal override void A(List<ShapeItem> A, Container B, Preferences C, ref double D)
	{
		if (D == 0.0)
		{
			float num = Math.Min(B.Height, B.Width) / 2f;
			D = Math.PI * Math.Pow(num, 2.0);
			D = D * (double)C.CircleScale / 100.0 / (double)A.Count;
		}
		this.C(A, (float)D);
		if (C.ScaleMode == ScaleMode.UniformArea)
		{
			D = 0.0;
		}
		this.C(A, B, C);
	}

	internal override void B(List<ShapeItem> A, Container B, Preferences C)
	{
		throw new NotImplementedException();
	}

	private void C(List<ShapeItem> A, Container B, Preferences C)
	{
		float b = C.RotationAngle;
		float c = (float)(360.0 / (double)A.Count);
		bool rotateShapes = C.RotateShapes;
		float d = Math.Min(B.Height, B.Width) / 2f;
		System.Windows.Point e = new System.Windows.Point(B.Left + B.Width / 2f, B.Top + B.Height / 2f);
		switch (C.CircleAlign)
		{
		case CircleAlign.Center:
			E(A, b, c, d, e, rotateShapes);
			break;
		case CircleAlign.Inside:
			this.C(A, b, c, d, e, rotateShapes);
			break;
		case CircleAlign.Outside:
			D(A, b, c, d, e, rotateShapes);
			break;
		}
	}

	private void C(List<ShapeItem> A, float B, float C, float D, System.Windows.Point E)
	{
		using List<ShapeItem>.Enumerator enumerator = A.GetEnumerator();
		Shape shape;
		for (; enumerator.MoveNext(); shape = null, B += C)
		{
			ShapeItem current = enumerator.Current;
			System.Windows.Point point = this.C(E, D, B);
			shape = current.Shape;
			if (B == 0f || B == 360f)
			{
				shape.Left = (float)(point.X - (double)(shape.Width / 2f));
				shape.Top = (float)point.Y;
				continue;
			}
			if (B > 0f)
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
				if (B < 90f)
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
					shape.Left = (float)(point.X - (double)shape.Width);
					shape.Top = (float)point.Y;
					continue;
				}
			}
			if (B == 90f)
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
				shape.Left = (float)(point.X - (double)shape.Width);
				shape.Top = (float)(point.Y - (double)(shape.Height / 2f));
				continue;
			}
			if (B > 90f)
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
				if (B < 180f)
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
					shape.Left = (float)(point.X - (double)shape.Width);
					shape.Top = (float)(point.Y - (double)shape.Height);
					continue;
				}
			}
			if (B == 180f)
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
				shape.Left = (float)(point.X - (double)(shape.Width / 2f));
				shape.Top = (float)(point.Y - (double)shape.Height);
				continue;
			}
			if (B > 180f)
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
				if (B < 270f)
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
					shape.Left = (float)point.X;
					shape.Top = (float)(point.Y - (double)shape.Height);
					continue;
				}
			}
			if (B == 270f)
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
				shape.Left = (float)point.X;
				shape.Top = (float)(point.Y - (double)(shape.Height / 2f));
			}
			else
			{
				shape.Left = (float)point.X;
				shape.Top = (float)point.Y;
			}
		}
	}

	private void C(List<ShapeItem> A, float B, float C, float D, System.Windows.Point E, bool F)
	{
		this.E(A, B, C, D, E, F);
		List<float> list = new List<float>();
		float num = 0f;
		using (List<ShapeItem>.Enumerator enumerator = A.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				Shape shape = enumerator.Current.Shape;
				num = (float)Math.Max(num, Math.Sqrt(Math.Pow(Math.Abs((double)shape.Top - E.Y), 2.0) + Math.Pow(Math.Abs((double)shape.Left - E.X), 2.0)));
				num = (float)Math.Max(num, Math.Sqrt(Math.Pow(Math.Abs((double)shape.Top - E.Y), 2.0) + Math.Pow(Math.Abs((double)(shape.Left + shape.Width) - E.X), 2.0)));
				num = (float)Math.Max(num, Math.Sqrt(Math.Pow(Math.Abs((double)(shape.Top + shape.Height) - E.Y), 2.0) + Math.Pow(Math.Abs((double)(shape.Left + shape.Width) - E.X), 2.0)));
				num = (float)Math.Max(num, Math.Sqrt(Math.Pow(Math.Abs((double)(shape.Top + shape.Height) - E.Y), 2.0) + Math.Pow(Math.Abs((double)shape.Left - E.X), 2.0)));
				shape = null;
				list.Add(num);
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
		}
		num = list.Max();
		D -= num - D;
		this.E(A, B, C, D, E, F);
	}

	private void D(List<ShapeItem> A, float B, float C, float D, System.Windows.Point E)
	{
		using List<ShapeItem>.Enumerator enumerator = A.GetEnumerator();
		Shape shape;
		for (; enumerator.MoveNext(); shape = null, B += C)
		{
			ShapeItem current = enumerator.Current;
			System.Windows.Point point = this.C(E, D, B);
			shape = current.Shape;
			if (B == 0f)
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
				shape.Left = (float)(point.X - (double)(shape.Width / 2f));
				shape.Top = (float)(point.Y - (double)shape.Height);
				continue;
			}
			if (B > 0f)
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
				if (B < 90f)
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
					shape.Left = (float)point.X;
					shape.Top = (float)(point.Y - (double)shape.Height);
					continue;
				}
			}
			if (B == 90f)
			{
				shape.Left = (float)point.X;
				shape.Top = (float)(point.Y - (double)(shape.Height / 2f));
				continue;
			}
			if (B > 90f && B < 180f)
			{
				shape.Left = (float)point.X;
				shape.Top = (float)point.Y;
				continue;
			}
			if (B == 180f)
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
				shape.Left = (float)(point.X - (double)(shape.Width / 2f));
				shape.Top = (float)point.Y;
				continue;
			}
			if (B > 180f)
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
				if (B < 270f)
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
					shape.Left = (float)(point.X - (double)shape.Width);
					shape.Top = (float)point.Y;
					continue;
				}
			}
			if (B == 270f)
			{
				shape.Left = 0 - ((point.X == (double)shape.Width) ? 1 : 0);
				shape.Top = (float)(point.Y - (double)(shape.Height / 2f));
			}
			else
			{
				shape.Left = (float)(point.X - (double)shape.Width);
				shape.Top = (float)(point.Y - (double)shape.Height);
			}
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void D(List<ShapeItem> A, float B, float C, float D, System.Windows.Point E, bool F)
	{
		this.E(A, B, C, D, E, F);
		List<float> list = new List<float>();
		float num = 1000000f;
		using (List<ShapeItem>.Enumerator enumerator = A.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				Shape shape = enumerator.Current.Shape;
				num = (float)Math.Min(num, Math.Sqrt(Math.Pow(Math.Abs((double)shape.Top - E.Y), 2.0) + Math.Pow(Math.Abs((double)shape.Left - E.X), 2.0)));
				num = (float)Math.Min(num, Math.Sqrt(Math.Pow(Math.Abs((double)shape.Top - E.Y), 2.0) + Math.Pow(Math.Abs((double)(shape.Left + shape.Width) - E.X), 2.0)));
				num = (float)Math.Min(num, Math.Sqrt(Math.Pow(Math.Abs((double)(shape.Top + shape.Height) - E.Y), 2.0) + Math.Pow(Math.Abs((double)(shape.Left + shape.Width) - E.X), 2.0)));
				num = (float)Math.Min(num, Math.Sqrt(Math.Pow(Math.Abs((double)(shape.Top + shape.Height) - E.Y), 2.0) + Math.Pow(Math.Abs((double)shape.Left - E.X), 2.0)));
				shape = null;
				list.Add(num);
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				break;
			}
		}
		num = list.Min();
		D += D - num;
		this.E(A, B, C, D, E, F);
	}

	private void C(List<ShapeItem> A, float B)
	{
		foreach (ShapeItem item in A)
		{
			C(item.Shape, B);
		}
	}

	private void E(List<ShapeItem> A, float B, float C, float D, System.Windows.Point E, bool F)
	{
		foreach (ShapeItem item in A)
		{
			System.Windows.Point point = this.C(E, D, B);
			Shape shape = item.Shape;
			shape.Left = (float)(point.X - (double)(shape.Width / 2f));
			shape.Top = (float)(point.Y - (double)(shape.Height / 2f));
			shape = null;
			Shape shape2 = item.Shape;
			float rotation;
			if (!F)
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
				rotation = 0f;
			}
			else
			{
				rotation = B;
			}
			shape2.Rotation = rotation;
			B += C;
		}
	}

	private System.Windows.Point C(System.Windows.Point A, float B, float C)
	{
		C -= 90f;
		double x = A.X + (double)B * Math.Cos(Math.PI * (double)C / 180.0);
		double y = A.Y + (double)B * Math.Sin(Math.PI * (double)C / 180.0);
		return new System.Windows.Point(x, y);
	}

	private float C(System.Windows.Point A, System.Windows.Point B)
	{
		return (float)Math.Sqrt(Math.Pow(Math.Abs(B.X - A.X), 2.0) + Math.Pow(Math.Abs(B.Y - A.Y), 2.0));
	}

	private void D(List<ShapeItem> A, Container B, Preferences C)
	{
		ScaleMode scaleMode = C.ScaleMode;
		if (scaleMode != ScaleMode.UniformArea)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			List<float> list = new List<float>();
			int num = 1;
			using (List<ShapeItem>.Enumerator enumerator = A.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					ShapeItem current = enumerator.Current;
					list.Add(current.Shape.Width * current.Shape.Height);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_006f;
					}
					continue;
					end_IL_006f:
					break;
				}
			}
			float b = (float)((double)list.Min() * Math.Pow(num, 2.0));
			this.C(A, b);
			this.C(A, B, C);
			bool flag = this.C(A, C);
			float num2 = A[1].Shape.Height * A[1].Shape.Width;
			if (flag)
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
				while (flag)
				{
					for (float num3 = num2; num3 >= 1f; num3 += -1f)
					{
						this.C(A, num3);
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_0125;
						}
						continue;
						end_IL_0125:
						break;
					}
					this.C(A, B, C);
					flag = this.C(A, C);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			while (!flag)
			{
				for (float num4 = num2; num4 <= 1000000f; num4 += 1f)
				{
					this.C(A, num4);
				}
				this.C(A, B, C);
				flag = this.C(A, C);
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
	}

	private bool C(List<ShapeItem> A, Preferences B)
	{
		List<RectangleF> list = C(A, B);
		checked
		{
			int num = list.Count - 1;
			int num2 = num;
			for (int i = 0; i <= num2; i++)
			{
				int num3 = num;
				for (int j = 0; j <= num3; j++)
				{
					if (i == j)
					{
						continue;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (!list[i].IntersectsWith(list[j]))
					{
						continue;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						return true;
					}
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_0073;
					}
					continue;
					end_IL_0073:
					break;
				}
			}
			return false;
		}
	}

	private List<RectangleF> C(List<ShapeItem> A, Preferences B)
	{
		List<RectangleF> list = new List<RectangleF>();
		foreach (ShapeItem item in A)
		{
			list.Add(C(item.Shape, B));
		}
		return list;
	}

	private RectangleF C(Shape A, Preferences B)
	{
		return new RectangleF(A.Left - B.MinColumnSpacing / 4f, A.Top - B.MinRowSpacing / 4f, A.Width + B.MinColumnSpacing / 2f, A.Height + B.MinRowSpacing / 2f);
	}
}
