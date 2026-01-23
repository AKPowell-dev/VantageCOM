using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Shapes.Arrange;

public sealed class PyramidArrangement : Arrangement
{
	[CompilerGenerated]
	private new List<RowItem> m_A;

	public List<RowItem> Rows
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public PyramidArrangement(int intShapes, int intSlides)
		: base(intSlides)
	{
		List<RowItem> list = new List<RowItem>();
		int num = 1;
		checked
		{
			for (int i = 0; i < intShapes; i += num)
			{
				list.Add(new RowItem(num, blnCenter: true));
				num++;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Rows = list;
				list = null;
				return;
			}
		}
	}

	internal override void A(List<ShapeItem> A, Container B, Preferences C, ref double D)
	{
		this.C(A, B, C);
		this.D(A, B, C);
	}

	internal override void B(List<ShapeItem> A, Container B, Preferences C)
	{
		throw new NotImplementedException();
	}

	private void C(List<ShapeItem> A, Container B, Preferences C)
	{
		int num = this.C(A.Count);
		ScaleMode scaleMode = C.ScaleMode;
		if (scaleMode != ScaleMode.UniformArea)
		{
			return;
		}
		checked
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				using (List<ShapeItem>.Enumerator enumerator = A.GetEnumerator())
				{
					while (enumerator.MoveNext())
					{
						ShapeItem current = enumerator.Current;
						base.C(current.Shape, 40000f);
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_006a;
						}
						continue;
						end_IL_006a:
						break;
					}
				}
				List<float> list = new List<float>();
				float num2 = 0f;
				int num3 = 0;
				int num4 = 1;
				int count = A.Count;
				float item;
				for (int i = 1; i <= count; i++)
				{
					num2 += A[i - 1].Shape.Width;
					num3++;
					if (num3 == num4)
					{
						item = (B.Width - (float)(num3 - 1) * C.MinColumnSpacing) / num2;
						list.Add(item);
						num2 = 0f;
						num3 = 0;
						num4++;
					}
				}
				item = list.Min();
				list = null;
				using (List<ShapeItem>.Enumerator enumerator2 = A.GetEnumerator())
				{
					while (enumerator2.MoveNext())
					{
						ShapeItem current2 = enumerator2.Current;
						current2.Shape.Width = Math.Max(1f, current2.Shape.Width * item);
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_0163;
						}
						continue;
						end_IL_0163:
						break;
					}
				}
				float num5 = 0f;
				float num6 = 0f;
				num4 = 1;
				int count2 = A.Count;
				for (int j = 1; j <= count2; j++)
				{
					num6 = Math.Max(num6, A[j - 1].Shape.Height);
					num3++;
					if (num3 == num4)
					{
						num5 += num6;
						num6 = 0f;
						num3 = 0;
						num4++;
					}
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					float num7 = B.Height - (float)(num - 1) * C.MinRowSpacing;
					item = ((!(num5 > num7)) ? 1f : (num7 / num5));
					foreach (ShapeItem item2 in A)
					{
						item2.Shape.Height = Math.Max(1f, item2.Shape.Height * item);
					}
					return;
				}
			}
		}
	}

	private void D(List<ShapeItem> A, Container B, Preferences C)
	{
		int num = this.C(A.Count);
		List<float> list = new List<float>();
		List<float> list2 = new List<float>();
		List<int> list3 = new List<int>();
		List<float> list4 = new List<float>();
		List<float> list5 = new List<float>();
		float num2 = 0f;
		float num3 = 0f;
		int num4 = 0;
		int num5 = 1;
		int count = A.Count;
		checked
		{
			for (int i = 1; i <= count; i++)
			{
				Shape shape = A[i - 1].Shape;
				num2 = Math.Max(num2, shape.Height);
				num3 += shape.Width;
				shape = null;
				num4++;
				if (num4 != num5)
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				list4.Add(num2);
				list.Add(num3);
				list3.Add(num4);
				list2.Add(C.MinColumnSpacing);
				num2 = 0f;
				num3 = 0f;
				num4 = 0;
				num5++;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				float num6;
				if (C.StretchHeight)
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
					num6 = Math.Max(C.MinRowSpacing, (B.Height - list4.Sum()) / (float)(num - 1));
				}
				else
				{
					num6 = C.MinRowSpacing;
				}
				float num7 = B.Top;
				using (List<float>.Enumerator enumerator = list4.GetEnumerator())
				{
					while (enumerator.MoveNext())
					{
						float num8 = enumerator.Current;
						list5.Add(num7);
						num7 += num8 + num6;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_017b;
						}
						continue;
						end_IL_017b:
						break;
					}
				}
				int num9 = 0;
				int num10 = num - 1;
				for (int j = 0; j <= num10; j++)
				{
					num4 = list3[j];
					float num11 = list2[j];
					float num12 = B.Left + (B.Width - list[j] - (float)(num4 - 1) * num11) / 2f;
					int num13 = num4;
					for (int k = 1; k <= num13; k++)
					{
						Shape shape2 = A[num9].Shape;
						shape2.Left = num12;
						num12 += shape2.Width + num11;
						shape2.Top = list5[j] + list4[j] / 2f - shape2.Height / 2f;
						shape2 = null;
						num9++;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_0267;
						}
						continue;
						end_IL_0267:
						break;
					}
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
		}
	}

	private float C(float[] A)
	{
		Array.Sort(A);
		if (A.Length % 2 == 0)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return (A[checked(unchecked(A.Length / 2) - 1)] + A[A.Length / 2]) / 2f;
				}
			}
		}
		return A[A.Length / 2];
	}

	private float D(float[] A)
	{
		double num = 0.0;
		double num2 = 0.0;
		int num3 = A.Length;
		checked
		{
			int num4 = num3 - 1;
			for (int i = 0; i <= num4; i++)
			{
				num += (double)A[i];
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
				num /= (double)num3;
				int num5 = num3 - 1;
				for (int j = 0; j <= num5; j++)
				{
					num2 += ((double)A[j] - num) * ((double)A[j] - num);
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					return (float)Math.Sqrt(num2 / (double)num3);
				}
			}
		}
	}

	private int C(int A)
	{
		if (A <= 15)
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
			switch (A)
			{
			case 6:
				return 3;
			case 10:
				return 4;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
			if (A == 15)
			{
				return 5;
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
		}
		else if (A <= 28)
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
			if (A == 21)
			{
				return 6;
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
			if (A == 28)
			{
				return 7;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		else
		{
			if (A == 36)
			{
				return 8;
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
			if (A == 45)
			{
				return 9;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		int result = default(int);
		return result;
	}
}
