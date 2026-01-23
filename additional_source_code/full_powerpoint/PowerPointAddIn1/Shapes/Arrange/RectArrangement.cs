using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.Shapes.Arrange;

public sealed class RectArrangement : Arrangement
{
	[CompilerGenerated]
	private new List<RowItem> m_A;

	[CompilerGenerated]
	private new int m_A;

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

	private int Columns
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

	public RectArrangement(List<RowItem> listRows, int intColumns, int intSlides)
		: base(intSlides)
	{
		Rows = listRows;
		Columns = intColumns;
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
				switch (3)
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
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_0068;
						}
						continue;
						end_IL_0068:
						break;
					}
				}
				List<float> list = new List<float>();
				float num2 = 0f;
				int num3 = 0;
				int count = A.Count;
				for (int i = 1; i <= count; i++)
				{
					num2 += A[i - 1].Shape.Width;
					num3++;
					if (this.C(i, A.Count))
					{
						float item = (B.Width - (float)(num3 - 1) * C.MinColumnSpacing) / num2;
						list.Add(item);
						num2 = 0f;
						num3 = 0;
					}
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					float item = list.Min();
					list = null;
					foreach (ShapeItem item2 in A)
					{
						item2.Shape.Width = Math.Max(1f, item2.Shape.Width * item);
					}
					float num4 = 0f;
					float num5 = 0f;
					int count2 = A.Count;
					for (int j = 1; j <= count2; j++)
					{
						num5 = Math.Max(num5, A[j - 1].Shape.Height);
						if (!this.C(j, A.Count))
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
						num4 += num5;
						num5 = 0f;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						float num6 = B.Height - (float)(num - 1) * C.MinRowSpacing;
						if (num4 > num6)
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
							item = num6 / num4;
						}
						else
						{
							item = 1f;
						}
						using List<ShapeItem>.Enumerator enumerator3 = A.GetEnumerator();
						while (enumerator3.MoveNext())
						{
							ShapeItem current3 = enumerator3.Current;
							current3.Shape.Height = Math.Max(1f, current3.Shape.Height * item);
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
		int num5 = 0;
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
				if (!this.C(i, A.Count))
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
				float num6;
				if (num4 > 1)
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
					num6 = ((!C.StretchWidth) ? C.MinColumnSpacing : Math.Max(C.MinColumnSpacing, (B.Width - num3) / (float)(num4 - 1)));
				}
				else
				{
					num6 = 0f;
				}
				num5++;
				if (num5 >= num)
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
					if (num != 1)
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
						if (num4 != Columns)
						{
							float num7 = list2.Max();
							if (num6 > num7)
							{
								list2.Add(num7);
							}
							else
							{
								list2.Add(num6);
							}
							goto IL_0184;
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
				}
				list2.Add(num6);
				goto IL_0184;
				IL_0184:
				num2 = 0f;
				num3 = 0f;
				num4 = 0;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				float num8;
				if (C.StretchHeight)
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
					num8 = Math.Max(C.MinRowSpacing, (B.Height - list4.Sum()) / (float)(num - 1));
				}
				else
				{
					num8 = C.MinRowSpacing;
				}
				float num9 = B.Top;
				foreach (float item in list4)
				{
					float num10 = item;
					list5.Add(num9);
					num9 += num10 + num8;
				}
				int num11 = 0;
				int num12 = num - 1;
				for (int num13 = 0; num13 <= num12; num13++)
				{
					num4 = list3[num13];
					float num14 = list2[num13];
					float num15;
					if (C.CenterShapes)
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
						if (num > 1)
						{
							num15 = B.Left + (B.Width - list[num13] - (float)(num4 - 1) * num14) / 2f;
							goto IL_02bc;
						}
					}
					num15 = B.Left;
					goto IL_02bc;
					IL_02bc:
					int num16 = num4;
					for (int j = 1; j <= num16; j++)
					{
						Shape shape2 = A[num11].Shape;
						shape2.Left = num15;
						num15 += shape2.Width + num14;
						shape2.Top = list5[num13] + list4[num13] / 2f - shape2.Height / 2f;
						shape2 = null;
						num11++;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_0339;
						}
						continue;
						end_IL_0339:
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
				switch (5)
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
				switch (4)
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
				return (float)Math.Sqrt(num2 / (double)num3);
			}
		}
	}

	private int C(int A)
	{
		return checked((int)Math.Ceiling((double)A / (double)Columns));
	}

	private bool C(int A, int B)
	{
		if (A % Columns != 0)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return A == B;
				}
			}
		}
		return true;
	}
}
