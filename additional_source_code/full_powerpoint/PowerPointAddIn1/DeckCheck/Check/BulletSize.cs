using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class BulletSize : BaseCheck
{
	private struct YB
	{
		public int A;

		public float A;

		public TextRange2 A;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<YB, O<float>> A;

		public static Func<YB, float> A;

		public static Func<YB, YB> A;

		public static Func<float, IEnumerable<YB>, P<float, IEnumerable<YB>>> A;

		public static Func<P<float, IEnumerable<YB>>, int> A;

		public static Func<P<float, IEnumerable<YB>>, Q<float, int>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal O<float> A(YB A)
		{
			return new O<float>(A.A);
		}

		[SpecialName]
		internal float A(YB A)
		{
			return A.A;
		}

		[SpecialName]
		internal YB A(YB A)
		{
			return A;
		}

		[SpecialName]
		internal P<float, IEnumerable<YB>> A(float A, IEnumerable<YB> B)
		{
			return new P<float, IEnumerable<YB>>(A, B);
		}

		[SpecialName]
		internal int A(P<float, IEnumerable<YB>> A)
		{
			return A.Group.Count();
		}

		[SpecialName]
		internal Q<float, int> A(P<float, IEnumerable<YB>> A)
		{
			return new Q<float, int>(A.size, A.Group.Count());
		}
	}

	[CompilerGenerated]
	internal sealed class ZB
	{
		public int A;

		public Func<YB, bool> A;

		public ZB(ZB A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(YB A)
		{
			return A.A == this.A;
		}
	}

	[CompilerGenerated]
	private List<YB> A;

	private List<YB> Sizes
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	}

	public BulletSize()
	{
		Sizes = new List<YB>();
	}

	public void CheckParagraph(TextRange2 para)
	{
		YB item = new YB
		{
			A = para
		};
		ParagraphFormat2 paragraphFormat = para.ParagraphFormat;
		item.A = paragraphFormat.IndentLevel;
		item.A = paragraphFormat.Bullet.RelativeSize;
		paragraphFormat = null;
		Sizes.Add(item);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		ZB a = default(ZB);
		ZB CS_0024_003C_003E8__locals5 = new ZB(a);
		if (Sizes.Count <= 0)
		{
			return;
		}
		int num = 1;
		do
		{
			CS_0024_003C_003E8__locals5.A = num;
			try
			{
				List<YB> sizes = Sizes;
				Func<YB, bool> predicate;
				if (CS_0024_003C_003E8__locals5.A != null)
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
					predicate = CS_0024_003C_003E8__locals5.A;
				}
				else
				{
					predicate = (CS_0024_003C_003E8__locals5.A = [SpecialName] (YB A) => A.A == CS_0024_003C_003E8__locals5.A);
				}
				List<YB> list = sizes.Where(predicate).ToList();
				List<YB> source = list;
				Func<YB, O<float>> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (YB A) => new O<float>(A.A));
				}
				else
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
					keySelector = _Closure_0024__.A;
				}
				if (source.GroupBy(keySelector).ToList().Count > 1)
				{
					List<TextRange2> list2 = new List<TextRange2>();
					foreach (YB item in list)
					{
						list2.Add(item.A);
					}
					List<float> list3 = new List<float>();
					List<string> list4 = new List<string>();
					List<YB> source2 = list;
					Func<YB, float> keySelector2;
					if (_Closure_0024__.A == null)
					{
						keySelector2 = (_Closure_0024__.A = [SpecialName] (YB A) => A.A);
					}
					else
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
						keySelector2 = _Closure_0024__.A;
					}
					Func<YB, YB> elementSelector;
					if (_Closure_0024__.A == null)
					{
						elementSelector = (_Closure_0024__.A = [SpecialName] (YB A) => A);
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
						elementSelector = _Closure_0024__.A;
					}
					Func<float, IEnumerable<YB>, P<float, IEnumerable<YB>>> resultSelector;
					if (_Closure_0024__.A == null)
					{
						resultSelector = (_Closure_0024__.A = [SpecialName] (float A, IEnumerable<YB> B) => new P<float, IEnumerable<YB>>(A, B));
					}
					else
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
						resultSelector = _Closure_0024__.A;
					}
					IEnumerable<P<float, IEnumerable<YB>>> source3 = source2.GroupBy(keySelector2, elementSelector, resultSelector);
					Func<P<float, IEnumerable<YB>>, int> keySelector3;
					if (_Closure_0024__.A == null)
					{
						keySelector3 = (_Closure_0024__.A = [SpecialName] (P<float, IEnumerable<YB>> A) => A.Group.Count());
					}
					else
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
						keySelector3 = _Closure_0024__.A;
					}
					IOrderedEnumerable<P<float, IEnumerable<YB>>> source4 = source3.OrderByDescending(keySelector3);
					Func<P<float, IEnumerable<YB>>, Q<float, int>> selector;
					if (_Closure_0024__.A == null)
					{
						selector = (_Closure_0024__.A = [SpecialName] (P<float, IEnumerable<YB>> A) => new Q<float, int>(A.size, A.Group.Count()));
					}
					else
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
						selector = _Closure_0024__.A;
					}
					IEnumerable<Q<float, int>> enumerable = source4.Select(selector);
					using (IEnumerator<Q<float, int>> enumerator2 = enumerable.GetEnumerator())
					{
						while (enumerator2.MoveNext())
						{
							Q<float, int> current = enumerator2.Current;
							list3.Add(current.sz);
							list4.Add(current.sz.ToString(AH.A(14595)) + AH.A(14248) + current.cnt + AH.A(14255));
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_028a;
							}
							continue;
							end_IL_028a:
							break;
						}
					}
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.BulletSize(sld, shp, list4, string.Join(AH.A(14258), list4.ToArray()), list2, list3));
					list3 = null;
					list4 = null;
					list2 = null;
				}
				list = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			num = checked(num + 1);
		}
		while (num <= 10);
		Sizes.Clear();
	}
}
