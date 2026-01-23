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

public sealed class BulletIndentation : BaseCheck
{
	private struct WB
	{
		public int A;

		public float A;

		public float B;

		public TextRange2 A;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<WB, H<float, float>> A;

		public static Func<WB, I<float, float>> A;

		public static Func<WB, WB> A;

		public static Func<I<float, float>, IEnumerable<WB>, J<float, float, IEnumerable<WB>>> A;

		public static Func<J<float, float, IEnumerable<WB>>, int> A;

		public static Func<J<float, float, IEnumerable<WB>>, K<float, float, int>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal H<float, float> A(WB A)
		{
			return new H<float, float>(A.A, A.B);
		}

		[SpecialName]
		internal I<float, float> A(WB A)
		{
			return new I<float, float>(A.A, A.B);
		}

		[SpecialName]
		internal WB A(WB A)
		{
			return A;
		}

		[SpecialName]
		internal J<float, float, IEnumerable<WB>> A(I<float, float> A, IEnumerable<WB> B)
		{
			return new J<float, float, IEnumerable<WB>>(A.left, A.first, B);
		}

		[SpecialName]
		internal int A(J<float, float, IEnumerable<WB>> A)
		{
			return A.Group.Count();
		}

		[SpecialName]
		internal K<float, float, int> A(J<float, float, IEnumerable<WB>> A)
		{
			return new K<float, float, int>(A.left, A.first, A.Group.Count());
		}
	}

	[CompilerGenerated]
	internal sealed class XB
	{
		public int A;

		public Func<WB, bool> A;

		public XB(XB A)
		{
			if (A == null)
			{
				return;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(WB A)
		{
			return A.A == this.A;
		}
	}

	[CompilerGenerated]
	private List<WB> A;

	private List<WB> Indentation
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

	public BulletIndentation()
	{
		Indentation = new List<WB>();
	}

	public void CheckParagraph(TextRange2 para)
	{
		WB item = new WB
		{
			A = para
		};
		ParagraphFormat2 paragraphFormat = para.ParagraphFormat;
		item.A = paragraphFormat.IndentLevel;
		item.A = paragraphFormat.LeftIndent;
		item.B = paragraphFormat.FirstLineIndent;
		paragraphFormat = null;
		Indentation.Add(item);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		XB a = default(XB);
		XB CS_0024_003C_003E8__locals5 = new XB(a);
		if (Indentation.Count <= 0)
		{
			return;
		}
		int num = 1;
		do
		{
			CS_0024_003C_003E8__locals5.A = num;
			try
			{
				List<WB> indentation = Indentation;
				Func<WB, bool> predicate;
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
					predicate = (CS_0024_003C_003E8__locals5.A = [SpecialName] (WB A) => A.A == CS_0024_003C_003E8__locals5.A);
				}
				List<WB> list = indentation.Where(predicate).ToList();
				if ((from A in list
					group A by new H<float, float>(A.A, A.B)).Count() > 1)
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
					List<TextRange2> list2 = new List<TextRange2>();
					foreach (WB item in list)
					{
						list2.Add(item.A);
					}
					List<BulletIndentFix> list3 = new List<BulletIndentFix>();
					List<string> list4 = new List<string>();
					List<WB> source = list;
					Func<WB, I<float, float>> keySelector;
					if (_Closure_0024__.A == null)
					{
						keySelector = (_Closure_0024__.A = [SpecialName] (WB A) => new I<float, float>(A.A, A.B));
					}
					else
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
						keySelector = _Closure_0024__.A;
					}
					Func<WB, WB> elementSelector;
					if (_Closure_0024__.A == null)
					{
						elementSelector = (_Closure_0024__.A = [SpecialName] (WB A) => A);
					}
					else
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
						elementSelector = _Closure_0024__.A;
					}
					Func<I<float, float>, IEnumerable<WB>, J<float, float, IEnumerable<WB>>> resultSelector;
					if (_Closure_0024__.A == null)
					{
						resultSelector = (_Closure_0024__.A = [SpecialName] (I<float, float> A, IEnumerable<WB> B) => new J<float, float, IEnumerable<WB>>(A.left, A.first, B));
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
					IEnumerable<J<float, float, IEnumerable<WB>>> source2 = source.GroupBy(keySelector, elementSelector, resultSelector);
					Func<J<float, float, IEnumerable<WB>>, int> keySelector2;
					if (_Closure_0024__.A == null)
					{
						keySelector2 = (_Closure_0024__.A = [SpecialName] (J<float, float, IEnumerable<WB>> A) => A.Group.Count());
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
						keySelector2 = _Closure_0024__.A;
					}
					IOrderedEnumerable<J<float, float, IEnumerable<WB>>> source3 = source2.OrderByDescending(keySelector2);
					Func<J<float, float, IEnumerable<WB>>, K<float, float, int>> selector;
					if (_Closure_0024__.A == null)
					{
						selector = (_Closure_0024__.A = [SpecialName] (J<float, float, IEnumerable<WB>> A) => new K<float, float, int>(A.left, A.first, A.Group.Count()));
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
						selector = _Closure_0024__.A;
					}
					IEnumerable<K<float, float, int>> enumerable = source3.Select(selector);
					using (IEnumerator<K<float, float, int>> enumerator2 = enumerable.GetEnumerator())
					{
						while (enumerator2.MoveNext())
						{
							K<float, float, int> current = enumerator2.Current;
							list3.Add(new BulletIndentFix
							{
								LeftIndent = current.l,
								FirstLineIndent = current.f
							});
							list4.Add(current.l.ToString(AH.A(14327)) + AH.A(14334) + current.f.ToString(AH.A(14327)) + AH.A(14365) + current.cnt + AH.A(14255));
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_02f4;
							}
							continue;
							end_IL_02f4:
							break;
						}
					}
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.BulletIndentation(sld, shp, list4, string.Join(AH.A(14258), list4.ToArray()), list2, list3));
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
		Indentation.Clear();
	}
}
