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

public sealed class BulletFontFamily : BaseCheck
{
	private struct UB
	{
		public int A;

		public string A;

		public TextRange2 A;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<UB, E<string>> A;

		public static Func<UB, string> A;

		public static Func<UB, UB> A;

		public static Func<string, IEnumerable<UB>, F<string, IEnumerable<UB>>> A;

		public static Func<F<string, IEnumerable<UB>>, int> A;

		public static Func<F<string, IEnumerable<UB>>, G<string, int>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal E<string> A(UB A)
		{
			return new E<string>(A.A);
		}

		[SpecialName]
		internal string A(UB A)
		{
			return A.A;
		}

		[SpecialName]
		internal UB A(UB A)
		{
			return A;
		}

		[SpecialName]
		internal F<string, IEnumerable<UB>> A(string A, IEnumerable<UB> B)
		{
			return new F<string, IEnumerable<UB>>(A, B);
		}

		[SpecialName]
		internal int A(F<string, IEnumerable<UB>> A)
		{
			return A.Group.Count();
		}

		[SpecialName]
		internal G<string, int> A(F<string, IEnumerable<UB>> A)
		{
			return new G<string, int>(A.family, A.Group.Count());
		}
	}

	[CompilerGenerated]
	internal sealed class VB
	{
		public int A;

		public Func<UB, bool> A;

		public VB(VB A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(UB A)
		{
			return A.A == this.A;
		}
	}

	[CompilerGenerated]
	private List<UB> A;

	private List<UB> Fonts
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

	public BulletFontFamily()
	{
		Fonts = new List<UB>();
	}

	public void CheckParagraph(TextRange2 para, string strText)
	{
		if (para.ParagraphFormat.Bullet.Type == MsoBulletType.msoBulletPicture)
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
			if (strText.Length <= 0)
			{
				return;
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				UB item = default(UB);
				ParagraphFormat2 paragraphFormat = para.ParagraphFormat;
				item.A = paragraphFormat.IndentLevel;
				item.A = paragraphFormat.Bullet.Font.Name;
				paragraphFormat = null;
				item.A = para;
				Fonts.Add(item);
				return;
			}
		}
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		VB a = default(VB);
		VB CS_0024_003C_003E8__locals2 = new VB(a);
		if (Fonts.Count <= 0)
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
			int num = 1;
			do
			{
				CS_0024_003C_003E8__locals2.A = num;
				try
				{
					List<UB> list = Fonts.Where([SpecialName] (UB A) => A.A == CS_0024_003C_003E8__locals2.A).ToList();
					List<UB> source = list;
					Func<UB, E<string>> keySelector;
					if (_Closure_0024__.A == null)
					{
						keySelector = (_Closure_0024__.A = [SpecialName] (UB A) => new E<string>(A.A));
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
						keySelector = _Closure_0024__.A;
					}
					if (source.GroupBy(keySelector).ToList().Count > 1)
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
						List<TextRange2> list2 = new List<TextRange2>();
						using (List<UB>.Enumerator enumerator = list.GetEnumerator())
						{
							while (enumerator.MoveNext())
							{
								list2.Add(enumerator.Current.A);
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_00f5;
								}
								continue;
								end_IL_00f5:
								break;
							}
						}
						List<string> list3 = new List<string>();
						List<string> list4 = new List<string>();
						List<UB> fonts = Fonts;
						Func<UB, string> keySelector2;
						if (_Closure_0024__.A == null)
						{
							keySelector2 = (_Closure_0024__.A = [SpecialName] (UB A) => A.A);
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
							keySelector2 = _Closure_0024__.A;
						}
						Func<UB, UB> elementSelector;
						if (_Closure_0024__.A == null)
						{
							elementSelector = (_Closure_0024__.A = [SpecialName] (UB A) => A);
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
						Func<string, IEnumerable<UB>, F<string, IEnumerable<UB>>> resultSelector;
						if (_Closure_0024__.A == null)
						{
							resultSelector = (_Closure_0024__.A = [SpecialName] (string A, IEnumerable<UB> B) => new F<string, IEnumerable<UB>>(A, B));
						}
						else
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
							resultSelector = _Closure_0024__.A;
						}
						IEnumerable<G<string, int>> enumerable = from A in fonts.GroupBy(keySelector2, elementSelector, resultSelector)
							orderby A.Group.Count() descending
							select new G<string, int>(A.family, A.Group.Count());
						using (IEnumerator<G<string, int>> enumerator2 = enumerable.GetEnumerator())
						{
							while (enumerator2.MoveNext())
							{
								G<string, int> current = enumerator2.Current;
								list3.Add(current.f);
								list4.Add(current.f + AH.A(14248) + current.cnt + AH.A(14255));
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_027e;
								}
								continue;
								end_IL_027e:
								break;
							}
						}
						Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.BulletFontFamily(sld, shp, list4, string.Join(AH.A(14258), list4.ToArray()), list2, list3));
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
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				Fonts.Clear();
				return;
			}
		}
	}
}
