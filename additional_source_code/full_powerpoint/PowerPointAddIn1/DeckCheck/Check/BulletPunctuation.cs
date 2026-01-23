using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class BulletPunctuation : BaseCheck
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<bool, L<bool>> A;

		public static Func<bool, bool> A;

		public static Func<bool, bool> B;

		public static Func<bool, IEnumerable<bool>, M<bool, IEnumerable<bool>>> A;

		public static Func<M<bool, IEnumerable<bool>>, int> A;

		public static Func<M<bool, IEnumerable<bool>>, N<bool, int>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal L<bool> A(bool A)
		{
			return new L<bool>(A);
		}

		[SpecialName]
		internal bool A(bool A)
		{
			return A;
		}

		[SpecialName]
		internal bool B(bool A)
		{
			return A;
		}

		[SpecialName]
		internal M<bool, IEnumerable<bool>> A(bool A, IEnumerable<bool> B)
		{
			return new M<bool, IEnumerable<bool>>(A, B);
		}

		[SpecialName]
		internal int A(M<bool, IEnumerable<bool>> A)
		{
			return A.Group.Count();
		}

		[SpecialName]
		internal N<bool, int> A(M<bool, IEnumerable<bool>> A)
		{
			return new N<bool, int>(A.punct, A.Group.Count());
		}
	}

	[CompilerGenerated]
	private List<bool> A;

	[CompilerGenerated]
	private List<TextRange2> A;

	private List<bool> Punctuation
	{
		[CompilerGenerated]
		get
		{
			return this.A;
		}
		[CompilerGenerated]
		set
		{
			this.A = value;
		}
	}

	private List<TextRange2> Paragraphs
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

	public BulletPunctuation()
	{
		Punctuation = new List<bool>();
		Paragraphs = new List<TextRange2>();
	}

	public void CheckParagraph(TextRange2 para, string strText)
	{
		if (strText.Length <= 0)
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
			if (Operators.CompareString(Strings.Right(strText, 4), AH.A(14408), TextCompare: false) == 0)
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
				Punctuation.Add(Operators.CompareString(Conversions.ToString(strText.Last()), AH.A(14417), TextCompare: false) == 0);
				Paragraphs.Add(para);
				return;
			}
		}
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		if (Punctuation.Count <= 0)
		{
			return;
		}
		checked
		{
			IEnumerator<N<bool, int>> enumerator = default(IEnumerator<N<bool, int>>);
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
				List<bool> punctuation = Punctuation;
				Func<bool, L<bool>> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (bool A) => new L<bool>(A));
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
					keySelector = _Closure_0024__.A;
				}
				if (punctuation.GroupBy(keySelector).ToList().Count > 1)
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
					int num = 0;
					int num2 = 0;
					List<bool> list = new List<bool>();
					List<string> list2 = new List<string>();
					List<bool> punctuation2 = Punctuation;
					Func<bool, bool> keySelector2;
					if (_Closure_0024__.A == null)
					{
						keySelector2 = (_Closure_0024__.A = [SpecialName] (bool A) => A);
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
					Func<bool, bool> elementSelector;
					if (_Closure_0024__.B == null)
					{
						elementSelector = (_Closure_0024__.B = [SpecialName] (bool A) => A);
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
						elementSelector = _Closure_0024__.B;
					}
					IEnumerable<N<bool, int>> enumerable = from A in punctuation2.GroupBy(keySelector2, elementSelector, [SpecialName] (bool A, IEnumerable<bool> B) => new M<bool, IEnumerable<bool>>(A, B))
						orderby A.Group.Count() descending
						select new N<bool, int>(A.punct, A.Group.Count());
					try
					{
						enumerator = enumerable.GetEnumerator();
						while (enumerator.MoveNext())
						{
							N<bool, int> current = enumerator.Current;
							list.Add(current.p);
							if (current.p)
							{
								list2.Add(AH.A(14420) + current.cnt + AH.A(14255));
								num += current.cnt;
							}
							else
							{
								list2.Add(AH.A(14467) + current.cnt + AH.A(14255));
								num2 += current.cnt;
							}
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_023f;
							}
							continue;
							end_IL_023f:
							break;
						}
					}
					finally
					{
						if (enumerator != null)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								enumerator.Dispose();
								break;
							}
						}
					}
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.BulletPunctuation(sld, shp, list2, string.Format(AH.A(14520), num, num2), Paragraphs.ToList(), list));
					list = null;
					list2 = null;
				}
				Punctuation.Clear();
				Paragraphs.Clear();
				return;
			}
		}
	}
}
