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

public sealed class LineSpacing : BaseCheck
{
	private struct AC
	{
		public TextRange2 A;

		public int A;

		public float A;

		public float B;

		public float C;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<AC, R<float, float, float>> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal R<float, float, float> A(AC A)
		{
			return new R<float, float, float>(A.A, A.B, A.C);
		}
	}

	[CompilerGenerated]
	internal sealed class BC
	{
		public int A;

		public Func<AC, bool> A;

		public BC(BC A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(AC A)
		{
			return A.A == this.A;
		}
	}

	[CompilerGenerated]
	private List<AC> A;

	private List<AC> LineSpacings
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

	public LineSpacing()
	{
		LineSpacings = new List<AC>();
	}

	public void CheckParagraph(TextRange2 para)
	{
		AC item = new AC
		{
			A = para
		};
		ParagraphFormat2 paragraphFormat = para.ParagraphFormat;
		item.A = paragraphFormat.IndentLevel;
		item.A = paragraphFormat.SpaceBefore;
		item.B = paragraphFormat.SpaceAfter;
		item.C = paragraphFormat.SpaceWithin;
		paragraphFormat = null;
		LineSpacings.Add(item);
	}

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		BC a = default(BC);
		BC CS_0024_003C_003E8__locals5 = new BC(a);
		if (LineSpacings.Count <= 0)
		{
			return;
		}
		int num = 1;
		IEnumerator<IGrouping<R<float, float, float>, AC>> enumerator2 = default(IEnumerator<IGrouping<R<float, float, float>, AC>>);
		do
		{
			CS_0024_003C_003E8__locals5.A = num;
			try
			{
				List<AC> lineSpacings = LineSpacings;
				Func<AC, bool> predicate;
				if (CS_0024_003C_003E8__locals5.A != null)
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
					predicate = CS_0024_003C_003E8__locals5.A;
				}
				else
				{
					predicate = (CS_0024_003C_003E8__locals5.A = [SpecialName] (AC A) => A.A == CS_0024_003C_003E8__locals5.A);
				}
				List<AC> list = lineSpacings.Where(predicate).ToList();
				List<AC> source = list;
				Func<AC, R<float, float, float>> keySelector;
				if (_Closure_0024__.A == null)
				{
					keySelector = (_Closure_0024__.A = [SpecialName] (AC A) => new R<float, float, float>(A.A, A.B, A.C));
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
					keySelector = _Closure_0024__.A;
				}
				IEnumerable<IGrouping<R<float, float, float>, AC>> enumerable = source.GroupBy(keySelector);
				if (enumerable.Count() > 1)
				{
					List<TextRange2> list2 = new List<TextRange2>();
					foreach (AC item2 in list)
					{
						list2.Add(item2.A);
					}
					List<LineSpacingFix> list3 = new List<LineSpacingFix>();
					List<string> list4 = new List<string>();
					List<string> list5 = new List<string>();
					try
					{
						enumerator2 = enumerable.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							IGrouping<R<float, float, float>, AC> current = enumerator2.Current;
							LineSpacingFix item = default(LineSpacingFix);
							R<float, float, float> key = current.Key;
							item.SpaceBefore = key.Before;
							item.SpaceAfter = key.After;
							item.SpaceWithin = key.Within;
							list4.Add(Conversions.ToString(key.Before) + AH.A(14600) + Conversions.ToString(key.After) + AH.A(14600) + Conversions.ToString(key.Within) + AH.A(14611) + current.Count() + AH.A(14255));
							list5.Add(Conversions.ToString(key.Before) + AH.A(14622) + Conversions.ToString(key.After) + AH.A(14622) + Conversions.ToString(key.Within) + AH.A(14611) + current.Count() + AH.A(14255));
							key = null;
							list3.Add(item);
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_02b2;
							}
							continue;
							end_IL_02b2:
							break;
						}
					}
					finally
					{
						if (enumerator2 != null)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								enumerator2.Dispose();
								break;
							}
						}
					}
					Main.Analysis.Errors.Add(new LineSpacingInconsistent(sld, shp, list4, string.Join(AH.A(14258), list5.ToArray()), list2, list3));
					list3 = null;
					list4 = null;
					list2 = null;
					list5 = null;
				}
				list = null;
				enumerable = null;
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
		LineSpacings.Clear();
	}
}
