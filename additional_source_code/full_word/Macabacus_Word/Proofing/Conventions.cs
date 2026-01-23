using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros.Proofing;
using Macabacus_Word.Colors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing;

public sealed class Conventions : Conventions
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Tuple<int, object>, int> A;

		public static Func<Tuple<int, object>, int> B;

		public static Func<Tuple<int, object>, int> C;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(Tuple<int, object> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal int B(Tuple<int, object> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal int C(Tuple<int, object> A)
		{
			return A.Item1;
		}
	}

	private List<Tuple<int, object>> m_A;

	private List<Tuple<int, object>> m_B;

	private List<Tuple<int, object>> C;

	public List<Tuple<int, object>> UsedFillColors
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
		}
	}

	public List<Tuple<int, object>> UsedBorderColors
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
		}
	}

	public List<Tuple<int, object>> UsedFontColors
	{
		get
		{
			return C;
		}
		set
		{
			C = value;
		}
	}

	public Conventions(Document doc, Settings options)
	{
		//IL_0166: Unknown result type (might be due to invalid IL or missing references)
		//IL_016b: Unknown result type (might be due to invalid IL or missing references)
		_ = doc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
		IEnumerator enumerator = doc.StoryRanges.GetEnumerator();
		try
		{
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Range range = (Range)enumerator.Current;
				do
				{
					A(range, options);
					try
					{
						WdStoryType storyType = range.StoryType;
						if ((uint)(storyType - 6) <= 5u)
						{
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
								if (range.ShapeRange.Count <= 0)
								{
									break;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									try
									{
										enumerator2 = range.ShapeRange.GetEnumerator();
										while (enumerator2.MoveNext())
										{
											if (((Microsoft.Office.Interop.Word.Shape)enumerator2.Current).TextFrame2.HasText != MsoTriState.msoTrue)
											{
												continue;
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
											A(range, options);
										}
										while (true)
										{
											switch (2)
											{
											case 0:
												break;
											default:
												goto end_IL_00e8;
											}
											continue;
											end_IL_00e8:
											break;
										}
									}
									finally
									{
										if (enumerator2 is IDisposable)
										{
											while (true)
											{
												switch (6)
												{
												case 0:
													continue;
												}
												(enumerator2 as IDisposable).Dispose();
												break;
											}
										}
									}
									break;
								}
								break;
							}
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					range = range.NextStoryRange;
				}
				while (range != null);
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
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_0145;
				}
				continue;
				end_IL_0145:
				break;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
		if ((int)((Settings)options).HyphenWordsInconsistent != 0)
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
			if (((Conventions)this).HyphenatedWords.Count > 0)
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
				IEnumerator enumerator3 = default(IEnumerator);
				try
				{
					enumerator3 = doc.StoryRanges.GetEnumerator();
					while (enumerator3.MoveNext())
					{
						Range range2 = (Range)enumerator3.Current;
						do
						{
							B(range2, ((Conventions)this).HyphenatedWords.Distinct().ToList());
							try
							{
								WdStoryType storyType2 = range2.StoryType;
								if ((uint)(storyType2 - 6) <= 5u)
								{
									while (true)
									{
										switch (4)
										{
										case 0:
											continue;
										}
										if (range2.ShapeRange.Count <= 0)
										{
											break;
										}
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											foreach (Microsoft.Office.Interop.Word.Shape item in range2.ShapeRange)
											{
												if (item.TextFrame2.HasText == MsoTriState.msoTrue)
												{
													B(range2, ((Conventions)this).HyphenatedWords.Distinct().ToList());
												}
											}
											break;
										}
										break;
									}
								}
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							range2 = range2.NextStoryRange;
						}
						while (range2 != null);
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_02bf;
						}
						continue;
						end_IL_02bf:
						break;
					}
				}
				finally
				{
					if (enumerator3 is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator3 as IDisposable).Dispose();
							break;
						}
					}
				}
				((Conventions)this).FilterHyphenWords();
			}
		}
		DeterminePaletteUsage();
		((Conventions)this).CleanUp();
	}

	private void A(Range A, Settings B)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.InlineShapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				InlineShape a = (InlineShape)enumerator.Current;
				this.A(a, B);
			}
			while (true)
			{
				switch (1)
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
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = A.ShapeRange.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape a2 = (Microsoft.Office.Interop.Word.Shape)enumerator2.Current;
				this.A(a2, B);
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_0092;
				}
				continue;
				end_IL_0092:
				break;
			}
		}
		finally
		{
			if (enumerator2 is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator2 as IDisposable).Dispose();
					break;
				}
			}
		}
		IEnumerator enumerator3 = A.Tables.GetEnumerator();
		try
		{
			while (enumerator3.MoveNext())
			{
				Table a3 = (Table)enumerator3.Current;
				this.A(a3, B);
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_00f2;
				}
				continue;
				end_IL_00f2:
				break;
			}
		}
		finally
		{
			IDisposable disposable = enumerator3 as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
		IEnumerator enumerator4 = default(IEnumerator);
		try
		{
			enumerator4 = A.Paragraphs.GetEnumerator();
			while (enumerator4.MoveNext())
			{
				Paragraph paragraph = (Paragraph)enumerator4.Current;
				this.B(paragraph.Range, B);
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_0151;
				}
				continue;
				end_IL_0151:
				break;
			}
		}
		finally
		{
			if (enumerator4 is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator4 as IDisposable).Dispose();
					break;
				}
			}
		}
		List<Tuple<int, object>> FontColors = UsedFontColors;
		Index.Font(A, A, ref FontColors);
		UsedFontColors = FontColors;
	}

	private void A(InlineShape A, Settings B)
	{
		List<Tuple<int, object>> FillColors = UsedFillColors;
		Index.Fill(A, A, ref FillColors);
		UsedFillColors = FillColors;
		FillColors = UsedBorderColors;
		Index.Border(A, A, ref FillColors);
		UsedBorderColors = FillColors;
	}

	private void A(Microsoft.Office.Interop.Word.Shape A, Settings B)
	{
		Microsoft.Office.Interop.Word.Shape shape = A;
		if (shape.Type != MsoShapeType.msoGroup)
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
			if (shape.HasSmartArt == MsoTriState.msoTrue)
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
				try
				{
					IEnumerator enumerator = shape.SmartArt.Nodes.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
							if (smartArtNode.TextFrame2.HasText != MsoTriState.msoTrue)
							{
								continue;
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
							this.A(smartArtNode.TextFrame2.TextRange, B);
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_009a;
							}
							continue;
							end_IL_009a:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
					SmartArt smartArt = shape.SmartArt;
					List<Tuple<int, object>> FontColors = UsedFontColors;
					List<Tuple<int, object>> FillColors = UsedBorderColors;
					List<Tuple<int, object>> BorderColors = UsedBorderColors;
					Macabacus_Word.Colors.Index.SmartArt(smartArt, blnFont: true, blnFill: true, blnBorder: true, ref FontColors, ref FillColors, ref BorderColors);
					UsedBorderColors = BorderColors;
					UsedBorderColors = FillColors;
					UsedFontColors = FontColors;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			else
			{
				MsoShapeType type = shape.Type;
				if (type == MsoShapeType.msoAutoShape || type == MsoShapeType.msoTextBox)
				{
					try
					{
						if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								this.A(shape.TextFrame2.TextRange, B);
								break;
							}
						}
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					if (!Base.IgnoreShapeType(A))
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
						List<Tuple<int, object>> BorderColors = UsedFillColors;
						Macabacus_Word.Colors.Index.Fill(A, A, ref BorderColors);
						UsedFillColors = BorderColors;
						BorderColors = UsedBorderColors;
						Macabacus_Word.Colors.Index.Border(A, A, ref BorderColors);
						UsedBorderColors = BorderColors;
						BorderColors = UsedFontColors;
						Macabacus_Word.Colors.Index.Font(A, A, ref BorderColors);
						UsedFontColors = BorderColors;
					}
				}
			}
		}
		else
		{
			int count = shape.GroupItems.Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				GroupShapes groupItems = shape.GroupItems;
				object Index = i;
				this.A(groupItems[ref Index], B);
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
		shape = null;
	}

	private void A(Table A, Settings B)
	{
		Table table = A;
		checked
		{
			if (table.Tables.Count == 0)
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
				int count = table.Rows.Count;
				int count2 = table.Columns.Count;
				int num = count;
				for (int i = 1; i <= num; i++)
				{
					int num2 = count2;
					for (int j = 1; j <= num2; j++)
					{
						this.B(table.Cell(i, j).Range, B);
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0070;
						}
						continue;
						end_IL_0070:
						break;
					}
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
				List<Tuple<int, object>> FontColors = UsedFontColors;
				List<Tuple<int, object>> FillColors = UsedBorderColors;
				List<Tuple<int, object>> BorderColors = UsedBorderColors;
				Index.Table(A, blnFont: true, blnFill: true, blnBorder: true, ref FontColors, ref FillColors, ref BorderColors);
				UsedBorderColors = BorderColors;
				UsedBorderColors = FillColors;
				UsedFontColors = FontColors;
			}
			else
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = table.Tables.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Table a = (Table)enumerator.Current;
						this.A(a, B);
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0105;
						}
						continue;
						end_IL_0105:
						break;
					}
				}
				finally
				{
					if (enumerator is IDisposable)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			table = null;
		}
	}

	private void A(TextRange2 A, Settings B)
	{
		int count = A.get_Paragraphs(-1, -1).Count;
		IEnumerator enumerator = default(IEnumerator);
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			string text = A.get_Paragraphs(i, -1).Text;
			try
			{
				enumerator = A.get_Paragraphs(i, -1).get_Runs(-1, -1).GetEnumerator();
				while (enumerator.MoveNext())
				{
					TextRange2 textRange = (TextRange2)enumerator.Current;
					try
					{
						((Conventions)this).FontFamilies.Add(textRange.Font.Name);
						((Conventions)this).FontSizes.Add(textRange.Font.Size);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (1)
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
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			this.A(text, B);
		}
	}

	private void B(Range A, Settings B)
	{
		string text = A.Text;
		try
		{
			((Conventions)this).FontFamilies.Add(A.Font.Name);
			((Conventions)this).FontSizes.Add(A.Font.Size);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		this.A(text, B);
	}

	private void A(string A, Settings B)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_003e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_006f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		if ((int)((Settings)B).PunctuationSpacingInconsistent == 0)
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
			if ((int)((Settings)B).PunctuationSpacingIncorrect == 0)
			{
				goto IL_0038;
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
		((Conventions)this).CheckPunctuationSpacing(A);
		goto IL_0038;
		IL_0038:
		if ((int)((Settings)B).HyphenWordsInconsistent != 0)
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
			((Conventions)this).CheckHyphenWords(A);
		}
		if ((int)((Settings)B).MillionsBillionsAbbreviation != 0)
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
			((Conventions)this).CheckAbbreviations(A);
		}
		if ((int)((Settings)B).QuotesStyle == 0)
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
			((Conventions)this).CheckQuotesStyle(A);
			return;
		}
	}

	private void A(Range A, List<string> B)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.InlineShapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				InlineShape a = (InlineShape)enumerator.Current;
				this.A(a, B);
			}
		}
		finally
		{
			if (enumerator is IDisposable)
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
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = A.ShapeRange.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape a2 = (Microsoft.Office.Interop.Word.Shape)enumerator2.Current;
				this.A(a2, B);
			}
		}
		finally
		{
			if (enumerator2 is IDisposable)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					(enumerator2 as IDisposable).Dispose();
					break;
				}
			}
		}
		foreach (Table table in A.Tables)
		{
			this.A(table, B);
		}
		IEnumerator enumerator4 = default(IEnumerator);
		try
		{
			enumerator4 = A.Paragraphs.GetEnumerator();
			while (enumerator4.MoveNext())
			{
				Paragraph paragraph = (Paragraph)enumerator4.Current;
				this.B(paragraph.Range, B);
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
		finally
		{
			if (enumerator4 is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator4 as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private void A(InlineShape A, List<string> B)
	{
	}

	private void A(Microsoft.Office.Interop.Word.Shape A, List<string> B)
	{
		Microsoft.Office.Interop.Word.Shape shape = A;
		if (shape.Type != MsoShapeType.msoGroup)
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
			if (shape.HasSmartArt == MsoTriState.msoTrue)
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
				try
				{
					IEnumerator enumerator = shape.SmartArt.Nodes.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
							if (smartArtNode.TextFrame2.HasText == MsoTriState.msoTrue)
							{
								this.A(smartArtNode.TextFrame2.TextRange, B);
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_008c;
							}
							continue;
							end_IL_008c:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
			else
			{
				MsoShapeType type = shape.Type;
				if (type != MsoShapeType.msoAutoShape)
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
					if (type != MsoShapeType.msoTextBox)
					{
						goto IL_017d;
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
				try
				{
					if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							this.A(shape.TextFrame2.TextRange, B);
							break;
						}
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
		}
		else
		{
			int count = shape.GroupItems.Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				GroupShapes groupItems = shape.GroupItems;
				object Index = i;
				this.A(groupItems[ref Index], B);
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		goto IL_017d;
		IL_017d:
		shape = null;
	}

	private void A(TextRange2 A, List<string> B)
	{
		int count = A.get_Paragraphs(-1, -1).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			this.A(A.get_Paragraphs(i, -1).Text, B);
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
			return;
		}
	}

	private void A(Table A, List<string> B)
	{
		Table table = A;
		checked
		{
			if (table.Tables.Count != 0)
			{
				{
					IEnumerator enumerator = table.Tables.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							Table a = (Table)enumerator.Current;
							this.A(a, B);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_00ca;
							}
							continue;
							end_IL_00ca:
							break;
						}
					}
					finally
					{
						IDisposable disposable = enumerator as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				int count = table.Rows.Count;
				int count2 = table.Columns.Count;
				int num = count;
				for (int i = 1; i <= num; i++)
				{
					int num2 = count2;
					for (int j = 1; j <= num2; j++)
					{
						this.B(table.Cell(i, j).Range, B);
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0074;
						}
						continue;
						end_IL_0074:
						break;
					}
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
			table = null;
		}
	}

	private void B(Range A, List<string> B)
	{
		this.A(A.Text, B);
	}

	private void A(string A, List<string> B)
	{
		MatchCollection matchCollection = base.WordsRegex.Matches(A);
		foreach (Match item in matchCollection)
		{
			foreach (string item2 in B)
			{
				if (Operators.CompareString(item.Groups[1].Value.ToLower(), item2.Replace(XC.A(6388), ""), TextCompare: false) != 0)
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				((Conventions)this).UnhyphenatedWords.Add(item.Groups[1].Value.ToLower());
			}
		}
		matchCollection = null;
	}

	public void DeterminePaletteUsage()
	{
		List<int> list = new List<int>();
		List<int> list2 = list;
		List<Tuple<int, object>> usedFillColors = UsedFillColors;
		Func<Tuple<int, object>, int> selector;
		if (_Closure_0024__.A == null)
		{
			selector = (_Closure_0024__.A = [SpecialName] (Tuple<int, object> A) => A.Item1);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			selector = _Closure_0024__.A;
		}
		list2.AddRange(usedFillColors.Select(selector));
		List<Tuple<int, object>> usedBorderColors = UsedBorderColors;
		Func<Tuple<int, object>, int> selector2;
		if (_Closure_0024__.B == null)
		{
			selector2 = (_Closure_0024__.B = [SpecialName] (Tuple<int, object> A) => A.Item1);
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
			selector2 = _Closure_0024__.B;
		}
		list2.AddRange(usedBorderColors.Select(selector2));
		list2.AddRange(UsedFontColors.Select([SpecialName] (Tuple<int, object> A) => A.Item1));
		_ = null;
		((Conventions)this).PopulatePalette(list);
		list = null;
	}
}
