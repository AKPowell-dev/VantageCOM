using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Macabacus_Word.TextOps.Redaction.Process;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.TextOps.Redaction.Redactors;

public sealed class TextRedactor
{
	public static void RedactLineOfWordsInterop(Range rng)
	{
		Dictionary<int, int> textLineList = RedactUtilities.GetTextLineList(rng);
		foreach (KeyValuePair<int, int> item in textLineList)
		{
			Document document = rng.Document;
			object Start = item.Key;
			object End = item.Value;
			RedactWordInterop(document.Range(ref Start, ref End));
		}
		textLineList = null;
	}

	public static void RedactWordsInterop(Range rng)
	{
		IEnumerable<Range> wordList = RedactUtilities.GetWordList(rng, trimList: true);
		IEnumerator<Range> enumerator = default(IEnumerator<Range>);
		try
		{
			enumerator = wordList.GetEnumerator();
			while (enumerator.MoveNext())
			{
				RedactWordInterop(enumerator.Current.Duplicate);
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
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		wordList = null;
	}

	public static void RedactWordInterop(Range rngWord)
	{
		A(rngWord);
		TextHighlighter.AddHighlight(rngWord, A(rngWord.Font));
	}

	public static void RedactWordsInAutoshape(Microsoft.Office.Interop.Word.Shape shp, Range rng)
	{
		IEnumerable<Range> wordList = RedactUtilities.GetWordList(rng, trimList: true);
		IEnumerator<Range> enumerator = default(IEnumerator<Range>);
		try
		{
			enumerator = wordList.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range current = enumerator.Current;
				A(current);
				TextHighlighter.AddHighlight(current, A(current.Font, shp.Fill.ForeColor.RGB));
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		wordList = null;
	}

	private static int A(Font A)
	{
		int num;
		int result;
		try
		{
			if (A.ColorIndex == WdColorIndex.wdAuto)
			{
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
					num = 0;
					break;
				}
			}
			else
			{
				num = A.TextColor.RGB;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = 0;
			ProjectData.ClearProjectError();
			goto IL_0043;
		}
		result = num;
		goto IL_0043;
		IL_0043:
		return result;
	}

	private static int A(Font A, int B)
	{
		int num;
		int result;
		try
		{
			num = A.TextColor.RGB;
			if (num == -16777216)
			{
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
					if (B != 0)
					{
						if (B != -16777216)
						{
							num = 0;
							break;
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
					num = 16777215;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = 0;
			ProjectData.ClearProjectError();
			goto IL_005c;
		}
		result = num;
		goto IL_005c;
		IL_005c:
		return result;
	}

	public static void RedactWordCore(TextRange2 rngWord)
	{
		checked
		{
			try
			{
				if (modFunctionsStr.IsBlank(rngWord.Text))
				{
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
						return;
					}
				}
				float boundWidth = rngWord.BoundWidth;
				float boundTop = rngWord.BoundTop;
				rngWord.Text = Strings.StrDup(rngWord.Text.Length, XC.A(19662));
				Font2 font = rngWord.Font;
				float boundWidth2 = rngWord.BoundWidth;
				float boundTop2 = rngWord.BoundTop;
				int num = 0;
				if (!(boundWidth2 > boundWidth))
				{
					if (!(boundTop2 > boundTop))
					{
						if (!(boundWidth2 < boundWidth))
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
							if (!(boundTop2 < boundTop))
							{
								goto IL_01fa;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								break;
							}
						}
						while (true)
						{
							font.Spacing += 0.5f;
							boundWidth2 = rngWord.BoundWidth;
							boundTop2 = rngWord.BoundTop;
							num++;
							if (boundWidth2 >= boundWidth)
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
								if (boundTop2 == boundTop)
								{
									break;
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
							if (num <= 20)
							{
								continue;
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									continue;
								}
								break;
							}
							break;
						}
						num = 0;
						while (boundWidth2 > boundWidth)
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
							if (num <= 10)
							{
								font.Spacing -= 0.05f;
								boundWidth2 = rngWord.BoundWidth;
								num++;
								continue;
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
							break;
						}
						goto IL_01fa;
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
				}
				while (true)
				{
					font.Spacing -= 0.5f;
					boundWidth2 = rngWord.BoundWidth;
					boundTop2 = rngWord.BoundTop;
					num++;
					if (boundWidth2 <= boundWidth)
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
						if (boundTop2 == boundTop)
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
							break;
						}
					}
					if (num <= 20)
					{
						continue;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
					break;
				}
				num = 0;
				while (boundWidth2 < boundWidth)
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
					if (num <= 10)
					{
						font.Spacing += 0.05f;
						boundWidth2 = rngWord.BoundWidth;
						num++;
						continue;
					}
					break;
				}
				goto IL_01fa;
				IL_01fa:
				Font2 font2 = font;
				float size = font2.Size;
				string name = font2.Name;
				int rGB = font2.Fill.ForeColor.RGB;
				font2.Highlight.RGB = rGB;
				font2.Size = size;
				font2.Name = name;
				font2.Fill.ForeColor.RGB = rGB;
				_ = null;
			}
			finally
			{
				Font2 font = null;
			}
		}
	}

	public static void RedactFirstOccurrence(Range rng, string strFind, Document doc)
	{
		Range range;
		object Start = (range = rng).Start;
		Range range2;
		object End = (range2 = rng).Start;
		Range range3 = doc.Range(ref Start, ref End);
		range2.Start = Conversions.ToInteger(End);
		range.Start = Conversions.ToInteger(Start);
		rng = range3;
		Find findObject = RedactUtilities.GetFindObject(rng, strFind);
		Find find = findObject;
		End = RuntimeHelpers.GetObjectValue(Missing.Value);
		Start = RuntimeHelpers.GetObjectValue(Missing.Value);
		object MatchWholeWord = RuntimeHelpers.GetObjectValue(Missing.Value);
		object MatchWildcards = RuntimeHelpers.GetObjectValue(Missing.Value);
		object MatchSoundsLike = RuntimeHelpers.GetObjectValue(Missing.Value);
		object MatchAllWordForms = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Forward = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Wrap = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Format = RuntimeHelpers.GetObjectValue(Missing.Value);
		object ReplaceWith = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Replace = RuntimeHelpers.GetObjectValue(Missing.Value);
		object MatchKashida = RuntimeHelpers.GetObjectValue(Missing.Value);
		object MatchDiacritics = RuntimeHelpers.GetObjectValue(Missing.Value);
		object MatchAlefHamza = RuntimeHelpers.GetObjectValue(Missing.Value);
		object MatchControl = RuntimeHelpers.GetObjectValue(Missing.Value);
		if (find.Execute(ref End, ref Start, ref MatchWholeWord, ref MatchWildcards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl))
		{
			RedactWordInterop(rng.Duplicate);
		}
		findObject = null;
	}

	private static void A(Range A)
	{
		object Cset = XC.A(18458);
		object Count = WdConstants.wdBackward;
		A.MoveEndWhile(ref Cset, ref Count);
		Count = XC.A(18461);
		Cset = WdConstants.wdBackward;
		A.MoveEndWhile(ref Count, ref Cset);
		int end = A.End;
		Range duplicate = A.Duplicate;
		duplicate.SetRange(end, end);
		float num = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
		float num2 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
		Font font;
		try
		{
			A.Text = Strings.StrDup(A.Text.Length, XC.A(19662));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			duplicate = null;
			font = null;
			ProjectData.ClearProjectError();
			return;
		}
		font = A.Font;
		duplicate.SetRange(end, end);
		float num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
		float num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
		num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
		num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
		int num5 = 0;
		checked
		{
			if (num4 != num2)
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
				while (true)
				{
					font.Spacing -= 0.5f;
					duplicate.SetRange(end, end);
					num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
					num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
					num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
					num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
					num5++;
					if (num4 == num2)
					{
						break;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_01ad;
						}
						continue;
						end_IL_01ad:
						break;
					}
					if (num5 <= 20)
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
						break;
					}
					break;
				}
			}
			if (num3 > num)
			{
				while (true)
				{
					font.Spacing -= 0.5f;
					duplicate.SetRange(end, end);
					num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
					num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
					num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
					num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
					num5++;
					if (num3 <= num)
					{
						break;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_0244;
						}
						continue;
						end_IL_0244:
						break;
					}
					if (num5 <= 20)
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
					break;
				}
				while (num3 < num)
				{
					if (num5 <= 10)
					{
						font.Spacing += 0.05f;
						duplicate.SetRange(end, end);
						num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
						num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
						num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
						num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
						num5++;
						if (num4 != num2)
						{
							font.Spacing -= 0.05f;
							break;
						}
						continue;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
					break;
				}
			}
			else
			{
				if (!(num3 < num))
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
					if (num4 == num2)
					{
						goto IL_048f;
					}
				}
				while (true)
				{
					font.Spacing += 0.5f;
					num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
					num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
					num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
					num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
					num5++;
					if (num4 != num2)
					{
						font.Spacing -= 0.5f;
						num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
						num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
						num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
						num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
						break;
					}
					if (num3 >= num)
					{
						if (num4 == num2)
						{
							break;
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
					if (num5 <= 20)
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
						break;
					}
					break;
				}
				num5 = 0;
				while (num3 > num)
				{
					if (num5 <= 10)
					{
						font.Spacing -= 0.05f;
						num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
						num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
						num3 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
						num4 = Conversions.ToSingle(duplicate.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
						num5++;
						continue;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
					break;
				}
			}
			goto IL_048f;
		}
		IL_048f:
		duplicate = null;
		font = null;
	}

	public static Range GetFirstWord(Range rng, Application wdApp)
	{
		IEnumerable<Range> wordList = RedactUtilities.GetWordList(rng, trimList: false);
		if (wordList.Any())
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Document activeDocument = wdApp.ActiveDocument;
					Range range;
					object Start = (range = wordList.ElementAtOrDefault(0)).Start;
					object End = A(wordList);
					Range result = activeDocument.Range(ref Start, ref End);
					range.Start = Conversions.ToInteger(Start);
					return result;
				}
				}
			}
		}
		return null;
	}

	private static int A(IEnumerable<Range> A)
	{
		int end = A.First().End;
		if (A.Count() == 1)
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
					return end;
				}
			}
		}
		IEnumerator<Range> enumerator = default(IEnumerator<Range>);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Range current = enumerator.Current;
				if (modFunctionsStr.IsBlank(current.Text.Trim()))
				{
					return end;
				}
				if (current.Text.EndsWith(XC.A(18458)))
				{
					return current.End;
				}
				end = current.End;
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_009b;
				}
				continue;
				end_IL_009b:
				break;
			}
		}
		finally
		{
			if (enumerator != null)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		return end;
	}
}
