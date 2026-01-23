using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Macabacus_Word.TextOps.Redaction.Redactors;
using Macabacus_Word.TextOps.Redaction.Values;
using Macabacus_Word.Values;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.TextOps.Redaction.Process;

public sealed class RedactUtilities
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Range, bool> A;

		public static Func<Range, Range> A;

		public static Func<TextRange2, bool> A;

		public static Func<TextRange2, TextRange2> A;

		public static Func<Range, bool> B;

		public static Func<Range, Range> B;

		public static Func<TextRange2, bool> B;

		public static Func<TextRange2, TextRange2> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(Range A)
		{
			if (!A.Text.Contains(XC.A(18455)))
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
				if (!A.Text.Contains(XC.A(18461)))
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
					if (!A.Text.Contains(XC.A(44217)))
					{
						return Operators.CompareString(A.Text.Trim(), null, TextCompare: false) != 0;
					}
				}
			}
			return false;
		}

		[SpecialName]
		internal Range A(Range A)
		{
			return A;
		}

		[SpecialName]
		internal bool A(TextRange2 A)
		{
			if (!A.Text.Contains(XC.A(18455)) && !A.Text.Contains(XC.A(18461)) && !A.Text.Contains(XC.A(44217)))
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
						return Operators.CompareString(A.Text.Trim(), null, TextCompare: false) != 0;
					}
				}
			}
			return false;
		}

		[SpecialName]
		internal TextRange2 B(TextRange2 A)
		{
			return A;
		}

		[SpecialName]
		internal bool B(Range A)
		{
			return !A.Text.Contains(XC.A(18455));
		}

		[SpecialName]
		internal Range C(Range A)
		{
			return A;
		}

		[SpecialName]
		internal bool B(TextRange2 A)
		{
			return !A.Text.Contains(XC.A(18455));
		}

		[SpecialName]
		internal TextRange2 D(TextRange2 A)
		{
			return A;
		}
	}

	public static void RedactFloatingShapes(List<IShape> iShapes)
	{
		using List<IShape>.Enumerator enumerator = iShapes.GetEnumerator();
		while (enumerator.MoveNext())
		{
			RedactFloatingShape(enumerator.Current);
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
			return;
		}
	}

	public static bool ShowYesNoDialogue(System.Windows.Window owner, string strMessage)
	{
		DialogResult dialogResult = Forms.YesNoCancelMessage2(owner, strMessage, (YesNoDefault)1);
		bool result = default(bool);
		if (dialogResult != DialogResult.Cancel)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (dialogResult == DialogResult.Yes)
					{
						return true;
					}
					return result;
				}
			}
		}
		return false;
	}

	public static void ExpandAtInsertionPoint(Selection sel)
	{
		if (sel.Type != WdSelectionType.wdSelectionIP)
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
			if (sel.Start == checked(sel.Paragraphs[1].Range.End - 1))
			{
				object Unit = WdUnits.wdWord;
				object Count = -1;
				sel.MoveStart(ref Unit, ref Count);
			}
			else
			{
				object Count = WdUnits.wdWord;
				sel.Expand(ref Count);
			}
			return;
		}
	}

	public static bool DoesRangeContainParagraphMark(Range rng)
	{
		return rng.Text.Contains(XC.A(18455));
	}

	public static Find GetFindObject(Range rng, string strFind)
	{
		Find find = rng.Find;
		find.ClearFormatting();
		find.MatchWildcards = false;
		find.MatchWholeWord = true;
		find.MatchCase = false;
		find.MatchPrefix = false;
		find.MatchSuffix = false;
		find.Forward = true;
		find.Text = strFind;
		find.Wrap = WdFindWrap.wdFindStop;
		_ = null;
		return find;
	}

	public static void RedactInlineShape(IShape iShape)
	{
		if (iShape is InlineShapeValue)
		{
			while (true)
			{
				InlineShape inlineShape;
				switch (3)
				{
				case 0:
					break;
				default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						inlineShape = ((InlineShapeValue)iShape).InlineShape;
						if (inlineShape.Type != WdInlineShapeType.wdInlineShapePicture)
						{
							if (inlineShape.Type != (WdInlineShapeType)clsUtilities.SVG_WD_INLINE_SHAPE_TYPE)
							{
								if (inlineShape.Type == WdInlineShapeType.wdInlineShapeChart)
								{
									ChartRedactor.RedactInlineChart(inlineShape);
								}
								else if (inlineShape.Type == WdInlineShapeType.wdInlineShapeSmartArt)
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
									SmartArtRedactor.RedactAllRangesInSmartArt(inlineShape.SmartArt);
								}
								goto IL_0090;
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
						PictureRedactor.RedactInlinePicture(inlineShape);
						goto IL_0090;
					}
					IL_0090:
					inlineShape = null;
					return;
				}
			}
		}
		Microsoft.Office.Interop.Word.Shape shape = ((ShapeValue)iShape).Shape;
		if (clsUtilities.DoesAutoShapeHaveText(shape))
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
			AutoShapeRedactor.RedactEntireRangeInAutoshape(shape);
		}
		shape = null;
	}

	public static void RedactFloatingShape(IShape iShape)
	{
		ShapeValue shapeValue = (ShapeValue)iShape;
		Microsoft.Office.Interop.Word.Shape shape = shapeValue.Shape;
		try
		{
			if (IsPicture(shape))
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
				if (!shapeValue.IsInsideGroup)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							PictureRedactor.RedactFloatingPicture(shape);
							return;
						}
					}
				}
			}
			if (IsChart(shape) && !shapeValue.IsInsideGroup)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						ChartRedactor.RedactFloatingChart(shape);
						return;
					}
				}
			}
			if (IsSmartArt(shape))
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						SmartArtRedactor.RedactAllRangesInSmartArt(shape.SmartArt);
						return;
					}
				}
			}
			if (clsUtilities.DoesAutoShapeHaveText(shape))
			{
				AutoShapeRedactor.RedactEntireRangeInAutoshape(shape);
			}
		}
		finally
		{
			shape = null;
		}
	}

	public static void RedactTextInShape(IShape iShape, Range rng)
	{
		AutoShapeRedactor.RedactRangeAutoshape(((ShapeValue)iShape).Shape, rng);
	}

	public static void RedactText(Range rng, bool redactEntireLine)
	{
		if (redactEntireLine)
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
					TextRedactor.RedactLineOfWordsInterop(rng);
					return;
				}
			}
		}
		TextRedactor.RedactWordsInterop(rng);
	}

	public static bool IsSmartArt(Microsoft.Office.Interop.Word.Shape shp)
	{
		if (shp.Type == MsoShapeType.msoSmartArt)
		{
			return true;
		}
		return false;
	}

	public static bool IsPicture(Microsoft.Office.Interop.Word.Shape shp)
	{
		if (shp.Type != MsoShapeType.msoPicture)
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
			if (shp.Type != (MsoShapeType)clsUtilities.SVG_WD_MSO_SHAPE_TYPE)
			{
				return false;
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
		return true;
	}

	public static bool IsChart(Microsoft.Office.Interop.Word.Shape shp)
	{
		if (shp.Type == MsoShapeType.msoChart)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		return false;
	}

	public static bool IsPictureOrChart(Microsoft.Office.Interop.Word.Shape shp)
	{
		if (!IsPicture(shp))
		{
			if (!IsChart(shp))
			{
				return false;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
		}
		return true;
	}

	public static UndoRecord BeginRedactionProcess(string name, Microsoft.Office.Interop.Word.Application wdApp)
	{
		UndoRecord undoRecord = wdApp.UndoRecord;
		undoRecord.StartCustomRecord(name);
		wdApp.ScreenUpdating = false;
		return undoRecord;
	}

	public static void CompleteRedactionProcess(string name, bool isFindAndRedact, ref UndoRecord undo, SelectionValue selectionValue)
	{
		if (selectionValue.ShapeRange.Count > 0)
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
			Microsoft.Office.Interop.Word.ShapeRange shapeRange = selectionValue.ShapeRange;
			object Replace = RuntimeHelpers.GetObjectValue(Missing.Value);
			shapeRange.Select(ref Replace);
		}
		else if (!isFindAndRedact)
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
				Document activeDocument = selectionValue.WdApp.ActiveDocument;
				object Replace = selectionValue.SelStart;
				object End = selectionValue.SelEnd;
				activeDocument.Range(ref Replace, ref End).Select();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		selectionValue.WdApp.ActiveWindow.View.Type = selectionValue.ViewType;
		selectionValue.WdApp.ScreenUpdating = true;
		selectionValue.WdApp.ScreenRefresh();
		undo.EndCustomRecord();
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, name);
	}

	public static Range TrimRange(Range rng, bool trimBefore, bool trimAfter)
	{
		if (Operators.CompareString(rng.Text, null, TextCompare: false) == 0)
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
					return null;
				}
			}
		}
		checked
		{
			if (trimBefore)
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
				while (rng.Text.StartsWith(XC.A(18458)))
				{
					rng.Start++;
					if (rng.Text != null)
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
						return null;
					}
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
			if (trimAfter)
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
				while (rng.Text.EndsWith(XC.A(18458)))
				{
					rng.End--;
					if (rng.Text != null)
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
						return null;
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
			return rng;
		}
	}

	public static string ReplaceSpaceWithSoftReturn(string str)
	{
		if (str != null && str.EndsWith(XC.A(18458)))
		{
			return str.Substring(0, checked(str.Length - 1)) + XC.A(18461);
		}
		return str;
	}

	public static Dictionary<int, int> GetTextLineList(Range rng)
	{
		Dictionary<int, int> dictionary = new Dictionary<int, int>();
		rng = TrimRange(rng, trimBefore: true, trimAfter: true);
		if (rng == null)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return dictionary;
				}
			}
		}
		List<Range> list = A(rng.Words.Cast<Range>().ToList());
		if (list.Count < 1)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					return dictionary;
				}
			}
		}
		Range range = list.First();
		object objectValue = RuntimeHelpers.GetObjectValue(range.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
		int num = range.Start;
		int end = range.End;
		foreach (Range item in list)
		{
			object objectValue2 = RuntimeHelpers.GetObjectValue(item.get_Information(WdInformation.wdVerticalPositionRelativeToPage));
			object value = item.Start;
			_ = (object)item.End;
			if (Operators.ConditionalCompareObjectNotEqual(objectValue2, objectValue, TextCompare: false))
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
				Document document = rng.Document;
				object Start = num;
				object End = end;
				Range range2 = document.Range(ref Start, ref End);
				end = Conversions.ToInteger(End);
				num = Conversions.ToInteger(Start);
				Range range3 = range2;
				range3.Text = ReplaceSpaceWithSoftReturn(range3.Text);
				dictionary.Add(range3.Start, range3.End);
				num = Conversions.ToInteger(value);
				end = item.End;
				objectValue = RuntimeHelpers.GetObjectValue(objectValue2);
			}
			else
			{
				end = item.End;
			}
		}
		dictionary.Add(num, list.Last().End);
		return dictionary;
	}

	private static List<Range> A(List<Range> A)
	{
		List<Range> result;
		try
		{
			Func<Range, bool> predicate;
			if (_Closure_0024__.A == null)
			{
				predicate = (_Closure_0024__.A = [SpecialName] (Range range) =>
				{
					if (!range.Text.Contains(XC.A(18455)))
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
						if (!range.Text.Contains(XC.A(18461)))
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
							if (!range.Text.Contains(XC.A(44217)))
							{
								return Operators.CompareString(range.Text.Trim(), null, TextCompare: false) != 0;
							}
						}
					}
					return false;
				});
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
				predicate = _Closure_0024__.A;
			}
			result = (from result2 in A.Where(predicate)
				select (result2)).ToList();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = new List<Range>();
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static List<TextRange2> B(List<TextRange2> A)
	{
		List<TextRange2> result;
		try
		{
			IEnumerable<TextRange2> source = A.Where([SpecialName] (TextRange2 textRange) =>
			{
				if (!textRange.Text.Contains(XC.A(18455)) && !textRange.Text.Contains(XC.A(18461)) && !textRange.Text.Contains(XC.A(44217)))
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
							return Operators.CompareString(textRange.Text.Trim(), null, TextCompare: false) != 0;
						}
					}
				}
				return false;
			});
			Func<TextRange2, TextRange2> selector;
			if (_Closure_0024__.A == null)
			{
				selector = (_Closure_0024__.A = [SpecialName] (TextRange2 result2) => result2);
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
			result = source.Select(selector).ToList();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = new List<TextRange2>();
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static IEnumerable<Range> GetWordList(Range rng, bool trimList)
	{
		List<Range> list = new List<Range>();
		rng = TrimRange(rng, trimBefore: true, trimAfter: true);
		IEnumerable<Range> result;
		if (rng == null)
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
			result = list;
		}
		else
		{
			list = rng.Words.Cast<Range>().ToList();
			try
			{
				if (trimList)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						result = A(list);
						break;
					}
				}
				else
				{
					List<Range> source = list;
					Func<Range, bool> predicate;
					if (_Closure_0024__.B == null)
					{
						predicate = (_Closure_0024__.B = [SpecialName] (Range A) => !A.Text.Contains(XC.A(18455)));
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
						predicate = _Closure_0024__.B;
					}
					IEnumerable<Range> source2 = source.Where(predicate);
					Func<Range, Range> selector;
					if (_Closure_0024__.B == null)
					{
						selector = (_Closure_0024__.B = [SpecialName] (Range A) => A);
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
						selector = _Closure_0024__.B;
					}
					result = source2.Select(selector).ToList();
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				result = Enumerable.Empty<Range>();
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	public static IEnumerable<TextRange2> GetWordListCore(TextRange2 rng, bool trimList)
	{
		List<TextRange2> list = new List<TextRange2>();
		IEnumerable<TextRange2> result;
		if (rng == null)
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
			result = list;
		}
		else
		{
			list = rng.get_Words(-1, -1).Cast<TextRange2>().ToList();
			try
			{
				if (trimList)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						result = B(list);
						break;
					}
				}
				else
				{
					List<TextRange2> source = list;
					Func<TextRange2, bool> predicate;
					if (_Closure_0024__.B == null)
					{
						predicate = (_Closure_0024__.B = [SpecialName] (TextRange2 A) => !A.Text.Contains(XC.A(18455)));
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
						predicate = _Closure_0024__.B;
					}
					result = (from A in source.Where(predicate)
						select (A)).ToList();
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				result = Enumerable.Empty<TextRange2>();
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}
}
