using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

public sealed class Highlight
{
	private static WdColorIndex m_A;

	private static WdColor m_A;

	private static int? m_A;

	private static readonly WdLineWidth m_A = WdLineWidth.wdLineWidth225pt;

	private static readonly WdLineStyle m_A = WdLineStyle.wdLineStyleSingle;

	private static readonly int? m_B = int.MinValue;

	private static readonly float m_A = -2.1474836E+09f;

	private static readonly float m_B = 0f;

	[CompilerGenerated]
	private static Dictionary<Document, LinkHighlights> m_A;

	private static Dictionary<Document, LinkHighlights> Highlights
	{
		[CompilerGenerated]
		get
		{
			return Highlight.m_A;
		}
		[CompilerGenerated]
		set
		{
			Highlight.m_A = value;
		}
	} = null;

	public static void LoadColor(int i)
	{
		if (Highlight.m_A.HasValue)
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
			Remove();
		}
		if (i != 0)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (i != 1)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								Highlight.m_A = WdColorIndex.wdBrightGreen;
								Highlight.m_A = WdColor.wdColorBrightGreen;
								Highlight.m_A = Information.RGB(0, 255, 0);
								return;
							}
						}
					}
					Highlight.m_A = WdColorIndex.wdYellow;
					Highlight.m_A = WdColor.wdColorYellow;
					Highlight.m_A = Information.RGB(255, 255, 0);
					return;
				}
			}
		}
		Highlight.m_A = WdColorIndex.wdTurquoise;
		Highlight.m_A = WdColor.wdColorTurquoise;
		Highlight.m_A = Information.RGB(0, 255, 255);
	}

	public static void Add()
	{
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		Document activeDocument = application.ActiveDocument;
		UndoRecord undoRecord = application.UndoRecord;
		LinkHighlights B = null;
		bool flag = false;
		if (Highlights == null)
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
			Highlights = new Dictionary<Document, LinkHighlights>();
		}
		if (!Highlights.ContainsKey(activeDocument))
		{
			Highlights.Add(activeDocument, new LinkHighlights());
		}
		else
		{
			B = Highlights[activeDocument];
		}
		undoRecord.StartCustomRecord(XC.A(13375));
		application.ScreenUpdating = false;
		if (activeDocument.TrackRevisions)
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
			if (activeDocument.TrackFormatting)
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
				activeDocument.TrackFormatting = false;
				flag = true;
			}
		}
		try
		{
			_ = activeDocument.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
			foreach (Range storyRange in activeDocument.StoryRanges)
			{
				Range range = storyRange;
				do
				{
					if (B != null)
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
						Highlight.B(range, ref B);
					}
					Range a = range;
					Dictionary<Document, LinkHighlights> highlights;
					Document key;
					LinkHighlights B2 = (highlights = Highlights)[key = activeDocument];
					A(a, ref B2);
					highlights[key] = B2;
					range = range.NextStoryRange;
				}
				while (range != null);
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
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		if (flag)
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
			activeDocument.TrackFormatting = true;
		}
		application.ScreenUpdating = true;
		undoRecord.EndCustomRecord();
		A(application);
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(13406)).AddEventHandler(application, new ApplicationEvents4_DocumentBeforeSaveEventHandler(A));
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(13443)).AddEventHandler(application, new ApplicationEvents4_DocumentBeforeCloseEventHandler(A));
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)10, XC.A(13375));
		activeDocument = null;
		undoRecord = null;
		application = null;
		B = null;
	}

	private static void A(Range A, ref LinkHighlights B)
	{
		foreach (ContentControl contentControl in A.ContentControls)
		{
			if (!Common.IsLinked(contentControl))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			B.Text.Add(contentControl.Range, contentControl.Range.HighlightColorIndex);
			contentControl.Range.HighlightColorIndex = Highlight.m_A;
		}
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			enumerator2 = A.InlineShapes.GetEnumerator();
			while (enumerator2.MoveNext())
			{
				InlineShape inlineShape = (InlineShape)enumerator2.Current;
				if (!Common.IsLinked(inlineShape))
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
				if (Images.B(inlineShape))
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
					B.InlineGlowShapes.Add(inlineShape, new GlowBorder(inlineShape.Glow));
					Highlight.B(inlineShape.Glow);
				}
				else
				{
					B.InlineShapes.Add(inlineShape, new InlineShapeBorder(inlineShape));
					Highlight.B(inlineShape);
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_011b;
				}
				continue;
				end_IL_011b:
				break;
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
		if (Helpers.A(A))
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
			foreach (Microsoft.Office.Interop.Word.Shape item in A.ShapeRange)
			{
				Highlight.A(item, ref B);
			}
		}
		IEnumerator enumerator4 = default(IEnumerator);
		try
		{
			enumerator4 = A.Tables.GetEnumerator();
			while (enumerator4.MoveNext())
			{
				Table table = (Table)enumerator4.Current;
				if (!Common.IsLinked(table))
				{
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
				B.Tables.Add(table, new TableBorders(table));
				WdBorderType[] array = new WdBorderType[4]
				{
					WdBorderType.wdBorderTop,
					WdBorderType.wdBorderRight,
					WdBorderType.wdBorderBottom,
					WdBorderType.wdBorderLeft
				};
				foreach (WdBorderType index in array)
				{
					Border border = table.Borders[index];
					border.LineStyle = Highlight.m_A;
					border.LineWidth = Highlight.m_A;
					border.Color = Highlight.m_A;
					_ = null;
				}
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
		finally
		{
			if (enumerator4 is IDisposable)
			{
				while (true)
				{
					switch (4)
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

	private static void A(Microsoft.Office.Interop.Word.Shape A, ref LinkHighlights B)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (Common.IsLinked(A))
					{
						if (Images.B(A))
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									B.GlowShapes.Add(A, new GlowBorder(A.Glow));
									Highlight.B(A.Glow);
									return;
								}
							}
						}
						B.Shapes.Add(A, new ShapeBorder(A));
						Highlight.B(A);
					}
					return;
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Highlight.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current, ref B);
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
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
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

	public static void Remove(Document doc = null)
	{
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		UndoRecord undoRecord = application.UndoRecord;
		LinkHighlights value = null;
		bool flag = false;
		bool flag2 = false;
		if (doc == null)
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
			doc = application.ActiveDocument;
		}
		if (Highlights != null)
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
			flag = Highlights.TryGetValue(doc, out value);
		}
		undoRecord.StartCustomRecord(XC.A(13482));
		application.ScreenUpdating = false;
		if (doc.TrackRevisions)
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
			if (doc.TrackFormatting)
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
				doc.TrackFormatting = false;
				flag2 = true;
			}
		}
		try
		{
			_ = doc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = doc.StoryRanges.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range range = (Range)enumerator.Current;
					do
					{
						B(range, ref value);
						range = range.NextStoryRange;
					}
					while (range != null);
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_011b;
					}
					continue;
					end_IL_011b:
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (7)
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		if (flag2)
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
			doc.TrackFormatting = true;
		}
		application.ScreenUpdating = true;
		undoRecord.EndCustomRecord();
		if (flag)
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
			Highlights.Remove(doc);
			if (!Highlights.Any())
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
				MC.A(Highlights);
				Highlights = null;
				A(application);
			}
		}
		doc = null;
		undoRecord = null;
		application = null;
	}

	private static void B(Range A, ref LinkHighlights B)
	{
		IEnumerator enumerator = default(IEnumerator);
		WdColorIndex C = default(WdColorIndex);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
		if (B != null)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					{
						enumerator = A.ContentControls.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								ContentControl contentControl = (ContentControl)enumerator.Current;
								if (Common.IsLinked(contentControl))
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
									if (Highlight.A(contentControl))
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
										if (Highlight.A(B.Text, contentControl.Range, ref C))
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
											contentControl.Range.HighlightColorIndex = C;
											B.Text.Remove(contentControl.Range);
										}
										else
										{
											Highlight.A(contentControl);
										}
									}
								}
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_00cb;
								}
								continue;
								end_IL_00cb:
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
					try
					{
						enumerator2 = A.InlineShapes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							InlineShape inlineShape = (InlineShape)enumerator2.Current;
							if (Common.IsLinked(inlineShape))
							{
								if (Images.B(inlineShape))
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
									if (Highlight.A(inlineShape.Glow))
									{
										GlowBorder glowBorder = Highlight.A(B.InlineGlowShapes, inlineShape.Range);
										if (glowBorder != null)
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
											Highlight.A(inlineShape.Glow, glowBorder);
											B.InlineGlowShapes.Remove(inlineShape);
										}
										else
										{
											Highlight.A(inlineShape.Glow);
										}
									}
								}
								else if (Highlight.A(inlineShape))
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
									InlineShapeBorder inlineShapeBorder = Highlight.A(B.InlineShapes, inlineShape.Range);
									if (inlineShapeBorder != null)
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
										Highlight.A(inlineShape, inlineShapeBorder);
										B.InlineShapes.Remove(inlineShape);
									}
									else
									{
										Highlight.A(inlineShape);
									}
								}
							}
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									(enumerator2 as IDisposable).Dispose();
									goto end_IL_0214;
								}
								continue;
								end_IL_0214:
								break;
							}
						}
					}
					if (Helpers.A(A))
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
						try
						{
							enumerator3 = A.ShapeRange.GetEnumerator();
							while (enumerator3.MoveNext())
							{
								Highlight.B((Microsoft.Office.Interop.Word.Shape)enumerator3.Current, ref B);
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_026d;
								}
								continue;
								end_IL_026d:
								break;
							}
						}
						finally
						{
							if (enumerator3 is IDisposable)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										(enumerator3 as IDisposable).Dispose();
										goto end_IL_0282;
									}
									continue;
									end_IL_0282:
									break;
								}
							}
						}
					}
					try
					{
						enumerator4 = A.Tables.GetEnumerator();
						while (enumerator4.MoveNext())
						{
							Table table = (Table)enumerator4.Current;
							if (Common.IsLinked(table))
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
								if (Highlight.A(table))
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
									TableBorders tableBorders = Highlight.A(B.Tables, table.Range);
									if (tableBorders != null)
									{
										TableBorders tableBorders2 = tableBorders;
										Highlight.A(table, WdBorderType.wdBorderTop, tableBorders2.Top);
										Highlight.A(table, WdBorderType.wdBorderRight, tableBorders2.Right);
										Highlight.A(table, WdBorderType.wdBorderBottom, tableBorders2.Bottom);
										Highlight.A(table, WdBorderType.wdBorderLeft, tableBorders2.Left);
										tableBorders2 = null;
										B.Tables.Remove(table);
									}
									else
									{
										Highlight.A(table);
									}
								}
							}
						}
						return;
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
									break;
								default:
									(enumerator4 as IDisposable).Dispose();
									goto end_IL_0388;
								}
								continue;
								end_IL_0388:
								break;
							}
						}
					}
				}
			}
		}
		IEnumerator enumerator5 = default(IEnumerator);
		try
		{
			enumerator5 = A.ContentControls.GetEnumerator();
			while (enumerator5.MoveNext())
			{
				ContentControl contentControl2 = (ContentControl)enumerator5.Current;
				if (!Common.IsLinked(contentControl2))
				{
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
				if (!Highlight.A(contentControl2))
				{
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
				Highlight.A(contentControl2);
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_03ff;
				}
				continue;
				end_IL_03ff:
				break;
			}
		}
		finally
		{
			if (enumerator5 is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator5 as IDisposable).Dispose();
					break;
				}
			}
		}
		IEnumerator enumerator6 = default(IEnumerator);
		try
		{
			enumerator6 = A.InlineShapes.GetEnumerator();
			while (enumerator6.MoveNext())
			{
				InlineShape inlineShape2 = (InlineShape)enumerator6.Current;
				if (!Common.IsLinked(inlineShape2))
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
				if (Images.B(inlineShape2))
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
					if (!Highlight.A(inlineShape2.Glow))
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
					Highlight.A(inlineShape2.Glow);
				}
				else if (Highlight.A(inlineShape2))
				{
					Highlight.A(inlineShape2);
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_04bd;
				}
				continue;
				end_IL_04bd:
				break;
			}
		}
		finally
		{
			if (enumerator6 is IDisposable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					(enumerator6 as IDisposable).Dispose();
					break;
				}
			}
		}
		if (Helpers.A(A))
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
			IEnumerator enumerator7 = default(IEnumerator);
			try
			{
				enumerator7 = A.ShapeRange.GetEnumerator();
				while (enumerator7.MoveNext())
				{
					Highlight.A((Microsoft.Office.Interop.Word.Shape)enumerator7.Current);
				}
			}
			finally
			{
				if (enumerator7 is IDisposable)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						(enumerator7 as IDisposable).Dispose();
						break;
					}
				}
			}
		}
		IEnumerator enumerator8 = default(IEnumerator);
		try
		{
			enumerator8 = A.Tables.GetEnumerator();
			while (enumerator8.MoveNext())
			{
				Table table2 = (Table)enumerator8.Current;
				if (!Common.IsLinked(table2) || !Highlight.A(table2))
				{
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
				Highlight.A(table2);
			}
			while (true)
			{
				switch (6)
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
			if (enumerator8 is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator8 as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private static void A(ContentControl A)
	{
		A.Range.HighlightColorIndex = WdColorIndex.wdAuto;
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			if (!Common.IsLinked(A))
			{
				return;
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
				if (!Highlight.A(A))
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
					Highlight.A(A.Line);
					return;
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Highlight.A((Microsoft.Office.Interop.Word.Shape)enumerator.Current);
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
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

	private static void A(Microsoft.Office.Interop.Word.LineFormat A)
	{
		A.Weight = 0f;
		A.Visible = MsoTriState.msoFalse;
	}

	private static void A(InlineShape A)
	{
		A.Line.Visible = MsoTriState.msoFalse;
		_ = null;
		A.Range.Borders.Enable = 0;
		Borders borders = A.Borders;
		borders.OutsideLineWidth = (WdLineWidth)0;
		borders.OutsideColor = WdColor.wdColorAutomatic;
		borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
		_ = null;
	}

	private static void A(GlowFormat A)
	{
		A.Radius = Highlight.m_B;
		A.Color.RGB = Highlight.m_B.Value;
		A.Transparency = Highlight.m_A;
	}

	private static void A(Table A)
	{
		WdBorderType[] array = new WdBorderType[4]
		{
			WdBorderType.wdBorderTop,
			WdBorderType.wdBorderRight,
			WdBorderType.wdBorderBottom,
			WdBorderType.wdBorderLeft
		};
		foreach (WdBorderType index in array)
		{
			A.Borders[index].LineStyle = WdLineStyle.wdLineStyleNone;
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

	private static void B(Microsoft.Office.Interop.Word.Shape A, ref LinkHighlights B)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			if (!Common.IsLinked(A))
			{
				return;
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
				if (Images.B(A))
				{
					if (Highlight.A(A.Glow))
					{
						GlowBorder glowBorder = Highlight.A(B.GlowShapes, A.Anchor);
						if (glowBorder != null)
						{
							Highlight.A(A.Glow, glowBorder);
							B.GlowShapes.Remove(A);
						}
						else
						{
							Highlight.A(A.Glow);
						}
					}
					return;
				}
				if (!Highlight.A(A))
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
					ShapeBorder shapeBorder = Highlight.A(B.Shapes, A.Anchor);
					if (shapeBorder != null)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								Highlight.A(A, shapeBorder);
								B.Shapes.Remove(A);
								return;
							}
						}
					}
					Highlight.A(A.Line);
					return;
				}
			}
		}
		foreach (Microsoft.Office.Interop.Word.Shape groupItem in A.GroupItems)
		{
			Highlight.B(groupItem, ref B);
		}
	}

	private static void A(Microsoft.Office.Interop.Word.Shape A, ShapeBorder B)
	{
		Microsoft.Office.Interop.Word.LineFormat line = A.Line;
		if (B.Visible == MsoTriState.msoTrue)
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
			line.Visible = MsoTriState.msoTrue;
		}
		if (line.Visible == MsoTriState.msoTrue)
		{
			line.ForeColor.RGB = B.Color;
			line.Style = (MsoLineStyle)Conversions.ToInteger(Interaction.IIf(B.Style == MsoLineStyle.msoLineStyleMixed, MsoLineStyle.msoLineSingle, B.Style));
			line.Weight = Conversions.ToSingle(Interaction.IIf(B.Weight < 1f, 0, B.Weight));
		}
		if (B.Visible == MsoTriState.msoFalse)
		{
			line.Visible = MsoTriState.msoFalse;
		}
		line = null;
	}

	private static void A(InlineShape A, InlineShapeBorder B)
	{
		if (B.Visible == MsoTriState.msoTrue)
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
			A.Line.Visible = MsoTriState.msoTrue;
		}
		A.Range.Borders.Enable = 0 - (B.Enable ? 1 : 0);
		Borders borders = A.Borders;
		borders.OutsideLineStyle = B.OutsideLineStyle;
		borders.OutsideColor = B.OutsideColor;
		try
		{
			borders.OutsideLineWidth = B.OutsideLineWidth;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		borders = null;
		if (B.Visible != MsoTriState.msoFalse)
		{
			return;
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			A.Line.Visible = MsoTriState.msoFalse;
			return;
		}
	}

	private static void A(GlowFormat A, GlowBorder B)
	{
		A.Radius = B.Radius;
		A.Color.RGB = B.RBGColor;
		A.Transparency = Conversions.ToSingle(Interaction.IIf(B.Transparency < 0f, 0, B.Transparency));
		_ = null;
	}

	private static void A(Table A, WdBorderType B, TableBorder C)
	{
		Border border = A.Borders[B];
		border.Visible = C.Visible;
		if (border.Visible)
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
			border.Color = C.Color;
			border.LineStyle = C.Style;
			border.LineWidth = C.Weight;
		}
		border = null;
	}

	private static bool A(ContentControl A)
	{
		return A.Range.HighlightColorIndex == Highlight.m_A;
	}

	private static bool A(Microsoft.Office.Interop.Word.Shape A)
	{
		Microsoft.Office.Interop.Word.LineFormat line = A.Line;
		int rGB = line.ForeColor.RGB;
		int? a = Highlight.m_A;
		bool? obj;
		if (!a.HasValue)
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
			obj = null;
		}
		else
		{
			obj = rGB == a.GetValueOrDefault();
		}
		bool? flag = obj;
		bool? flag2 = obj;
		bool? flag3;
		if (flag2.HasValue)
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
			if (flag != true)
			{
				flag3 = false;
				goto IL_00a5;
			}
		}
		if (line.Weight != 3f)
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
			flag3 = false;
		}
		else
		{
			flag3 = flag;
		}
		goto IL_00a5;
		IL_00ee:
		bool? obj2;
		bool? flag4 = (bool?)obj2;
		return flag4.Value;
		IL_00a5:
		flag4 = flag3;
		flag = flag3;
		if (flag.HasValue)
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
			if (flag4 != true)
			{
				obj2 = false;
				goto IL_00ee;
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
		obj2 = (line.Style == MsoLineStyle.msoLineSingle) & flag4;
		goto IL_00ee;
	}

	private static bool A(InlineShape A)
	{
		Borders borders = A.Borders;
		bool? obj;
		if (borders.OutsideLineStyle == WdLineStyle.wdLineStyleSingle)
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
			if (borders.OutsideLineWidth == WdLineWidth.wdLineWidth225pt)
			{
				int outsideColor = (int)borders.OutsideColor;
				int? a = Highlight.m_A;
				obj = (a.HasValue ? new bool?(outsideColor == a.GetValueOrDefault()) : ((bool?)null));
				goto IL_006f;
			}
		}
		obj = false;
		goto IL_006f;
		IL_006f:
		bool? flag = obj;
		return flag.Value;
	}

	private static bool A(GlowFormat A)
	{
		int rGB = A.Color.RGB;
		int? a = Highlight.m_A;
		bool? flag2;
		bool? flag = (flag2 = (a.HasValue ? new bool?(rGB == a.GetValueOrDefault()) : ((bool?)null)));
		bool? flag3;
		if (flag.HasValue)
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
			if (flag2 != true)
			{
				flag3 = false;
				goto IL_00a0;
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
		if (A.Radius != 3f)
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
			flag3 = false;
		}
		else
		{
			flag3 = flag2;
		}
		goto IL_00a0;
		IL_00eb:
		bool? flag5;
		bool? flag4 = (bool?)flag5;
		return flag4.Value;
		IL_00a0:
		flag4 = flag3;
		flag2 = flag3;
		if (flag2.HasValue)
		{
			if (flag4 != true)
			{
				flag5 = false;
				goto IL_00eb;
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
		if (A.Transparency != 0f)
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
			flag5 = false;
		}
		else
		{
			flag5 = flag4;
		}
		goto IL_00eb;
	}

	private static bool A(Table A)
	{
		WdBorderType[] array = new WdBorderType[4]
		{
			WdBorderType.wdBorderTop,
			WdBorderType.wdBorderRight,
			WdBorderType.wdBorderBottom,
			WdBorderType.wdBorderLeft
		};
		int num = 0;
		while (num < array.Length)
		{
			WdBorderType index = array[num];
			Border border = A.Borders[index];
			if (border.LineStyle == Highlight.m_A && border.LineWidth == Highlight.m_A)
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
				if (border.Color == Highlight.m_A)
				{
					border = null;
					num = checked(num + 1);
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
			}
			return false;
		}
		return true;
	}

	private static void B(Microsoft.Office.Interop.Word.Shape A)
	{
		Microsoft.Office.Interop.Word.LineFormat line = A.Line;
		line.Visible = MsoTriState.msoTrue;
		line.Style = MsoLineStyle.msoLineSingle;
		line.Weight = 2f;
		line.ForeColor.RGB = Highlight.m_A.Value;
		_ = null;
	}

	private static void B(InlineShape A)
	{
		A.Line.Visible = MsoTriState.msoTrue;
		_ = null;
		A.Range.Borders.Enable = -1;
		Borders borders = A.Borders;
		borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
		borders.OutsideLineWidth = WdLineWidth.wdLineWidth225pt;
		borders.OutsideColor = (WdColor)Highlight.m_A.Value;
		_ = null;
	}

	private static void B(GlowFormat A)
	{
		A.Radius = 3f;
		A.Transparency = 0f;
		A.Color.RGB = Highlight.m_A.Value;
		_ = null;
	}

	private static bool A(Dictionary<Range, WdColorIndex> A, Range B, ref WdColorIndex C)
	{
		using (Dictionary<Range, WdColorIndex>.Enumerator enumerator = A.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				KeyValuePair<Range, WdColorIndex> current = enumerator.Current;
				if (!current.Key.IsEqual(B))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					C = current.Value;
					return true;
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_0051;
				}
				continue;
				end_IL_0051:
				break;
			}
		}
		return false;
	}

	private static ShapeBorder A(Dictionary<Microsoft.Office.Interop.Word.Shape, ShapeBorder> A, Range B)
	{
		using (Dictionary<Microsoft.Office.Interop.Word.Shape, ShapeBorder>.Enumerator enumerator = A.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				KeyValuePair<Microsoft.Office.Interop.Word.Shape, ShapeBorder> current = enumerator.Current;
				if (!current.Key.Anchor.IsEqual(B))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return current.Value;
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_0055;
				}
				continue;
				end_IL_0055:
				break;
			}
		}
		return null;
	}

	private static InlineShapeBorder A(Dictionary<InlineShape, InlineShapeBorder> A, Range B)
	{
		using (Dictionary<InlineShape, InlineShapeBorder>.Enumerator enumerator = A.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				KeyValuePair<InlineShape, InlineShapeBorder> current = enumerator.Current;
				if (!current.Key.Range.IsEqual(B))
				{
					continue;
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
					return current.Value;
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_0055;
				}
				continue;
				end_IL_0055:
				break;
			}
		}
		return null;
	}

	private static GlowBorder A(Dictionary<Microsoft.Office.Interop.Word.Shape, GlowBorder> A, Range B)
	{
		using (Dictionary<Microsoft.Office.Interop.Word.Shape, GlowBorder>.Enumerator enumerator = A.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				KeyValuePair<Microsoft.Office.Interop.Word.Shape, GlowBorder> current = enumerator.Current;
				if (current.Key.Anchor.IsEqual(B))
				{
					return current.Value;
				}
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
				break;
			}
		}
		return null;
	}

	private static GlowBorder A(Dictionary<InlineShape, GlowBorder> A, Range B)
	{
		using (Dictionary<InlineShape, GlowBorder>.Enumerator enumerator = A.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				KeyValuePair<InlineShape, GlowBorder> current = enumerator.Current;
				if (!current.Key.Range.IsEqual(B))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return current.Value;
				}
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_0051;
				}
				continue;
				end_IL_0051:
				break;
			}
		}
		return null;
	}

	private static TableBorders A(Dictionary<Table, TableBorders> A, Range B)
	{
		foreach (KeyValuePair<Table, TableBorders> item in A)
		{
			if (item.Key.Range.IsEqual(B))
			{
				return item.Value;
			}
		}
		return null;
	}

	private static void A(Microsoft.Office.Interop.Word.Application A)
	{
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(13406)).RemoveEventHandler(A, new ApplicationEvents4_DocumentBeforeSaveEventHandler(Highlight.A));
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(13443)).RemoveEventHandler(A, new ApplicationEvents4_DocumentBeforeCloseEventHandler(Highlight.A));
	}

	private static void A(Document A, ref bool B)
	{
		if (Highlights == null || !Highlights.ContainsKey(A))
		{
			return;
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
			if (A.Saved)
			{
				return;
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				MessageBoxResult messageBoxResult = MessageBox.Show(XC.A(13517), XC.A(2438), MessageBoxButton.YesNoCancel, MessageBoxImage.Exclamation);
				if (messageBoxResult != MessageBoxResult.Cancel)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							if (messageBoxResult != MessageBoxResult.Yes)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										if (messageBoxResult != MessageBoxResult.No)
										{
											while (true)
											{
												switch (1)
												{
												case 0:
													break;
												default:
													return;
												}
											}
										}
										Highlights.Remove(A);
										return;
									}
								}
							}
							Remove(A);
							return;
						}
					}
				}
				B = true;
				return;
			}
		}
	}

	private static void A(Document A, ref bool B, ref bool C)
	{
		bool flag;
		try
		{
			flag = Conversions.ToBoolean(NewLateBinding.LateGet(A, null, XC.A(13685), new object[0], null, null, null));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			flag = false;
			ProjectData.ClearProjectError();
		}
		if (flag)
		{
			return;
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
			if (Highlights == null)
			{
				return;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				if (!Highlights.ContainsKey(A))
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
					if (C)
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
						MessageBoxResult messageBoxResult = MessageBox.Show(XC.A(13706), XC.A(2438), MessageBoxButton.YesNoCancel, MessageBoxImage.Exclamation);
						if (messageBoxResult != MessageBoxResult.Cancel)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									if (messageBoxResult != MessageBoxResult.Yes)
									{
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												_ = 7;
												return;
											}
										}
									}
									Remove(A);
									return;
								}
							}
						}
						C = true;
						return;
					}
				}
			}
		}
	}
}
