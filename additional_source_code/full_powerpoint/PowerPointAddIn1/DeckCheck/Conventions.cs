using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.Proofing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Colors;
using PowerPointAddIn1.DeckCheck.Fix;
using PowerPointAddIn1.DeckCheck.Reformat;
using PowerPointAddIn1.Slides;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.DeckCheck;

public sealed class Conventions : Conventions
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Tuple<int, IndexedObject>, int> A;

		public static Func<Tuple<int, IndexedObject>, int> B;

		public static Func<Tuple<int, IndexedObject>, int> C;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(Tuple<int, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal int B(Tuple<int, IndexedObject> A)
		{
			return A.Item1;
		}

		[SpecialName]
		internal int C(Tuple<int, IndexedObject> A)
		{
			return A.Item1;
		}
	}

	[CompilerGenerated]
	private Dictionary<Slide, List<ShapeBounds>> m_A;

	[CompilerGenerated]
	private Dictionary<Slide, List<Microsoft.Office.Interop.PowerPoint.Shape>> m_A;

	[CompilerGenerated]
	private Dictionary<Slide, Microsoft.Office.Interop.PowerPoint.Shape> m_A;

	[CompilerGenerated]
	private Dictionary<Slide, List<int>> m_A;

	[CompilerGenerated]
	private int m_A;

	[CompilerGenerated]
	private int m_B;

	[CompilerGenerated]
	private int m_C;

	[CompilerGenerated]
	private int m_D;

	[CompilerGenerated]
	private List<Tuple<int, IndexedObject>> m_A;

	[CompilerGenerated]
	private List<Tuple<float, IndexedObject>> m_A;

	[CompilerGenerated]
	private List<Tuple<int, IndexedObject>> m_B;

	[CompilerGenerated]
	private List<Tuple<int, IndexedObject>> m_C;

	[CompilerGenerated]
	private List<Tuple<string, IndexedObject>> m_A;

	[CompilerGenerated]
	private List<Tuple<FontStyle, IndexedObject>> m_A;

	[CompilerGenerated]
	private List<Tuple<TextDecoration, IndexedObject>> m_A;

	[CompilerGenerated]
	private List<Tuple<int, IndexedObject>> m_D;

	[CompilerGenerated]
	private List<Tuple<float, IndexedObject>> m_B;

	[CompilerGenerated]
	private List<Tuple<Margins, IndexedObject>> m_A;

	[CompilerGenerated]
	private List<Tuple<Margins, IndexedObject>> m_B;

	[CompilerGenerated]
	private List<Tuple<LineSpacing, IndexedObject>> m_A;

	[CompilerGenerated]
	private List<Tuple<Indent, IndexedObject>> m_A;

	[CompilerGenerated]
	private List<Tuple<BulletStyle, IndexedObject>> m_A;

	[CompilerGenerated]
	private List<string> m_A;

	[CompilerGenerated]
	private int? m_A;

	[CompilerGenerated]
	private int? m_B;

	public Dictionary<Slide, List<ShapeBounds>> TextRangeBounds
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

	public Dictionary<Slide, List<Microsoft.Office.Interop.PowerPoint.Shape>> LayoutPlaceholders
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

	public Dictionary<Slide, Microsoft.Office.Interop.PowerPoint.Shape> FootnoteShapes
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

	public Dictionary<Slide, List<int>> FootnoteNumbers
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

	public int ContentSlidesCount
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

	public int SlideNumbersCount
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	public int SentenceCaseTitlesCount
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	public int TitleCaseTitlesCount
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	public List<Tuple<int, IndexedObject>> UsedFillColors
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

	public List<Tuple<float, IndexedObject>> UsedFillTransparencies
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

	public List<Tuple<int, IndexedObject>> UsedBorderColors
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	public List<Tuple<int, IndexedObject>> UsedFontColors
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	public List<Tuple<string, IndexedObject>> UsedFontFamilies
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

	public List<Tuple<FontStyle, IndexedObject>> UsedFontStyles
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

	public List<Tuple<TextDecoration, IndexedObject>> UsedTextDecorations
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

	public List<Tuple<int, IndexedObject>> UsedBorderDashStyles
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	public List<Tuple<float, IndexedObject>> UsedBorderWeights
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	public List<Tuple<Margins, IndexedObject>> UsedTextBoxMargins
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

	public List<Tuple<Margins, IndexedObject>> UsedCellMargins
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	public List<Tuple<LineSpacing, IndexedObject>> UsedParagraphSpacing
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

	public List<Tuple<Indent, IndexedObject>> UsedIndents
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

	public List<Tuple<BulletStyle, IndexedObject>> UsedBulletStyles
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

	public List<string> LegalFonts
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

	public int? MinFontSize
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

	public int? MaxFontSize
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	public Conventions(Microsoft.Office.Interop.PowerPoint.Presentation pres, Settings options, List<string> unexpectedErrors)
	{
		//IL_011f: Unknown result type (might be due to invalid IL or missing references)
		//IL_013d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0142: Unknown result type (might be due to invalid IL or missing references)
		//IL_0205: Unknown result type (might be due to invalid IL or missing references)
		//IL_020a: Unknown result type (might be due to invalid IL or missing references)
		//IL_020e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0213: Unknown result type (might be due to invalid IL or missing references)
		//IL_0151: Unknown result type (might be due to invalid IL or missing references)
		//IL_0156: Unknown result type (might be due to invalid IL or missing references)
		//IL_0275: Unknown result type (might be due to invalid IL or missing references)
		//IL_027a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0474: Unknown result type (might be due to invalid IL or missing references)
		//IL_084c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0851: Unknown result type (might be due to invalid IL or missing references)
		//IL_0165: Unknown result type (might be due to invalid IL or missing references)
		//IL_016a: Unknown result type (might be due to invalid IL or missing references)
		//IL_05ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_05f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_0179: Unknown result type (might be due to invalid IL or missing references)
		//IL_017e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0183: Unknown result type (might be due to invalid IL or missing references)
		//IL_0188: Unknown result type (might be due to invalid IL or missing references)
		//IL_0197: Unknown result type (might be due to invalid IL or missing references)
		//IL_019c: Unknown result type (might be due to invalid IL or missing references)
		//IL_019f: Invalid comparison between Unknown and I4
		try
		{
			PowerPointAddIn1.Template.Settings settings = new PowerPointAddIn1.Template.Settings(pres);
			LegalFonts = settings.LegalFontTypes;
			MinFontSize = settings.MinFontSize;
			MaxFontSize = settings.MaxFontSize;
			settings = null;
			UsedFontColors = new List<Tuple<int, IndexedObject>>();
			UsedFontFamilies = new List<Tuple<string, IndexedObject>>();
			UsedFontStyles = new List<Tuple<FontStyle, IndexedObject>>();
			UsedTextDecorations = new List<Tuple<TextDecoration, IndexedObject>>();
			UsedFillColors = new List<Tuple<int, IndexedObject>>();
			UsedFillTransparencies = new List<Tuple<float, IndexedObject>>();
			UsedBorderColors = new List<Tuple<int, IndexedObject>>();
			UsedBorderDashStyles = new List<Tuple<int, IndexedObject>>();
			UsedBorderWeights = new List<Tuple<float, IndexedObject>>();
			UsedTextBoxMargins = new List<Tuple<Margins, IndexedObject>>();
			UsedCellMargins = new List<Tuple<Margins, IndexedObject>>();
			UsedParagraphSpacing = new List<Tuple<LineSpacing, IndexedObject>>();
			UsedIndents = new List<Tuple<Indent, IndexedObject>>();
			UsedBulletStyles = new List<Tuple<BulletStyle, IndexedObject>>();
			ContentSlidesCount = 0;
			SlideNumbersCount = 0;
			SentenceCaseTitlesCount = 0;
			TitleCaseTitlesCount = 0;
			TextRangeBounds = new Dictionary<Slide, List<ShapeBounds>>();
			LayoutPlaceholders = new Dictionary<Slide, List<Microsoft.Office.Interop.PowerPoint.Shape>>();
			FootnoteNumbers = new Dictionary<Slide, List<int>>();
			FootnoteShapes = new Dictionary<Slide, Microsoft.Office.Interop.PowerPoint.Shape>();
			int num;
			if ((int)options.CheckPlaceholderLayoutMismatch == 0)
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
				if ((int)options.CheckPlaceholderFillMismatch == 0)
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
					if ((int)options.CheckPlaceholderFontColorMismatch == 0)
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
						if ((int)options.CheckPlaceholderFontStyleMismatch == 0)
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
							if ((int)options.CheckPlaceholderMarginMismatch == 0 && (int)options.CheckPlaceholderIndentMismatch == 0)
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
								num = (((int)options.CheckPlaceholderBulletMismatch > 0) ? 1 : 0);
								goto IL_01a4;
							}
						}
					}
				}
			}
			num = 1;
			goto IL_01a4;
			IL_01a4:
			bool flag = (byte)num != 0;
			IEnumerator enumerator = pres.Slides.GetEnumerator();
			try
			{
				IEnumerator enumerator2 = default(IEnumerator);
				IEnumerator enumerator5 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					Slide slide = (Slide)enumerator.Current;
					List<ShapeBounds> list = new List<ShapeBounds>();
					List<int> list2 = new List<int>();
					List<Microsoft.Office.Interop.PowerPoint.Shape> list3 = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
					Microsoft.Office.Interop.PowerPoint.Shape shape = null;
					float height = slide.CustomLayout.Height;
					float width = slide.CustomLayout.Width;
					_ = options.CheckAgendaUpdated;
					if ((int)options.SlideNumbers != 0)
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
						checked
						{
							if (slide.SlideIndex > 1)
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
								if (!Helpers.IsSpecialSlide(slide))
								{
									ContentSlidesCount++;
									if (slide.HeadersFooters.SlideNumber.Visible == MsoTriState.msoTrue)
									{
										SlideNumbersCount++;
									}
								}
							}
						}
					}
					if ((int)options.SlideTitles != 0)
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
						checked
						{
							if (slide.Shapes.HasTitle == MsoTriState.msoTrue)
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
								Microsoft.Office.Interop.PowerPoint.Shape title = slide.Shapes.Title;
								if (title.HasTextFrame == MsoTriState.msoTrue && title.TextFrame2.HasText == MsoTriState.msoTrue)
								{
									string text = Strings.Trim(title.TextFrame2.TextRange.Text);
									if (text.Length > 0)
									{
										if (Text.CountCapsInTitle(text) == 0)
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
											SentenceCaseTitlesCount++;
										}
										else
										{
											TitleCaseTitlesCount++;
										}
									}
								}
								title = null;
							}
						}
					}
					try
					{
						enumerator2 = slide.CustomLayout.Shapes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
							if (flag)
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
								if (shape2.Type == MsoShapeType.msoPlaceholder)
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
									list3.Add(shape2);
								}
							}
							try
							{
								if (shape2.Type == MsoShapeType.msoPlaceholder)
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
									if (shape2.Visible != MsoTriState.msoTrue)
									{
										break;
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											continue;
										}
										if (shape2.HasTextFrame == MsoTriState.msoTrue && shape2.TextFrame2.HasText == MsoTriState.msoTrue)
										{
											A(slide, shape2, shape2.TextFrame2.TextRange);
										}
										break;
									}
									break;
								}
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
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0416;
							}
							continue;
							end_IL_0416:
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
					foreach (Microsoft.Office.Interop.PowerPoint.Shape shape5 in slide.Shapes)
					{
						A(slide, shape5, options);
						Microsoft.Office.Interop.PowerPoint.Shape shape4;
						if ((int)options.OverlappingText != 0)
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
							shape4 = shape5;
							MsoShapeType type = shape4.Type;
							if (type != MsoShapeType.msoAutoShape)
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
								if (type != MsoShapeType.msoPlaceholder)
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
									if (type != MsoShapeType.msoTextBox)
									{
										goto IL_05e9;
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
							}
							if (shape4.HasTextFrame == MsoTriState.msoTrue)
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
								if (shape4.TextFrame2.HasText == MsoTriState.msoTrue)
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
									foreach (TextRange2 item in shape4.TextFrame2.TextRange.get_Paragraphs(-1, -1))
									{
										if (Text.PrintableText(item.Text).Length > 0)
										{
											TextRange2 textRange2 = item;
											list.Add(new ShapeBounds
											{
												Top = textRange2.BoundTop,
												Right = textRange2.BoundLeft + textRange2.BoundWidth,
												Bottom = textRange2.BoundTop + textRange2.BoundHeight,
												Left = textRange2.BoundLeft,
												Zorder = shape5.ZOrderPosition
											});
											textRange2 = null;
										}
									}
								}
							}
							goto IL_05e9;
						}
						goto IL_05ec;
						IL_05ec:
						if ((int)options.Footnotes == 0 || shape != null)
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
						if (shape5.HasTextFrame != MsoTriState.msoTrue)
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
						if (!((double)shape5.Top >= 0.75 * (double)height))
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
						if (!((double)shape5.Width >= 0.5 * (double)width))
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
						if (!(shape5.TextFrame2.TextRange.Font.Size <= 9f))
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
						try
						{
							enumerator5 = shape5.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
							while (enumerator5.MoveNext())
							{
								TextRange2 textRange3 = (TextRange2)enumerator5.Current;
								if (textRange3.ParagraphFormat.Bullet.Type == MsoBulletType.msoBulletNumbered)
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
									list2.Add(textRange3.ParagraphFormat.Bullet.Number);
									continue;
								}
								Match match = Regex.Match(textRange3.Text, AH.A(47180));
								if (match.Success)
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
									list2.Add(Conversions.ToInteger(match.Groups[1].Value));
								}
								match = null;
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_0775;
								}
								continue;
								end_IL_0775:
								break;
							}
						}
						finally
						{
							if (enumerator5 is IDisposable)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									(enumerator5 as IDisposable).Dispose();
									break;
								}
							}
						}
						if (list2.Count <= 0)
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
						shape = shape5;
						continue;
						IL_05e9:
						shape4 = null;
						goto IL_05ec;
					}
					TextRangeBounds.Add(slide, list);
					LayoutPlaceholders.Add(slide, list3);
					FootnoteShapes.Add(slide, shape);
					FootnoteNumbers.Add(slide, list2);
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0829;
					}
					continue;
					end_IL_0829:
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
			if ((int)((Settings)options).HyphenWordsInconsistent != 0 && ((Conventions)this).HyphenatedWords.Count > 0)
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
				IEnumerator enumerator6 = default(IEnumerator);
				try
				{
					enumerator6 = pres.Slides.GetEnumerator();
					IEnumerator enumerator7 = default(IEnumerator);
					while (enumerator6.MoveNext())
					{
						Slide slide2 = (Slide)enumerator6.Current;
						try
						{
							enumerator7 = slide2.Shapes.GetEnumerator();
							while (enumerator7.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator7.Current;
								A(a, ((Conventions)this).HyphenatedWords.Distinct().ToList());
							}
						}
						finally
						{
							if (enumerator7 is IDisposable)
							{
								while (true)
								{
									switch (7)
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
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0912;
						}
						continue;
						end_IL_0912:
						break;
					}
				}
				finally
				{
					if (enumerator6 is IDisposable)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							(enumerator6 as IDisposable).Dispose();
							break;
						}
					}
				}
				((Conventions)this).FilterHyphenWords();
			}
			DeterminePaletteUsage();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			unexpectedErrors.Add(string.Format(AH.A(47235), AH.A(7894), ex4.Message));
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
		finally
		{
			List<ShapeBounds> list = null;
			List<int> list2 = null;
			List<Microsoft.Office.Interop.PowerPoint.Shape> list3 = null;
			Microsoft.Office.Interop.PowerPoint.Shape shape = null;
			((Conventions)this).CleanUp();
		}
	}

	private void A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, Settings C)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = B;
		if (shape.Type != MsoShapeType.msoGroup)
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
			if (shape.HasTable == MsoTriState.msoTrue)
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
				this.A(shape.Table, A, B, C);
			}
			else if (shape.HasChart == MsoTriState.msoTrue)
			{
				this.A(shape.Chart, A, B, C);
			}
			else
			{
				if (shape.HasSmartArt != MsoTriState.msoTrue)
				{
					MsoShapeType type = shape.Type;
					if (type != MsoShapeType.msoAutoShape)
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
						if (type != MsoShapeType.msoPlaceholder)
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
							if (type != MsoShapeType.msoTextBox)
							{
								goto IL_01a6;
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
					}
					if (shape.HasTextFrame == MsoTriState.msoTrue)
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
						if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
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
							this.A(shape.TextFrame2.TextRange, C);
						}
						if (shape.Type != MsoShapeType.msoPlaceholder)
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
							if (shape.Type != MsoShapeType.msoTextBox)
							{
								goto IL_01a6;
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
						Margins item = default(Margins);
						Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shape.TextFrame2;
						item.Top = textFrame.MarginTop;
						item.Right = textFrame.MarginRight;
						item.Bottom = textFrame.MarginBottom;
						item.Left = textFrame.MarginLeft;
						textFrame = null;
						UsedTextBoxMargins.Add(new Tuple<Margins, IndexedObject>(item, new IndexedObject(A, B, B)));
					}
					goto IL_01a6;
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
				this.A(shape.SmartArt, A, B, C);
			}
		}
		else
		{
			int count = shape.GroupItems.Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				this.A(A, shape.GroupItems[i], C);
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
		goto IL_0332;
		IL_01a6:
		if (!Base.A(B))
		{
			List<Tuple<int, IndexedObject>> FillColors = UsedFillColors;
			Index.Fill(A, B, B, ref FillColors);
			UsedFillColors = FillColors;
			FillColors = UsedBorderColors;
			Index.Border(A, B, B, ref FillColors);
			UsedBorderColors = FillColors;
			try
			{
				if (B.Fill.Visible == MsoTriState.msoTrue)
				{
					UsedFillTransparencies.Add(new Tuple<float, IndexedObject>(B.Fill.Transparency, new IndexedObject(A, B, B)));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			try
			{
				if (B.HasTextFrame == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						this.A(A, B, B.TextFrame2.TextRange);
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
			try
			{
				Microsoft.Office.Interop.PowerPoint.LineFormat line = B.Line;
				if (line.Visible == MsoTriState.msoTrue)
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
					UsedBorderDashStyles.Add(new Tuple<int, IndexedObject>((int)line.DashStyle, new IndexedObject(A, B, B)));
					UsedBorderWeights.Add(new Tuple<float, IndexedObject>(line.Weight, new IndexedObject(A, B, B)));
				}
				line = null;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
		}
		goto IL_0332;
		IL_0332:
		shape = null;
	}

	private void A(Table A, Slide B, Microsoft.Office.Interop.PowerPoint.Shape C, Settings D)
	{
		Table table = A;
		int count = table.Rows.Count;
		int count2 = table.Columns.Count;
		int num = count;
		checked
		{
			for (int i = 1; i <= num; i++)
			{
				int num2 = count2;
				for (int j = 1; j <= num2; j++)
				{
					Cell cell = table.Cell(i, j);
					Microsoft.Office.Interop.PowerPoint.Shape shape = cell.Shape;
					Cell obj = A.Cell(i, j);
					List<Tuple<int, IndexedObject>> FillColors = UsedFillColors;
					Index.Fill(B, shape, obj, ref FillColors);
					UsedFillColors = FillColors;
					Microsoft.Office.Interop.PowerPoint.Shape shape2 = cell.Shape;
					Cell obj2 = A.Cell(i, j);
					FillColors = UsedBorderColors;
					Index.Border(B, shape2, obj2, ref FillColors);
					UsedBorderColors = FillColors;
					try
					{
						if (cell.Shape.Fill.Visible == MsoTriState.msoTrue)
						{
							UsedFillTransparencies.Add(new Tuple<float, IndexedObject>(C.Fill.Transparency, new IndexedObject(B, cell.Shape, A.Cell(i, j))));
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					Microsoft.Office.Interop.PowerPoint.Shape shape3 = cell.Shape;
					if (shape3.HasTextFrame == MsoTriState.msoTrue)
					{
						Margins item = default(Margins);
						Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shape3.TextFrame2;
						if (textFrame.HasText == MsoTriState.msoTrue)
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
							this.A(textFrame.TextRange, D);
						}
						this.A(B, C, textFrame.TextRange);
						item.Top = textFrame.MarginTop;
						item.Right = textFrame.MarginRight;
						item.Bottom = textFrame.MarginBottom;
						item.Left = textFrame.MarginLeft;
						textFrame = null;
						UsedCellMargins.Add(new Tuple<Margins, IndexedObject>(item, new IndexedObject(B, A.Cell(i, j).Shape, A.Cell(i, j))));
					}
					shape3 = null;
					cell = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_01ec;
					}
					continue;
					end_IL_01ec:
					break;
				}
			}
			table = null;
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Chart A, Slide B, Microsoft.Office.Interop.PowerPoint.Shape C, Settings D)
	{
		Microsoft.Office.Interop.PowerPoint.Chart chart = A;
		if (!(chart.ChartArea.Format.Fill.Transparency < 0f))
		{
			Microsoft.Office.Interop.PowerPoint.FillFormat fill = chart.ChartArea.Format.Fill;
			Microsoft.Office.Interop.PowerPoint.ChartArea chartArea = A.ChartArea;
			List<Tuple<int, IndexedObject>> listColors = UsedFillColors;
			Index.FillFormat(fill, B, C, chartArea, ref listColors);
			UsedFillColors = listColors;
		}
		if (!(chart.PlotArea.Format.Fill.Transparency < 0f))
		{
			Microsoft.Office.Interop.PowerPoint.FillFormat fill2 = chart.PlotArea.Format.Fill;
			PlotArea plotArea = A.PlotArea;
			List<Tuple<int, IndexedObject>> listColors = UsedFillColors;
			Index.FillFormat(fill2, B, C, plotArea, ref listColors);
			UsedFillColors = listColors;
		}
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat line = chart.ChartArea.Format.Line;
			if (line.Visible == MsoTriState.msoTrue)
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
				if (line.Weight > 0f)
				{
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor = line.ForeColor;
					Microsoft.Office.Interop.PowerPoint.ChartArea chartArea2 = A.ChartArea;
					List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
					Index.AddColorIfNotTransparent(foreColor, B, C, chartArea2, ref listColors);
					UsedBorderColors = listColors;
				}
			}
			line = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			Microsoft.Office.Interop.PowerPoint.LineFormat line2 = chart.PlotArea.Format.Line;
			if (line2.Visible == MsoTriState.msoTrue)
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
				if (line2.Weight > 0f)
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
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor2 = line2.ForeColor;
					PlotArea plotArea2 = A.PlotArea;
					List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
					Index.AddColorIfNotTransparent(foreColor2, B, C, plotArea2, ref listColors);
					UsedBorderColors = listColors;
				}
			}
			line2 = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		if (chart.HasTitle)
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
			ChartTitle chartTitle = chart.ChartTitle;
			try
			{
				ChartTitle chartTitle2 = chartTitle;
				List<Tuple<int, IndexedObject>> listColors;
				if (!(chartTitle2.Format.Fill.Transparency < 0f))
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
					Microsoft.Office.Interop.PowerPoint.FillFormat fill3 = chartTitle2.Format.Fill;
					ChartTitle obj = chartTitle;
					listColors = UsedFillColors;
					Index.FillFormat(fill3, B, C, obj, ref listColors);
					UsedFillColors = listColors;
				}
				Microsoft.Office.Interop.PowerPoint.LineFormat line3 = chartTitle2.Format.Line;
				ChartTitle obj2 = chartTitle;
				listColors = UsedBorderColors;
				Index.LineFormat(line3, B, C, obj2, ref listColors);
				UsedBorderColors = listColors;
				if (clsCharts.ImplsFont(chartTitle))
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
					Microsoft.Office.Interop.PowerPoint.ChartFont font = chartTitle2.Font;
					UsedFontColors.Add(Index.NewTuple(Conversions.ToInteger(font.Color), B, C, chartTitle));
					UsedFontFamilies.Add(new Tuple<string, IndexedObject>(Conversions.ToString(font.Name), new IndexedObject(B, C, chartTitle)));
					UsedFontStyles.Add(new Tuple<FontStyle, IndexedObject>(this.A(chartTitle.Font), new IndexedObject(B, C, chartTitle)));
					font = null;
					this.A(chartTitle2.Font, B, C, chartTitle);
				}
				chartTitle2 = null;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception a = ex5;
				Main.A(a, null, A);
				ProjectData.ClearProjectError();
			}
			chartTitle = null;
		}
		int count = ((ChartGroups)chart.ChartGroups(RuntimeHelpers.GetObjectValue(Missing.Value))).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			ChartGroup chartGroup = (ChartGroup)chart.ChartGroups(i);
			ChartGroup chartGroup2 = chartGroup;
			if (chartGroup2.HasUpDownBars)
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
				Microsoft.Office.Interop.PowerPoint.FillFormat fill4 = chartGroup2.UpBars.Format.Fill;
				Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor3 = fill4.ForeColor;
				UpBars upBars = chartGroup.UpBars;
				List<Tuple<int, IndexedObject>> listColors = UsedFillColors;
				Index.AddColorIfNotTransparent(foreColor3, B, C, upBars, ref listColors);
				UsedFillColors = listColors;
				Microsoft.Office.Interop.PowerPoint.ColorFormat backColor = fill4.BackColor;
				UpBars upBars2 = chartGroup.UpBars;
				listColors = UsedFillColors;
				Index.AddColorIfNotTransparent(backColor, B, C, upBars2, ref listColors);
				UsedFillColors = listColors;
				_ = null;
				Microsoft.Office.Interop.PowerPoint.FillFormat fill5 = chartGroup2.DownBars.Format.Fill;
				Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor4 = fill5.ForeColor;
				DownBars downBars = chartGroup.DownBars;
				listColors = UsedFillColors;
				Index.AddColorIfNotTransparent(foreColor4, B, C, downBars, ref listColors);
				UsedFillColors = listColors;
				Microsoft.Office.Interop.PowerPoint.ColorFormat backColor2 = fill5.BackColor;
				DownBars downBars2 = chartGroup.DownBars;
				listColors = UsedFillColors;
				Index.AddColorIfNotTransparent(backColor2, B, C, downBars2, ref listColors);
				UsedFillColors = listColors;
				_ = null;
				Microsoft.Office.Interop.PowerPoint.LineFormat line4 = chartGroup2.UpBars.Format.Line;
				Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor5 = line4.ForeColor;
				UpBars upBars3 = chartGroup.UpBars;
				listColors = UsedBorderColors;
				Index.AddColorIfNotTransparent(foreColor5, B, C, upBars3, ref listColors);
				UsedBorderColors = listColors;
				Microsoft.Office.Interop.PowerPoint.ColorFormat backColor3 = line4.BackColor;
				UpBars upBars4 = chartGroup.UpBars;
				listColors = UsedBorderColors;
				Index.AddColorIfNotTransparent(backColor3, B, C, upBars4, ref listColors);
				UsedBorderColors = listColors;
				_ = null;
				Microsoft.Office.Interop.PowerPoint.LineFormat line5 = chartGroup2.DownBars.Format.Line;
				Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor6 = line5.ForeColor;
				DownBars downBars3 = chartGroup.DownBars;
				listColors = UsedBorderColors;
				Index.AddColorIfNotTransparent(foreColor6, B, C, downBars3, ref listColors);
				UsedBorderColors = listColors;
				Microsoft.Office.Interop.PowerPoint.ColorFormat backColor4 = line5.BackColor;
				DownBars downBars4 = chartGroup.DownBars;
				listColors = UsedBorderColors;
				Index.AddColorIfNotTransparent(backColor4, B, C, downBars4, ref listColors);
				UsedBorderColors = listColors;
				_ = null;
			}
			if (chartGroup2.HasHiLoLines)
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
				try
				{
					Microsoft.Office.Interop.PowerPoint.LineFormat line6 = chartGroup2.HiLoLines.Format.Line;
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor7 = line6.ForeColor;
					HiLoLines hiLoLines = chartGroup.HiLoLines;
					List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
					Index.AddColorIfNotTransparent(foreColor7, B, C, hiLoLines, ref listColors);
					UsedBorderColors = listColors;
					Microsoft.Office.Interop.PowerPoint.ColorFormat backColor5 = line6.BackColor;
					HiLoLines hiLoLines2 = chartGroup.HiLoLines;
					listColors = UsedBorderColors;
					Index.AddColorIfNotTransparent(backColor5, B, C, hiLoLines2, ref listColors);
					UsedBorderColors = listColors;
					_ = null;
				}
				catch (Exception ex6)
				{
					ProjectData.SetProjectError(ex6);
					Exception ex7 = ex6;
					ProjectData.ClearProjectError();
				}
			}
			if (chartGroup2.HasDropLines)
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
					Microsoft.Office.Interop.PowerPoint.LineFormat line7 = chartGroup2.DropLines.Format.Line;
					Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor8 = line7.ForeColor;
					DropLines dropLines = chartGroup.DropLines;
					List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
					Index.AddColorIfNotTransparent(foreColor8, B, C, dropLines, ref listColors);
					UsedBorderColors = listColors;
					Microsoft.Office.Interop.PowerPoint.ColorFormat backColor6 = line7.BackColor;
					DropLines dropLines2 = chartGroup.DropLines;
					listColors = UsedBorderColors;
					Index.AddColorIfNotTransparent(backColor6, B, C, dropLines2, ref listColors);
					UsedBorderColors = listColors;
					_ = null;
				}
				catch (Exception ex8)
				{
					ProjectData.SetProjectError(ex8);
					Exception ex9 = ex8;
					ProjectData.ClearProjectError();
				}
			}
			if (Charts.HasRadarAxisLabels(chartGroup))
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
				try
				{
					int? c = Conversions.ToInteger(chartGroup2.RadarAxisLabels.Font.Color);
					TickLabels radarAxisLabels = chartGroup2.RadarAxisLabels;
					List<Tuple<int, IndexedObject>> listColors = UsedFontColors;
					Index.AddColorIntIfNotTransparent(c, B, C, radarAxisLabels, ref listColors, isMarker: false, areRadarLabels: true);
					UsedFontColors = listColors;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
			}
			chartGroup2 = null;
			chartGroup = null;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			bool flag = clsCharts.UsesPointsForSeriesClrs(A);
			bool flag2 = clsCharts.UsesLegendsForSeriesClrs(A);
			bool flag3 = clsCharts.UsesFormatFillForSeriesClrs(A);
			if (flag2)
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
				if (A.HasLegend)
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
						enumerator = ((IEnumerable)A.Legend.LegendEntries(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator.MoveNext())
						{
							object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
							IMsoLegendKey msoLegendKey;
							try
							{
								msoLegendKey = (IMsoLegendKey)NewLateBinding.LateGet(objectValue, null, AH.A(13177), new object[0], null, null, null);
								IMsoLegendKey obj3 = msoLegendKey;
								List<Tuple<int, IndexedObject>> listColors = UsedFillColors;
								Index.FillFormat(B, C, obj3, ref listColors, isMarker: false);
								UsedFillColors = listColors;
								IMsoLegendKey obj4 = msoLegendKey;
								listColors = UsedFillColors;
								Index.FillFormat(B, C, obj4, ref listColors, isMarker: true);
								UsedFillColors = listColors;
							}
							catch (Exception projectError2)
							{
								ProjectData.SetProjectError(projectError2);
								ProjectData.ClearProjectError();
							}
							msoLegendKey = null;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_07e6;
							}
							continue;
							end_IL_07e6:
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
			}
			enumerator2 = ((IEnumerable)chart.SeriesCollection(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
			try
			{
				while (enumerator2.MoveNext())
				{
					IMsoSeries msoSeries = (IMsoSeries)enumerator2.Current;
					try
					{
						if (Charts.ImplsAndHasErrorBars(msoSeries))
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								Microsoft.Office.Core.LineFormat line8 = msoSeries.ErrorBars.Format.Line;
								IMsoErrorBars errorBars = msoSeries.ErrorBars;
								List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
								Index.LineFormat(line8, B, C, errorBars, ref listColors);
								UsedBorderColors = listColors;
								break;
							}
						}
					}
					catch (Exception ex10)
					{
						ProjectData.SetProjectError(ex10);
						Exception a2 = ex10;
						Main.A(a2, null, A);
						ProjectData.ClearProjectError();
					}
					try
					{
						if (msoSeries.HasLeaderLines)
						{
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								Microsoft.Office.Core.LineFormat line9 = msoSeries.LeaderLines.Format.Line;
								IMsoLeaderLines leaderLines = msoSeries.LeaderLines;
								List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
								Index.LineFormat(line9, B, C, leaderLines, ref listColors);
								UsedBorderColors = listColors;
								break;
							}
						}
					}
					catch (Exception ex11)
					{
						ProjectData.SetProjectError(ex11);
						Exception a3 = ex11;
						Main.A(a3, null, A);
						ProjectData.ClearProjectError();
					}
					try
					{
						if (Charts.ImplsTrendLines(msoSeries))
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								try
								{
									enumerator3 = ((IEnumerable)msoSeries.Trendlines(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
									while (enumerator3.MoveNext())
									{
										IMsoTrendline msoTrendline = (IMsoTrendline)enumerator3.Current;
										Microsoft.Office.Core.LineFormat line10 = msoTrendline.Format.Line;
										List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
										Index.LineFormat(line10, B, C, msoTrendline, ref listColors);
										UsedBorderColors = listColors;
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											break;
										default:
											goto end_IL_09a9;
										}
										continue;
										end_IL_09a9:
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
												continue;
											}
											(enumerator3 as IDisposable).Dispose();
											break;
										}
									}
								}
								break;
							}
						}
					}
					catch (Exception ex12)
					{
						ProjectData.SetProjectError(ex12);
						Exception a4 = ex12;
						Main.A(a4, null, A);
						ProjectData.ClearProjectError();
					}
					if (flag3)
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
						Microsoft.Office.Core.FillFormat fill6 = msoSeries.Format.Fill;
						List<Tuple<int, IndexedObject>> listColors = UsedFillColors;
						Index.FillFormat(fill6, B, C, msoSeries, ref listColors);
						UsedFillColors = listColors;
						Microsoft.Office.Core.LineFormat line11 = msoSeries.Format.Line;
						listColors = UsedBorderColors;
						Index.LineFormat(line11, B, C, msoSeries, ref listColors);
						UsedBorderColors = listColors;
					}
					try
					{
						if (msoSeries.HasDataLabels)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								UsedFontColors.Add(Index.NewTuple(Conversions.ToInteger(((IMsoDataLabels)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Font.Color), B, C, RuntimeHelpers.GetObjectValue(msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value)))));
								break;
							}
						}
					}
					catch (NullReferenceException ex13)
					{
						ProjectData.SetProjectError(ex13);
						NullReferenceException a5 = ex13;
						if (!object.Equals((int)A.ChartType, 122))
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
							Main.A(a5, null, A);
						}
						ProjectData.ClearProjectError();
					}
					catch (Exception ex14)
					{
						ProjectData.SetProjectError(ex14);
						Exception a6 = ex14;
						Main.A(a6, -2147467259, A, new int[8] { 83, 84, 85, 86, 119, 123, 117, 120 });
						ProjectData.ClearProjectError();
					}
					if (!Charts.ImplsPoints(msoSeries))
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
					try
					{
						enumerator4 = ((IEnumerable)msoSeries.Points(RuntimeHelpers.GetObjectValue(Missing.Value))).GetEnumerator();
						while (enumerator4.MoveNext())
						{
							ChartPoint chartPoint = (ChartPoint)enumerator4.Current;
							IMsoChartFormat format = chartPoint.Format;
							if (!flag2)
							{
								try
								{
									List<Tuple<int, IndexedObject>> listColors;
									if (format.Fill.Visible == MsoTriState.msoTrue)
									{
										if (flag)
										{
											goto IL_0c50;
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
										if (format.Fill.ForeColor.RGB != msoSeries.Format.Fill.ForeColor.RGB)
										{
											goto IL_0c50;
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
										if (format.Fill.BackColor.RGB != msoSeries.Format.Fill.BackColor.RGB)
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
											goto IL_0c50;
										}
									}
									else if (msoSeries.Format.Fill.Visible == MsoTriState.msoFalse)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											Microsoft.Office.Core.FillFormat fill7 = format.Fill;
											listColors = UsedFillColors;
											Index.FillFormat(fill7, B, C, chartPoint, ref listColors);
											UsedFillColors = listColors;
											break;
										}
									}
									goto end_IL_0bae;
									IL_0c50:
									Microsoft.Office.Core.FillFormat fill8 = format.Fill;
									listColors = UsedFillColors;
									Index.FillFormat(fill8, B, C, chartPoint, ref listColors);
									UsedFillColors = listColors;
									end_IL_0bae:;
								}
								catch (Exception ex15)
								{
									ProjectData.SetProjectError(ex15);
									Exception ex16 = ex15;
									ProjectData.ClearProjectError();
								}
							}
							try
							{
								if (format.Line.Visible == MsoTriState.msoTrue)
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											continue;
										}
										if (!flag)
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
											if (format.Line.ForeColor.RGB == msoSeries.Format.Line.ForeColor.RGB)
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
												if (format.Line.BackColor.RGB == msoSeries.Format.Line.BackColor.RGB)
												{
													break;
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
										}
										Microsoft.Office.Core.LineFormat line12 = format.Line;
										List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
										Index.LineFormat(line12, B, C, chartPoint, ref listColors);
										UsedBorderColors = listColors;
										break;
									}
								}
								else if (msoSeries.Format.Line.Visible == MsoTriState.msoFalse)
								{
									Microsoft.Office.Core.LineFormat line13 = format.Line;
									List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
									Index.LineFormat(line13, B, C, chartPoint, ref listColors);
									UsedBorderColors = listColors;
								}
							}
							catch (Exception ex17)
							{
								ProjectData.SetProjectError(ex17);
								Exception ex18 = ex17;
								ProjectData.ClearProjectError();
							}
							format = null;
							try
							{
								if (chartPoint.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone && chartPoint.MarkerBackgroundColor != Base.TRANSPARENT)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										UsedFillColors.Add(Index.NewTuple(chartPoint.MarkerBackgroundColor, B, C, chartPoint, isMarker: true));
										break;
									}
								}
							}
							catch (Exception ex19)
							{
								ProjectData.SetProjectError(ex19);
								Exception ex20 = ex19;
								ProjectData.ClearProjectError();
							}
							try
							{
								if (chartPoint.MarkerStyle != XlMarkerStyle.xlMarkerStyleNone)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										if (chartPoint.MarkerForegroundColor != Base.TRANSPARENT)
										{
											UsedBorderColors.Add(Index.NewTuple(chartPoint.MarkerForegroundColor, B, C, chartPoint, isMarker: true));
										}
										break;
									}
								}
							}
							catch (Exception ex21)
							{
								ProjectData.SetProjectError(ex21);
								Exception ex22 = ex21;
								ProjectData.ClearProjectError();
							}
							try
							{
								IMsoDataLabel dataLabel = chartPoint.DataLabel;
								if (msoSeries.HasDataLabels)
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
									if (Operators.ConditionalCompareObjectNotEqual(dataLabel.Font.Color, ((IMsoDataLabels)msoSeries.DataLabels(RuntimeHelpers.GetObjectValue(Missing.Value))).Font.Color, TextCompare: false))
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
										UsedFontColors.Add(Index.NewTuple(Conversions.ToInteger(dataLabel.Font.Color), B, C, chartPoint.DataLabel));
									}
								}
								else
								{
									UsedFontColors.Add(Index.NewTuple(Conversions.ToInteger(dataLabel.Font.Color), B, C, chartPoint.DataLabel));
								}
								if (dataLabel.Text.Length > 0)
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
									Microsoft.Office.Core.ChartFont font2 = dataLabel.Font;
									UsedFontFamilies.Add(new Tuple<string, IndexedObject>(Conversions.ToString(font2.Name), new IndexedObject(B, C, chartPoint.DataLabel)));
									UsedFontStyles.Add(new Tuple<FontStyle, IndexedObject>(this.A(chartPoint.DataLabel.Font), new IndexedObject(B, C, chartPoint.DataLabel)));
									font2 = null;
									this.A(dataLabel.Font, B, C, chartPoint.DataLabel);
								}
								dataLabel = null;
							}
							catch (Exception ex23)
							{
								ProjectData.SetProjectError(ex23);
								Exception ex24 = ex23;
								ProjectData.ClearProjectError();
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
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_1083;
					}
					continue;
					end_IL_1083:
					break;
				}
			}
			finally
			{
				IDisposable disposable = enumerator2 as IDisposable;
				if (disposable != null)
				{
					disposable.Dispose();
				}
			}
			using (List<Axis>.Enumerator enumerator5 = modCharts.AxesList(A).GetEnumerator())
			{
				while (enumerator5.MoveNext())
				{
					Axis current = enumerator5.Current;
					try
					{
						Axis axis = current;
						Microsoft.Office.Interop.PowerPoint.XlAxisType type = current.Type;
						_ = current.AxisGroup;
						if (axis.HasTitle)
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
							AxisTitle axisTitle = axis.AxisTitle;
							AxisTitle axisTitle2 = axisTitle;
							List<Tuple<int, IndexedObject>> listColors;
							if (!(axisTitle2.Format.Fill.Transparency < 0f))
							{
								Microsoft.Office.Interop.PowerPoint.FillFormat fill9 = axisTitle2.Format.Fill;
								AxisTitle obj5 = axisTitle;
								listColors = UsedFillColors;
								Index.FillFormat(fill9, B, C, obj5, ref listColors);
								UsedFillColors = listColors;
							}
							Microsoft.Office.Interop.PowerPoint.LineFormat line14 = axisTitle2.Format.Line;
							AxisTitle obj6 = axisTitle;
							listColors = UsedBorderColors;
							Index.LineFormat(line14, B, C, obj6, ref listColors);
							UsedBorderColors = listColors;
							Microsoft.Office.Interop.PowerPoint.ChartFont font3 = axisTitle2.Font;
							UsedFontColors.Add(Index.NewTuple(Conversions.ToInteger(font3.Color), B, C, axisTitle));
							UsedFontFamilies.Add(new Tuple<string, IndexedObject>(Conversions.ToString(font3.Name), new IndexedObject(B, C, axisTitle)));
							UsedFontStyles.Add(new Tuple<FontStyle, IndexedObject>(this.A(axisTitle.Font), new IndexedObject(B, C, axisTitle)));
							font3 = null;
							this.A(axisTitle2.Font, B, C, axisTitle);
							axisTitle2 = null;
							axisTitle = null;
						}
						if (axis.HasMajorGridlines)
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
							Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor9 = axis.MajorGridlines.Format.Line.ForeColor;
							Gridlines majorGridlines = axis.MajorGridlines;
							List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
							Index.AddColorIfNotTransparent(foreColor9, B, C, majorGridlines, ref listColors);
							UsedBorderColors = listColors;
						}
						if (axis.HasMinorGridlines)
						{
							Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor10 = axis.MinorGridlines.Format.Line.ForeColor;
							Gridlines minorGridlines = axis.MinorGridlines;
							List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
							Index.AddColorIfNotTransparent(foreColor10, B, C, minorGridlines, ref listColors);
							UsedBorderColors = listColors;
						}
						Microsoft.Office.Interop.PowerPoint.LineFormat line15 = axis.Format.Line;
						if (line15.Weight > 0f)
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
							if (line15.Visible == MsoTriState.msoTrue)
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
								Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor11 = line15.ForeColor;
								object objectValue2 = RuntimeHelpers.GetObjectValue(A.Axes(type));
								List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
								Index.AddColorIfNotTransparent(foreColor11, B, C, objectValue2, ref listColors);
								UsedBorderColors = listColors;
							}
						}
						line15 = null;
						try
						{
							if (axis.TickLabelPosition != XlTickLabelPosition.xlTickLabelPositionNone)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									TickLabels tickLabels = axis.TickLabels;
									Microsoft.Office.Interop.PowerPoint.ChartFont font4 = tickLabels.Font;
									UsedFontColors.Add(Index.NewTuple(Conversions.ToInteger(font4.Color), B, C, tickLabels));
									UsedFontFamilies.Add(new Tuple<string, IndexedObject>(Conversions.ToString(font4.Name), new IndexedObject(B, C, tickLabels)));
									UsedFontStyles.Add(new Tuple<FontStyle, IndexedObject>(this.A(tickLabels.Font), new IndexedObject(B, C, tickLabels)));
									font4 = null;
									this.A(tickLabels.Font, B, C, tickLabels);
									tickLabels = null;
									break;
								}
							}
						}
						catch (Exception ex25)
						{
							ProjectData.SetProjectError(ex25);
							Exception a7 = ex25;
							Main.A(a7, null, A, new int[5] { 118, 122, 121, 119, 123 });
							ProjectData.ClearProjectError();
						}
						axis = null;
					}
					catch (Exception ex26)
					{
						ProjectData.SetProjectError(ex26);
						Exception a8 = ex26;
						Main.A(a8, null, A);
						ProjectData.ClearProjectError();
					}
					finally
					{
						current = null;
					}
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_1463;
					}
					continue;
					end_IL_1463:
					break;
				}
			}
			if (chart.HasLegend)
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
				Legend legend = chart.Legend;
				try
				{
					Legend legend2 = legend;
					List<Tuple<int, IndexedObject>> listColors;
					if (!(legend2.Format.Fill.Transparency < 0f))
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
						Microsoft.Office.Interop.PowerPoint.FillFormat fill10 = legend2.Format.Fill;
						Legend obj7 = legend;
						listColors = UsedFillColors;
						Index.FillFormat(fill10, B, C, obj7, ref listColors);
						UsedFillColors = listColors;
					}
					Microsoft.Office.Interop.PowerPoint.LineFormat line16 = legend2.Format.Line;
					Legend obj8 = legend;
					listColors = UsedBorderColors;
					Index.LineFormat(line16, B, C, obj8, ref listColors);
					UsedBorderColors = listColors;
					Microsoft.Office.Interop.PowerPoint.ChartFont font5 = legend2.Font;
					UsedFontColors.Add(Index.NewTuple(Conversions.ToInteger(font5.Color), B, C, legend));
					UsedFontFamilies.Add(new Tuple<string, IndexedObject>(Conversions.ToString(font5.Name), new IndexedObject(B, C, legend)));
					UsedFontStyles.Add(new Tuple<FontStyle, IndexedObject>(this.A(legend.Font), new IndexedObject(B, C, legend)));
					font5 = null;
					this.A(legend2.Font, B, C, legend);
					legend2 = null;
				}
				catch (Exception ex27)
				{
					ProjectData.SetProjectError(ex27);
					Exception a9 = ex27;
					Main.A(a9, null, A);
					ProjectData.ClearProjectError();
				}
				legend = null;
			}
			if (chart.HasDataTable)
			{
				DataTable dataTable = chart.DataTable;
				try
				{
					DataTable dataTable2 = dataTable;
					Microsoft.Office.Interop.PowerPoint.FillFormat fill11 = dataTable2.Format.Fill;
					DataTable obj9 = dataTable;
					List<Tuple<int, IndexedObject>> listColors = UsedFillColors;
					Index.FillFormat(fill11, B, C, obj9, ref listColors);
					UsedFillColors = listColors;
					Microsoft.Office.Interop.PowerPoint.LineFormat line17 = dataTable2.Format.Line;
					DataTable obj10 = dataTable;
					listColors = UsedBorderColors;
					Index.LineFormat(line17, B, C, obj10, ref listColors);
					UsedBorderColors = listColors;
					Microsoft.Office.Interop.PowerPoint.ChartFont font6 = dataTable2.Font;
					UsedFontColors.Add(Index.NewTuple(Conversions.ToInteger(font6.Color), B, C, dataTable));
					UsedFontFamilies.Add(new Tuple<string, IndexedObject>(Conversions.ToString(font6.Name), new IndexedObject(B, C, dataTable)));
					UsedFontStyles.Add(new Tuple<FontStyle, IndexedObject>(this.A(dataTable.Font), new IndexedObject(B, C, dataTable)));
					font6 = null;
					this.A(dataTable2.Font, B, C, dataTable);
					dataTable2 = null;
				}
				catch (Exception ex28)
				{
					ProjectData.SetProjectError(ex28);
					Exception ex29 = ex28;
					clsReporting.LogException(ex29);
					ProjectData.ClearProjectError();
				}
				dataTable = null;
			}
			chart = null;
			return;
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.ChartFont A, Slide B, Microsoft.Office.Interop.PowerPoint.Shape C, object D)
	{
		Microsoft.Office.Interop.PowerPoint.ChartFont chartFont = A;
		if (!Conversions.ToBoolean(chartFont.Bold) && !Conversions.ToBoolean(chartFont.Italic))
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
			if ((XlUnderlineStyle)chartFont.Underline == XlUnderlineStyle.xlUnderlineStyleNone)
			{
				goto IL_0103;
			}
		}
		TextDecoration item = new TextDecoration
		{
			Bold = Conversions.ToBoolean(chartFont.Bold),
			Italic = Conversions.ToBoolean(chartFont.Italic)
		};
		XlUnderlineStyle xlUnderlineStyle = (XlUnderlineStyle)chartFont.Underline;
		if (xlUnderlineStyle != XlUnderlineStyle.xlUnderlineStyleNone)
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
			if (xlUnderlineStyle != XlUnderlineStyle.xlUnderlineStyleDouble)
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
				switch (xlUnderlineStyle)
				{
				case XlUnderlineStyle.xlUnderlineStyleSingle:
				case XlUnderlineStyle.xlUnderlineStyleSingleAccounting:
					item.UnderlineStyle = MsoTextUnderlineType.msoUnderlineSingleLine;
					goto IL_00e0;
				case XlUnderlineStyle.xlUnderlineStyleDoubleAccounting:
					break;
				default:
					goto IL_00e0;
				}
			}
			item.UnderlineStyle = MsoTextUnderlineType.msoUnderlineDoubleLine;
		}
		else
		{
			item.UnderlineStyle = MsoTextUnderlineType.msoNoUnderline;
		}
		goto IL_00e0;
		IL_00e0:
		UsedTextDecorations.Add(new Tuple<TextDecoration, IndexedObject>(item, new IndexedObject(B, C, RuntimeHelpers.GetObjectValue(D))));
		goto IL_0103;
		IL_0103:
		chartFont = null;
	}

	private void A(Microsoft.Office.Core.ChartFont A, Slide B, Microsoft.Office.Interop.PowerPoint.Shape C, object D)
	{
		Microsoft.Office.Core.ChartFont chartFont = A;
		if (!Conversions.ToBoolean(chartFont.Bold))
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
			if (!Conversions.ToBoolean(chartFont.Italic))
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
				if ((XlUnderlineStyle)chartFont.Underline == XlUnderlineStyle.xlUnderlineStyleNone)
				{
					goto IL_0113;
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
		}
		TextDecoration item = new TextDecoration
		{
			Bold = Conversions.ToBoolean(chartFont.Bold),
			Italic = Conversions.ToBoolean(chartFont.Italic)
		};
		XlUnderlineStyle xlUnderlineStyle = (XlUnderlineStyle)chartFont.Underline;
		if (xlUnderlineStyle != XlUnderlineStyle.xlUnderlineStyleNone)
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
			switch (xlUnderlineStyle)
			{
			case XlUnderlineStyle.xlUnderlineStyleSingle:
			case XlUnderlineStyle.xlUnderlineStyleSingleAccounting:
				item.UnderlineStyle = MsoTextUnderlineType.msoUnderlineSingleLine;
				break;
			case XlUnderlineStyle.xlUnderlineStyleDouble:
			case XlUnderlineStyle.xlUnderlineStyleDoubleAccounting:
				item.UnderlineStyle = MsoTextUnderlineType.msoUnderlineDoubleLine;
				break;
			}
		}
		else
		{
			item.UnderlineStyle = MsoTextUnderlineType.msoNoUnderline;
		}
		UsedTextDecorations.Add(new Tuple<TextDecoration, IndexedObject>(item, new IndexedObject(B, C, RuntimeHelpers.GetObjectValue(D))));
		goto IL_0113;
		IL_0113:
		chartFont = null;
	}

	private void A(SmartArt A, Slide B, Microsoft.Office.Interop.PowerPoint.Shape C, Settings D)
	{
		B.Application.StartNewUndoEntry();
		A.Reset();
		B.Application.CommandBars.ExecuteMso(AH.A(40491));
		System.Windows.Forms.Application.DoEvents();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Nodes.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
				try
				{
					if (smartArtNode.TextFrame2.HasText == MsoTriState.msoTrue)
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
						this.A(smartArtNode.TextFrame2.TextRange, D);
					}
					this.A(B, C, smartArtNode.TextFrame2.TextRange);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				try
				{
					enumerator2 = smartArtNode.Shapes.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Microsoft.Office.Core.Shape shape = (Microsoft.Office.Core.Shape)enumerator2.Current;
						try
						{
							Microsoft.Office.Core.FillFormat fill = shape.Fill;
							if (fill.Visible == MsoTriState.msoTrue)
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
								Microsoft.Office.Core.ColorFormat foreColor = fill.ForeColor;
								List<Tuple<int, IndexedObject>> listColors = UsedFillColors;
								Index.AddColorIfNotTransparent(foreColor, B, C, shape, ref listColors);
								UsedFillColors = listColors;
								Microsoft.Office.Core.ColorFormat backColor = fill.BackColor;
								listColors = UsedFillColors;
								Index.AddColorIfNotTransparent(backColor, B, C, shape, ref listColors);
								UsedFillColors = listColors;
							}
							fill = null;
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
						try
						{
							Microsoft.Office.Core.LineFormat line = shape.Line;
							if (line.Visible == MsoTriState.msoTrue)
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
								Microsoft.Office.Core.ColorFormat foreColor2 = line.ForeColor;
								List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
								Index.AddColorIfNotTransparent(foreColor2, B, C, shape, ref listColors);
								UsedBorderColors = listColors;
								Microsoft.Office.Core.ColorFormat backColor2 = line.BackColor;
								listColors = UsedBorderColors;
								Index.AddColorIfNotTransparent(backColor2, B, C, shape, ref listColors);
								UsedBorderColors = listColors;
							}
							line = null;
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							ProjectData.ClearProjectError();
						}
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_01e7;
						}
						continue;
						end_IL_01e7:
						break;
					}
				}
				finally
				{
					if (enumerator2 is IDisposable)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_021e;
				}
				continue;
				end_IL_021e:
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
		try
		{
			if (C.Fill.Visible == MsoTriState.msoTrue)
			{
				List<Tuple<int, IndexedObject>> listColors = UsedFillColors;
				Index.Fill(B, C, C, ref listColors);
				UsedFillColors = listColors;
			}
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		try
		{
			if (C.Line.Visible == MsoTriState.msoTrue)
			{
				List<Tuple<int, IndexedObject>> listColors = UsedBorderColors;
				Index.Border(B, C, C, ref listColors);
				UsedBorderColors = listColors;
			}
		}
		catch (Exception ex9)
		{
			ProjectData.SetProjectError(ex9);
			Exception ex10 = ex9;
			ProjectData.ClearProjectError();
		}
	}

	private void A(TextRange2 A, Settings B)
	{
		//IL_00d6: Unknown result type (might be due to invalid IL or missing references)
		//IL_0101: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_011a: Unknown result type (might be due to invalid IL or missing references)
		//IL_011f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0135: Unknown result type (might be due to invalid IL or missing references)
		//IL_013a: Unknown result type (might be due to invalid IL or missing references)
		int count = A.get_Paragraphs(-1, -1).Count;
		IEnumerator enumerator = default(IEnumerator);
		for (int num = 1; num <= count; num = checked(num + 1))
		{
			string text = A.get_Paragraphs(num, -1).Text;
			try
			{
				enumerator = A.get_Paragraphs(num, -1).get_Runs(-1, -1).GetEnumerator();
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
			if ((int)((Settings)B).PunctuationSpacingInconsistent == 0)
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
				if ((int)((Settings)B).PunctuationSpacingIncorrect == 0)
				{
					goto IL_0100;
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
			((Conventions)this).CheckPunctuationSpacing(text);
			goto IL_0100;
			IL_0100:
			if ((int)((Settings)B).HyphenWordsInconsistent != 0)
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
				((Conventions)this).CheckHyphenWords(text);
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
				((Conventions)this).CheckAbbreviations(text);
			}
			if ((int)((Settings)B).QuotesStyle != 0)
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
				((Conventions)this).CheckQuotesStyle(text);
			}
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, List<string> B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		checked
		{
			if (shape.Type != MsoShapeType.msoGroup)
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
				if (shape.HasTable == MsoTriState.msoTrue)
				{
					Table table = shape.Table;
					int count = table.Rows.Count;
					int count2 = table.Columns.Count;
					int num = count;
					for (int i = 1; i <= num; i++)
					{
						int num2 = count2;
						for (int j = 1; j <= num2; j++)
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape2 = table.Cell(i, j).Shape;
							if (shape2.HasTextFrame == MsoTriState.msoTrue)
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
								if (shape2.TextFrame2.HasText == MsoTriState.msoTrue)
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
									this.A(shape2.TextFrame2.TextRange, B);
								}
							}
							shape2 = null;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_00d0;
							}
							continue;
							end_IL_00d0:
							break;
						}
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
					table = null;
				}
				else if (shape.HasSmartArt == MsoTriState.msoTrue)
				{
					try
					{
						IEnumerator enumerator = default(IEnumerator);
						try
						{
							enumerator = shape.SmartArt.AllNodes.GetEnumerator();
							while (enumerator.MoveNext())
							{
								SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
								if (smartArtNode.TextFrame2.HasText != MsoTriState.msoTrue)
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
								this.A(smartArtNode.TextFrame2.TextRange, B);
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_016a;
								}
								continue;
								end_IL_016a:
								break;
							}
						}
						finally
						{
							if (enumerator is IDisposable)
							{
								while (true)
								{
									switch (5)
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
						ProjectData.ClearProjectError();
					}
				}
				else
				{
					MsoShapeType type = shape.Type;
					if (type != MsoShapeType.msoAutoShape && type != MsoShapeType.msoPlaceholder)
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
						if (type != MsoShapeType.msoTextBox)
						{
							goto IL_0270;
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
					if (shape.HasTextFrame == MsoTriState.msoTrue)
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
						if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
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
							this.A(shape.TextFrame2.TextRange, B);
						}
					}
				}
			}
			else
			{
				int count3 = shape.GroupItems.Count;
				for (int k = 1; k <= count3; k++)
				{
					this.A(shape.GroupItems[k], B);
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
			goto IL_0270;
		}
		IL_0270:
		shape = null;
	}

	private void A(TextRange2 A, List<string> B)
	{
		int count = A.get_Paragraphs(-1, -1).Count;
		for (int i = 1; i <= count; i = checked(i + 1))
		{
			MatchCollection matchCollection = base.WordsRegex.Matches(A.get_Paragraphs(i, -1).Text);
			foreach (Match item in matchCollection)
			{
				using List<string>.Enumerator enumerator2 = B.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					string current = enumerator2.Current;
					if (Operators.CompareString(item.Groups[1].Value.ToLower(), current.Replace(AH.A(17319), ""), TextCompare: false) == 0)
					{
						((Conventions)this).UnhyphenatedWords.Add(item.Groups[1].Value.ToLower());
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
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			MatchCollection matchCollection = null;
			return;
		}
	}

	private void A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, TextRange2 C)
	{
		IEnumerator enumerator = default(IEnumerator);
		if (C.get_Runs(-1, -1).Count == 1)
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
					this.B(A, B, C);
					try
					{
						enumerator = C.get_Paragraphs(-1, -1).GetEnumerator();
						while (enumerator.MoveNext())
						{
							TextRange2 c = (TextRange2)enumerator.Current;
							this.C(A, B, c);
							D(A, B, c);
							E(A, B, c);
						}
						while (true)
						{
							switch (5)
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
								switch (2)
								{
								case 0:
									break;
								default:
									(enumerator as IDisposable).Dispose();
									goto end_IL_008f;
								}
								continue;
								end_IL_008f:
								break;
							}
						}
					}
				}
			}
		}
		IEnumerator enumerator2 = C.get_Paragraphs(-1, -1).GetEnumerator();
		try
		{
			IEnumerator enumerator3 = default(IEnumerator);
			while (enumerator2.MoveNext())
			{
				TextRange2 textRange = (TextRange2)enumerator2.Current;
				if (textRange.get_Runs(-1, -1).Count == 1)
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
					this.B(A, B, textRange);
				}
				else
				{
					{
						enumerator3 = textRange.get_Runs(-1, -1).GetEnumerator();
						try
						{
							while (enumerator3.MoveNext())
							{
								TextRange2 c2 = (TextRange2)enumerator3.Current;
								this.B(A, B, c2);
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_012d;
								}
								continue;
								end_IL_012d:
								break;
							}
						}
						finally
						{
							IDisposable disposable2 = enumerator3 as IDisposable;
							if (disposable2 != null)
							{
								disposable2.Dispose();
							}
						}
					}
				}
				this.C(A, B, textRange);
				D(A, B, textRange);
				E(A, B, textRange);
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
			IDisposable disposable = enumerator2 as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	private void B(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, TextRange2 C)
	{
		if (Text.PrintableText(Strings.Trim(C.Text)).Length <= 0)
		{
			return;
		}
		UsedFontFamilies.Add(new Tuple<string, IndexedObject>(C.Font.Name, new IndexedObject(A, B, C)));
		UsedFontStyles.Add(new Tuple<FontStyle, IndexedObject>(this.A(C.Font), new IndexedObject(A, B, C)));
		List<Tuple<int, IndexedObject>> listColors = UsedFontColors;
		Index.TextRange(C, A, B, C, ref listColors);
		UsedFontColors = listColors;
		Font2 font = C.Font;
		if (font.Bold != MsoTriState.msoTrue)
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
			if (font.Italic != MsoTriState.msoTrue)
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
				if (font.UnderlineStyle == MsoTextUnderlineType.msoNoUnderline)
				{
					goto IL_0133;
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
		}
		TextDecoration item = new TextDecoration
		{
			Bold = (font.Bold == MsoTriState.msoTrue),
			Italic = (font.Italic == MsoTriState.msoTrue),
			UnderlineStyle = font.UnderlineStyle
		};
		UsedTextDecorations.Add(new Tuple<TextDecoration, IndexedObject>(item, new IndexedObject(A, B, C)));
		goto IL_0133;
		IL_0133:
		font = null;
	}

	private void C(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, TextRange2 C)
	{
		BulletFormat2 bullet = C.ParagraphFormat.Bullet;
		if (bullet.Type == MsoBulletType.msoBulletUnnumbered || bullet.Type == MsoBulletType.msoBulletNumbered)
		{
			if (bullet.UseTextColor == MsoTriState.msoFalse)
			{
				Microsoft.Office.Core.FillFormat fill = bullet.Font.Fill;
				UsedFontColors.Add(Index.NewTuple(fill.ForeColor.RGB, A, B, C.ParagraphFormat.Bullet));
				if (fill.BackColor.RGB != Base.TRANSPARENT && fill.BackColor.RGB != fill.ForeColor.RGB)
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
					UsedFontColors.Add(Index.NewTuple(fill.BackColor.RGB, A, B, C.ParagraphFormat.Bullet));
				}
				fill = null;
			}
			if (bullet.Type == MsoBulletType.msoBulletNumbered)
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
				BulletStyle item = new BulletStyle
				{
					Style = bullet.Style,
					IndentLevel = C.ParagraphFormat.IndentLevel
				};
				UsedBulletStyles.Add(new Tuple<BulletStyle, IndexedObject>(item, new IndexedObject(A, B, C.ParagraphFormat.Bullet)));
			}
		}
		bullet = null;
	}

	private void D(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, TextRange2 C)
	{
		LineSpacing item = default(LineSpacing);
		ParagraphFormat2 paragraphFormat = C.ParagraphFormat;
		item.Before = paragraphFormat.SpaceBefore;
		item.After = paragraphFormat.SpaceAfter;
		item.Within = paragraphFormat.SpaceWithin;
		item.LineRuleWithin = paragraphFormat.LineRuleWithin;
		UsedParagraphSpacing.Add(new Tuple<LineSpacing, IndexedObject>(item, new IndexedObject(A, B, C)));
		paragraphFormat = null;
	}

	private void E(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, TextRange2 C)
	{
		ParagraphFormat2 paragraphFormat = C.ParagraphFormat;
		if (paragraphFormat.LeftIndent != 0f)
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
			if (paragraphFormat.FirstLineIndent != 0f)
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
				Indent item = new Indent
				{
					LeftIndent = paragraphFormat.LeftIndent,
					FirstLineIndent = paragraphFormat.FirstLineIndent,
					IndentLevel = paragraphFormat.IndentLevel
				};
				UsedIndents.Add(new Tuple<Indent, IndexedObject>(item, new IndexedObject(A, B, C)));
			}
		}
		paragraphFormat = null;
	}

	private FontStyle A(Font2 A)
	{
		FontStyle result = default(FontStyle);
		Font2 font = A;
		result.Family = font.Name;
		result.Size = font.Size;
		font = null;
		return result;
	}

	private FontStyle A(Microsoft.Office.Core.ChartFont A)
	{
		FontStyle result = default(FontStyle);
		Microsoft.Office.Core.ChartFont chartFont = A;
		result.Family = Conversions.ToString(chartFont.Name);
		result.Size = Conversions.ToSingle(chartFont.Size);
		chartFont = null;
		return result;
	}

	private FontStyle A(Microsoft.Office.Interop.PowerPoint.ChartFont A)
	{
		FontStyle result = default(FontStyle);
		Microsoft.Office.Interop.PowerPoint.ChartFont chartFont = A;
		result.Family = Conversions.ToString(chartFont.Name);
		result.Size = Conversions.ToSingle(chartFont.Size);
		chartFont = null;
		return result;
	}

	public void DeterminePaletteUsage()
	{
		List<int> list = new List<int>();
		List<int> list2 = list;
		List<Tuple<int, IndexedObject>> usedFillColors = UsedFillColors;
		Func<Tuple<int, IndexedObject>, int> selector;
		if (_Closure_0024__.A == null)
		{
			selector = (_Closure_0024__.A = [SpecialName] (Tuple<int, IndexedObject> A) => A.Item1);
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
		list2.AddRange(UsedBorderColors.Select([SpecialName] (Tuple<int, IndexedObject> A) => A.Item1));
		List<Tuple<int, IndexedObject>> usedFontColors = UsedFontColors;
		Func<Tuple<int, IndexedObject>, int> selector2;
		if (_Closure_0024__.C == null)
		{
			selector2 = (_Closure_0024__.C = [SpecialName] (Tuple<int, IndexedObject> A) => A.Item1);
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
			selector2 = _Closure_0024__.C;
		}
		list2.AddRange(usedFontColors.Select(selector2));
		_ = null;
		((Conventions)this).PopulatePalette(list);
		list = null;
	}
}
