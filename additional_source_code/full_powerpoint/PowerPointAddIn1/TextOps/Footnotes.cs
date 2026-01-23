using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.MasterShapes;

namespace PowerPointAddIn1.TextOps;

public sealed class Footnotes
{
	public enum FootnoteType
	{
		Numbers,
		Lowercase,
		Uppercase
	}

	public enum FootnoteOrder
	{
		TopThenLeft,
		LeftThenTop
	}

	public enum FootnoteDecoration
	{
		Parentheses,
		Brackets,
		None
	}

	private struct ZF
	{
		public bool A;

		public FootnoteOrder A;

		public FootnoteType A;

		public FootnoteDecoration A;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, float> A;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, float> B;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, float> C;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, float> D;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal float A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A.Left;
		}

		[SpecialName]
		internal float B(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A.Top;
		}

		[SpecialName]
		internal float C(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A.Top;
		}

		[SpecialName]
		internal float D(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A.Left;
		}
	}

	[CompilerGenerated]
	internal sealed class AG
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public AG(AG A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.Copy();
		}
	}

	private static readonly List<string> m_A = new List<string>(new string[26]
	{
		AH.A(9078),
		AH.A(9081),
		AH.A(9084),
		AH.A(9087),
		AH.A(9090),
		AH.A(9093),
		AH.A(9096),
		AH.A(9099),
		AH.A(9102),
		AH.A(9105),
		AH.A(9110),
		AH.A(9115),
		AH.A(9120),
		AH.A(9125),
		AH.A(9130),
		AH.A(9135),
		AH.A(9140),
		AH.A(9145),
		AH.A(9150),
		AH.A(9155),
		AH.A(9160),
		AH.A(9165),
		AH.A(9170),
		AH.A(9175),
		AH.A(9180),
		AH.A(9185)
	});

	private static readonly List<string> m_B = new List<string>(new string[26]
	{
		AH.A(8100),
		AH.A(8103),
		AH.A(8106),
		AH.A(8109),
		AH.A(8112),
		AH.A(8115),
		AH.A(8118),
		AH.A(8121),
		AH.A(8124),
		AH.A(8127),
		AH.A(8130),
		AH.A(8133),
		AH.A(8136),
		AH.A(8139),
		AH.A(8142),
		AH.A(8145),
		AH.A(8148),
		AH.A(8151),
		AH.A(8154),
		AH.A(8157),
		AH.A(8160),
		AH.A(8163),
		AH.A(8166),
		AH.A(8169),
		AH.A(8172),
		AH.A(8175)
	});

	private static readonly List<string> C = new List<string>(new string[26]
	{
		AH.A(7902),
		AH.A(7905),
		AH.A(7908),
		AH.A(7911),
		AH.A(7914),
		AH.A(7917),
		AH.A(7920),
		AH.A(7923),
		AH.A(7926),
		AH.A(7929),
		AH.A(7932),
		AH.A(7935),
		AH.A(7938),
		AH.A(7941),
		AH.A(7944),
		AH.A(7947),
		AH.A(7950),
		AH.A(7953),
		AH.A(7956),
		AH.A(7959),
		AH.A(7962),
		AH.A(7965),
		AH.A(7968),
		AH.A(7971),
		AH.A(7974),
		AH.A(7977)
	});

	private static readonly string m_A = AH.A(155506);

	private static readonly string m_B = AH.A(155537);

	public static void Add()
	{
		if (!Licensing.AllowAdvancedTextOperation())
		{
			return;
		}
		checked
		{
			ZF B = default(ZF);
			int C = default(int);
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
				Selection A = null;
				Microsoft.Office.Interop.PowerPoint.Shape b;
				TextRange2 textRange;
				TextRange2 textRange2;
				if (Footnotes.A(ref A, ref B))
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
					try
					{
						b = A.ShapeRange[1];
						textRange = A.TextRange2;
						string d;
						if (!Footnotes.A(textRange, b))
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
							FootnoteDecoration a = B.A;
							string newText;
							if (a != FootnoteDecoration.Parentheses)
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
								if (a != FootnoteDecoration.Brackets)
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
									newText = AH.A(139285);
								}
								else
								{
									newText = AH.A(154368);
								}
							}
							else
							{
								newText = AH.A(154361);
							}
							textRange.InsertAfter(newText).Select();
							textRange.Font.Superscript = MsoTriState.msoTrue;
							textRange2 = textRange;
							d = "";
						}
						else
						{
							textRange2 = Footnotes.A(textRange, b);
							string text = textRange2.Text;
							d = text;
							if (Regex.IsMatch(text, AH.A(154375) + Footnotes.m_A + AH.A(154382)))
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
								if (B.A)
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
									int num = text.IndexOf(AH.A(14255));
									text = Strings.Left(text, num) + AH.A(154389) + Strings.Mid(text, num + 1);
								}
								else
								{
									text += AH.A(154361);
								}
							}
							else if (Regex.IsMatch(text, AH.A(71202) + Footnotes.m_A + AH.A(71209)))
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
								if (B.A)
								{
									int num = text.IndexOf(AH.A(15138));
									text = Strings.Left(text, num) + AH.A(154389) + Strings.Mid(text, num + 1);
								}
								else
								{
									text += AH.A(154368);
								}
							}
							else if (Regex.IsMatch(text, AH.A(154394) + Footnotes.m_A + AH.A(154397)))
							{
								text += AH.A(154389);
							}
							else if (Regex.IsMatch(text, AH.A(154400) + Footnotes.m_A + AH.A(154409)))
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
								text += AH.A(154361);
							}
							else if (Regex.IsMatch(text, AH.A(154426) + Footnotes.m_A + AH.A(154435)))
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
								text += AH.A(154368);
							}
							else if (Regex.IsMatch(text, AH.A(154394) + Footnotes.m_A + AH.A(154452) + Footnotes.m_A + AH.A(154457)))
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
								text += AH.A(154389);
							}
							else if (Regex.IsMatch(text, AH.A(154394) + Footnotes.m_A + AH.A(154464) + Footnotes.m_A + AH.A(154457)))
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
								text += AH.A(154473);
							}
							else if (Regex.IsMatch(text, AH.A(154480) + Footnotes.m_A + AH.A(154452) + Footnotes.m_A + AH.A(154495)))
							{
								int num = text.IndexOf(AH.A(12717));
								text = Strings.Left(text, num) + AH.A(154389) + Strings.Mid(text, num + 1);
							}
							else if (Regex.IsMatch(text, AH.A(154480) + Footnotes.m_A + AH.A(154464) + Footnotes.m_A + AH.A(154495)))
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
								int num = text.IndexOf(AH.A(12717));
								text = Strings.Left(text, num) + AH.A(154473) + Strings.Mid(text, num + 1);
							}
							else
							{
								Forms.WarningMessage(AH.A(154514));
							}
							textRange2.Text = text;
						}
						Footnotes.A(A, B, ref C);
						Footnotes.A(A.SlideRange[1], B, C: true, d, textRange2.Text);
						textRange2.Select();
						if (C > Footnotes.m_A.Count)
						{
							Forms.ErrorMessage(AH.A(154712) + Footnotes.m_A.Count + AH.A(154753));
						}
						Base.LogActivity(AH.A(154796));
					}
					catch (BG bG)
					{
						ProjectData.SetProjectError(bG);
						BG a2 = bG;
						Footnotes.A(a2);
						ProjectData.ClearProjectError();
					}
				}
				else
				{
					Forms.WarningMessage(AH.A(154821));
				}
				b = null;
				textRange = null;
				textRange2 = null;
				A = null;
				return;
			}
		}
	}

	public static void Remove()
	{
		if (!Licensing.AllowAdvancedTextOperation())
		{
			return;
		}
		Selection A = null;
		ZF B = default(ZF);
		Microsoft.Office.Interop.PowerPoint.Shape b;
		TextRange2 textRange;
		TextRange2 textRange2;
		if (Footnotes.A(ref A, ref B))
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
			try
			{
				b = A.ShapeRange[1];
				textRange = A.TextRange2;
				if (!Footnotes.A(textRange, b))
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						Forms.WarningMessage(AH.A(154928));
						break;
					}
				}
				else
				{
					textRange2 = Footnotes.A(textRange, b);
					string text = textRange2.Text;
					string d = text;
					if (Regex.IsMatch(text, AH.A(154375) + Footnotes.m_A + AH.A(154382)))
					{
						text = "";
					}
					else if (Regex.IsMatch(text, AH.A(71202) + Footnotes.m_A + AH.A(71209)))
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
						text = "";
					}
					else if (Regex.IsMatch(text, AH.A(154394) + Footnotes.m_A + AH.A(154397)))
					{
						text = "";
					}
					else if (Regex.IsMatch(text, AH.A(154400) + Footnotes.m_A + AH.A(154409)))
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
						text = Regex.Replace(text, AH.A(17514) + Footnotes.m_A + AH.A(154382), "");
					}
					else if (Regex.IsMatch(text, AH.A(154426) + Footnotes.m_A + AH.A(154435)))
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
						text = Regex.Replace(text, AH.A(17462) + Footnotes.m_A + AH.A(71209), "");
					}
					else if (Regex.IsMatch(text, AH.A(154394) + Footnotes.m_A + AH.A(154452) + Footnotes.m_A + AH.A(154457)))
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
						text = Regex.Replace(text, AH.A(12717) + Footnotes.m_A + AH.A(154397), "");
					}
					else if (Regex.IsMatch(text, AH.A(154394) + Footnotes.m_A + AH.A(154464) + Footnotes.m_A + AH.A(154457)))
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
						text = Regex.Replace(text, AH.A(154991) + Footnotes.m_A + AH.A(154397), "");
					}
					else if (Regex.IsMatch(text, AH.A(154480) + Footnotes.m_A + AH.A(154452) + Footnotes.m_A + AH.A(154495)))
					{
						text = Regex.Replace(text, AH.A(12717) + Footnotes.m_A + AH.A(154998), AH.A(44617));
					}
					else if (Regex.IsMatch(text, AH.A(154480) + Footnotes.m_A + AH.A(154464) + Footnotes.m_A + AH.A(154495)))
					{
						text = Regex.Replace(text, AH.A(154991) + Footnotes.m_A + AH.A(154998), AH.A(44617));
					}
					else
					{
						Forms.WarningMessage(AH.A(155017));
					}
					textRange2.Text = text;
					int C = default(int);
					Footnotes.A(A, B, ref C);
					Footnotes.A(A.SlideRange[1], B, C: false, d, textRange2.Text);
					textRange2.Select();
					Base.LogActivity(AH.A(155221));
				}
			}
			catch (BG bG)
			{
				ProjectData.SetProjectError(bG);
				BG a = bG;
				Footnotes.A(a);
				ProjectData.ClearProjectError();
			}
		}
		else
		{
			Forms.WarningMessage(AH.A(155252));
		}
		b = null;
		textRange = null;
		textRange2 = null;
		A = null;
	}

	private static void A(BG A)
	{
		NG.A.Application.CommandBars.ExecuteMso(AH.A(40491));
		System.Windows.Forms.Application.DoEvents();
		Forms.WarningMessage(A.Message);
	}

	private static bool A(ref Selection A, ref ZF B)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		bool result = false;
		try
		{
			A = application.ActiveWindow.Selection;
			try
			{
				if (A.Type == PpSelectionType.ppSelectionText)
				{
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
						result = true;
						B = Footnotes.A();
						application.StartNewUndoEntry();
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.WarningMessage(AH.A(13552));
			ProjectData.ClearProjectError();
		}
		application = null;
		return result;
	}

	private static ZF A()
	{
		ZF result = default(ZF);
		try
		{
			XmlNode xmlNode = KG.A.SettingsXml.SelectSingleNode(AH.A(155331));
			result.A = (FootnoteDecoration)Conversions.ToInteger(xmlNode.SelectSingleNode(AH.A(155364)).InnerText);
			result.A = Conversions.ToBoolean(xmlNode.SelectSingleNode(AH.A(96093)).InnerText);
			result.A = (FootnoteOrder)Conversions.ToInteger(xmlNode.SelectSingleNode(AH.A(155385)).InnerText);
			result.A = (FootnoteType)Conversions.ToInteger(xmlNode.SelectSingleNode(AH.A(120502)).InnerText);
			xmlNode = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result.A = FootnoteDecoration.Parentheses;
			result.A = true;
			result.A = FootnoteOrder.TopThenLeft;
			result.A = FootnoteType.Numbers;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static bool A(TextRange2 A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		bool result = false;
		if (A.Font.Superscript == MsoTriState.msoFalse)
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
			try
			{
				if (B.TextFrame2.TextRange.get_Characters(A.Start, 1).Font.Superscript == MsoTriState.msoTrue)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						result = true;
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
		}
		else
		{
			result = true;
		}
		return result;
	}

	private static TextRange2 A(TextRange2 A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		TextRange2 textRange = B.TextFrame2.TextRange;
		int num = A.Start;
		checked
		{
			try
			{
				int num2 = A.Start - 1;
				while (true)
				{
					if (num2 >= 1)
					{
						if (textRange.get_Characters(num2, 1).Font.Superscript != MsoTriState.msoTrue)
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							num = num2;
							num2 += -1;
							break;
						}
						continue;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_005c;
						}
						continue;
						end_IL_005c:
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
			int num3 = A.Start - 1;
			try
			{
				int start = A.Start;
				int count = textRange.get_Characters(-1, -1).Count;
				int num4 = start;
				while (true)
				{
					IL_016b:
					if (num4 <= count)
					{
						string text = textRange.get_Characters(num4, 1).Text;
						if (Operators.CompareString(text, AH.A(7894), TextCompare: false) == 0)
						{
							break;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							if (Operators.CompareString(text, AH.A(47331), TextCompare: false) == 0)
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
								if (Operators.CompareString(text, AH.A(47334), TextCompare: false) == 0)
								{
									break;
								}
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									if (Operators.CompareString(text, AH.A(155396), TextCompare: false) == 0)
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
										if (textRange.get_Characters(num4, 1).Font.Superscript != MsoTriState.msoTrue)
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
												goto end_IL_0158;
											}
											continue;
											end_IL_0158:
											break;
										}
										num3 = num4;
										num4++;
										goto IL_016b;
									}
									break;
								}
								break;
							}
							break;
						}
						break;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_0174;
						}
						continue;
						end_IL_0174:
						break;
					}
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			TextRange2 result = textRange.get_Characters(num, num3 - num + 1);
			textRange = null;
			return result;
		}
	}

	private static void A(Selection A, ZF B, ref int C)
	{
		List<Microsoft.Office.Interop.PowerPoint.Shape> A2 = A.SlideRange[1].Shapes.Cast<Microsoft.Office.Interop.PowerPoint.Shape>().ToList();
		if (B.A == FootnoteOrder.TopThenLeft)
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
			Footnotes.B(ref A2);
		}
		else
		{
			Footnotes.A(ref A2);
		}
		C = 0;
		checked
		{
			IEnumerator enumerator2 = default(IEnumerator);
			IEnumerator enumerator3 = default(IEnumerator);
			foreach (Microsoft.Office.Interop.PowerPoint.Shape item in A2)
			{
				if (Footnotes.A(item))
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
					Footnotes.A(item.TextFrame2.TextRange, B.A, ref C);
				}
				else if (item.Type == MsoShapeType.msoGroup)
				{
					try
					{
						enumerator2 = item.GroupItems.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
							if (!Footnotes.A(shape))
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
							Footnotes.A(shape.TextFrame2.TextRange, B.A, ref C);
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_0110;
							}
							continue;
							end_IL_0110:
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
				}
				else if (item.HasTable == MsoTriState.msoTrue)
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
					Table table = item.Table;
					int count = table.Rows.Count;
					int count2 = table.Columns.Count;
					if (B.A == FootnoteOrder.TopThenLeft)
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
						int num = count;
						for (int i = 1; i <= num; i++)
						{
							float height = table.Rows[i].Height;
							int num2 = count2;
							for (int j = 1; j <= num2; j++)
							{
								if (Tables.IsCellMergedOrSplit(table.Cell(i, j), height, table.Columns[j].Width))
								{
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											throw new BG();
										}
									}
								}
								if (!Footnotes.A(table.Cell(i, j).Shape))
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
								Footnotes.A(table.Cell(i, j).Shape.TextFrame2.TextRange, B.A, ref C);
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_0264;
								}
								continue;
								end_IL_0264:
								break;
							}
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
					else
					{
						int num3 = count2;
						for (int k = 1; k <= num3; k++)
						{
							float width = table.Columns[k].Width;
							int num4 = count;
							for (int l = 1; l <= num4; l++)
							{
								if (Tables.IsCellMergedOrSplit(table.Cell(l, k), table.Rows[l].Height, width))
								{
									while (true)
									{
										switch (4)
										{
										case 0:
											break;
										default:
											throw new BG();
										}
									}
								}
								if (!Footnotes.A(table.Cell(l, k).Shape))
								{
									continue;
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
								Footnotes.A(table.Cell(l, k).Shape.TextFrame2.TextRange, B.A, ref C);
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_0360;
								}
								continue;
								end_IL_0360:
								break;
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
					table = null;
				}
				else
				{
					if (item.HasSmartArt != MsoTriState.msoTrue)
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
					{
						enumerator3 = item.SmartArt.AllNodes.GetEnumerator();
						try
						{
							while (enumerator3.MoveNext())
							{
								SmartArtNode smartArtNode = (SmartArtNode)enumerator3.Current;
								if (smartArtNode.TextFrame2.HasText != MsoTriState.msoTrue)
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
								Footnotes.A(smartArtNode.TextFrame2.TextRange, B.A, ref C);
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_040d;
								}
								continue;
								end_IL_040d:
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
					}
				}
			}
			JG.A(A2);
		}
	}

	private static void A(TextRange2 A, FootnoteType B, ref int C)
	{
		Regex regex = Footnotes.A();
		List<TextRange2> list = null;
		List<string> list2;
		if (B != FootnoteType.Numbers)
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
			list2 = ((B != FootnoteType.Lowercase) ? Footnotes.C : Footnotes.m_B);
		}
		else
		{
			list2 = Footnotes.m_A;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.get_Runs(-1, -1).GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				IEnumerator enumerator3 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					TextRange2 textRange = (TextRange2)enumerator.Current;
					if (textRange.Font.Superscript != MsoTriState.msoTrue)
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
					MatchCollection matchCollection = regex.Matches(textRange.Text);
					if (matchCollection.Count > 0)
					{
						list = new List<TextRange2>();
						try
						{
							enumerator2 = matchCollection.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Match match = (Match)enumerator2.Current;
								int num = match.Groups.Count - 1;
								for (int i = 1; i <= num; i++)
								{
									if (match.Groups[i].Value.Length <= 0)
									{
										continue;
									}
									try
									{
										enumerator3 = match.Groups[i].Captures.GetEnumerator();
										while (enumerator3.MoveNext())
										{
											Capture capture = (Capture)enumerator3.Current;
											string value = capture.Value;
											if (Operators.CompareString(value, AH.A(155399), TextCompare: false) == 0)
											{
												continue;
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
											if (Operators.CompareString(value, AH.A(155404), TextCompare: false) == 0)
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
											if (Operators.CompareString(value, AH.A(155409), TextCompare: false) == 0 || Operators.CompareString(value, AH.A(155414), TextCompare: false) == 0)
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
											if (Operators.CompareString(value, AH.A(155419), TextCompare: false) == 0)
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
											list.Add(textRange.get_Characters(capture.Index + 1, capture.Length));
										}
										while (true)
										{
											switch (1)
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
										if (enumerator3 is IDisposable)
										{
											while (true)
											{
												switch (4)
												{
												case 0:
													continue;
												}
												(enumerator3 as IDisposable).Dispose();
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
										continue;
									}
									break;
								}
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_026f;
								}
								continue;
								end_IL_026f:
								break;
							}
						}
						finally
						{
							if (enumerator2 is IDisposable)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										continue;
									}
									(enumerator2 as IDisposable).Dispose();
									break;
								}
							}
						}
						using List<TextRange2>.Enumerator enumerator4 = list.GetEnumerator();
						while (enumerator4.MoveNext())
						{
							TextRange2 current = enumerator4.Current;
							if (C < list2.Count)
							{
								current.Text = list2[C];
							}
							C++;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_02e1;
							}
							continue;
							end_IL_02e1:
							break;
						}
					}
					matchCollection = null;
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
			regex = null;
			list2 = null;
			JG.A(list);
		}
	}

	private static Regex A()
	{
		return new Regex(AH.A(155422) + Footnotes.m_A + AH.A(155447) + Footnotes.m_A + AH.A(155466) + Footnotes.m_A + AH.A(155501), RegexOptions.IgnoreCase);
	}

	private static void A(ref List<Microsoft.Office.Interop.PowerPoint.Shape> A)
	{
		IOrderedEnumerable<Microsoft.Office.Interop.PowerPoint.Shape> source = A.OrderBy([SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => shape.Left);
		Func<Microsoft.Office.Interop.PowerPoint.Shape, float> keySelector;
		if (_Closure_0024__.B == null)
		{
			keySelector = (_Closure_0024__.B = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => shape.Top);
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
			keySelector = _Closure_0024__.B;
		}
		A = source.ThenBy(keySelector).ToList();
	}

	private static void B(ref List<Microsoft.Office.Interop.PowerPoint.Shape> A)
	{
		List<Microsoft.Office.Interop.PowerPoint.Shape> source = A;
		Func<Microsoft.Office.Interop.PowerPoint.Shape, float> keySelector;
		if (_Closure_0024__.C == null)
		{
			keySelector = (_Closure_0024__.C = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => shape.Top);
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
			keySelector = _Closure_0024__.C;
		}
		A = source.OrderBy(keySelector).ThenBy([SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => shape.Left).ToList();
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (A.HasTextFrame == MsoTriState.msoTrue)
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
					return A.TextFrame2.HasText == MsoTriState.msoTrue;
				}
			}
		}
		return false;
	}

	private static void A(Slide A, ZF B, bool C, string D, string E)
	{
		AG a = default(AG);
		AG CS_0024_003C_003E8__locals9 = new AG(a);
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		CS_0024_003C_003E8__locals9.A = null;
		_ = (float)(0.75 * (double)A.CustomLayout.Height);
		try
		{
			CS_0024_003C_003E8__locals9.A = ((Microsoft.Office.Interop.PowerPoint.Presentation)A.Parent).Designs[1].SlideMaster.Shapes[Footnotes.m_B];
			if (CS_0024_003C_003E8__locals9.A.HasTextFrame == MsoTriState.msoFalse)
			{
				goto IL_00a5;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (CS_0024_003C_003E8__locals9.A.Visible == MsoTriState.msoTrue)
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
				goto IL_00a5;
			}
			goto end_IL_002c;
			IL_00a5:
			CS_0024_003C_003E8__locals9.A = null;
			end_IL_002c:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (CS_0024_003C_003E8__locals9.A == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			try
			{
				AddRemove.Master master = AddRemove.MasterShapeProperties(CS_0024_003C_003E8__locals9.A);
				try
				{
					enumerator = A.Shapes.GetEnumerator();
					while (true)
					{
						if (enumerator.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
							if (!PowerPointAddIn1.MasterShapes.Base.A(shape2, master.Id))
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
								shape = shape2;
								break;
							}
							break;
						}
						while (true)
						{
							switch (2)
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
						break;
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
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				if (C && shape == null)
				{
					clsClipboard.CopyWithWait((Action)([SpecialName] () =>
					{
						CS_0024_003C_003E8__locals9.A.Copy();
					}), 4000);
					shape = AddRemove.TryPaste(CS_0024_003C_003E8__locals9.A, A.Shapes);
					AddRemove.PrepShape(shape, master);
					shape.Visible = MsoTriState.msoTrue;
				}
				return;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}
}
