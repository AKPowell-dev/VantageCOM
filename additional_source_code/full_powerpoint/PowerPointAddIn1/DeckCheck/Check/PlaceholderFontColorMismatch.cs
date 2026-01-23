using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class PlaceholderFontColorMismatch
{
	public void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, Microsoft.Office.Interop.PowerPoint.Shape placeholder)
	{
		if (shp.HasTextFrame != MsoTriState.msoTrue)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
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
			TextRange2 textRange = shp.TextFrame2.TextRange;
			int rGB = textRange.Font.Fill.ForeColor.RGB;
			int rGB2 = placeholder.TextFrame2.TextRange.Font.Fill.ForeColor.RGB;
			List<TextRange2> C;
			if (rGB2 < 0)
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
				Dictionary<int, int> dictionary = new Dictionary<int, int>();
				try
				{
					enumerator = placeholder.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
					while (enumerator.MoveNext())
					{
						TextRange2 textRange2 = (TextRange2)enumerator.Current;
						ParagraphFormat2 paragraphFormat = textRange2.ParagraphFormat;
						if (!dictionary.ContainsKey(paragraphFormat.IndentLevel))
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
							dictionary.Add(paragraphFormat.IndentLevel, textRange2.Font.Fill.ForeColor.RGB);
						}
						paragraphFormat = null;
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
				using (Dictionary<int, int>.Enumerator enumerator2 = dictionary.GetEnumerator())
				{
					while (enumerator2.MoveNext())
					{
						KeyValuePair<int, int> current = enumerator2.Current;
						C = new List<TextRange2>();
						try
						{
							enumerator3 = textRange.get_Paragraphs(-1, -1).GetEnumerator();
							while (enumerator3.MoveNext())
							{
								TextRange2 textRange3 = (TextRange2)enumerator3.Current;
								if (textRange3.ParagraphFormat.IndentLevel == current.Key)
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
									A(textRange3, current.Value, ref C);
								}
								_ = null;
							}
						}
						finally
						{
							if (enumerator3 is IDisposable)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										continue;
									}
									(enumerator3 as IDisposable).Dispose();
									break;
								}
							}
						}
						if (C.Count <= 0)
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
						Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.PlaceholderFontColorMismatch(sld, shp, C, current.Value));
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_0226;
						}
						continue;
						end_IL_0226:
						break;
					}
				}
				dictionary = null;
			}
			else if (rGB < 0)
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
				C = new List<TextRange2>();
				try
				{
					enumerator4 = textRange.get_Paragraphs(-1, -1).GetEnumerator();
					while (enumerator4.MoveNext())
					{
						TextRange2 a = (TextRange2)enumerator4.Current;
						A(a, rGB2, ref C);
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0298;
						}
						continue;
						end_IL_0298:
						break;
					}
				}
				finally
				{
					if (enumerator4 is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator4 as IDisposable).Dispose();
							break;
						}
					}
				}
				if (C.Count > 0)
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
					Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.PlaceholderFontColorMismatch(sld, shp, C, rGB2));
				}
			}
			else if (rGB != rGB2)
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
				Main.Analysis.Errors.Add(new PowerPointAddIn1.DeckCheck.Errors.PlaceholderFontColorMismatch(sld, shp, new List<TextRange2>(new TextRange2[1] { textRange }), rGB2));
			}
			C = null;
			textRange = null;
			return;
		}
	}

	private void A(TextRange2 A, int B, ref List<TextRange2> C)
	{
		TextRange2 textRange = A;
		if (textRange.Font.Fill.ForeColor.RGB < 0)
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
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = textRange.get_Runs(-1, -1).GetEnumerator();
				while (enumerator.MoveNext())
				{
					TextRange2 textRange2 = (TextRange2)enumerator.Current;
					if (textRange2.Font.Fill.ForeColor.RGB == B)
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
					C.Add(textRange2);
				}
				while (true)
				{
					switch (5)
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
		}
		else if (textRange.Font.Fill.ForeColor.RGB != B)
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
			C.Add(A);
		}
		textRange = null;
	}
}
