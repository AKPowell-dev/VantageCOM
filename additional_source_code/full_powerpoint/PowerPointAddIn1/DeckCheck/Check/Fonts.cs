using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.DeckCheck.Errors;

namespace PowerPointAddIn1.DeckCheck.Check;

public sealed class Fonts
{
	internal static void A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, TextRange2 C)
	{
		int? minFontSize = Main.Analysis.Conventions.MinFontSize;
		int? b = Main.Analysis.Conventions.MaxFontSize;
		if (b.HasValue && A.Shapes.HasTitle == MsoTriState.msoTrue && A.Shapes.Title == B)
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
			b = null;
		}
		if (!b.HasValue)
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
			if (!minFontSize.HasValue)
			{
				return;
			}
		}
		List<TextRange2> list = new List<TextRange2>();
		List<TextRange2> list2 = new List<TextRange2>();
		try
		{
			if (C.get_Runs(-1, -1).Count == 1)
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
				Font2 font = C.Font;
				if (Fonts.B(font.Size, b))
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
					list2.Add(C);
				}
				else if (Fonts.A(font.Size, minFontSize))
				{
					list.Add(C);
				}
				font = null;
			}
			else
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = C.get_Paragraphs(-1, -1).GetEnumerator();
					IEnumerator enumerator2 = default(IEnumerator);
					while (enumerator.MoveNext())
					{
						TextRange2 textRange = (TextRange2)enumerator.Current;
						if (textRange.get_Runs(-1, -1).Count == 1)
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
							Font2 font2 = textRange.Font;
							if (Fonts.B(font2.Size, b))
							{
								list2.Add(textRange);
							}
							else if (Fonts.A(font2.Size, minFontSize))
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
								list.Add(textRange);
							}
							font2 = null;
							continue;
						}
						{
							enumerator2 = textRange.get_Runs(-1, -1).GetEnumerator();
							try
							{
								while (enumerator2.MoveNext())
								{
									TextRange2 textRange2 = (TextRange2)enumerator2.Current;
									Font2 font3 = textRange2.Font;
									if (Fonts.B(font3.Size, b))
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
										list2.Add(textRange2);
									}
									else if (Fonts.A(font3.Size, minFontSize))
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
										list.Add(textRange2);
									}
									font3 = null;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_022e;
									}
									continue;
									end_IL_022e:
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
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_025e;
						}
						continue;
						end_IL_025e:
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
				if (Fonts.A(list2) == 1)
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
					list2.Clear();
					list2.Add(C);
				}
				if (Fonts.A(list) == 1)
				{
					list.Clear();
					list.Add(C);
				}
			}
			if (list2.Count > 0)
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
				Main.Analysis.Errors.Add(new MaxFontSize(A, B, list2, b.Value));
			}
			if (list.Count > 0)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					Main.Analysis.Errors.Add(new MinFontSize(A, B, list, minFontSize.Value));
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
		list = null;
		list2 = null;
	}

	internal static bool A(float A, int? B)
	{
		if (B.HasValue)
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
					return A < (float)B.Value;
				}
			}
		}
		return false;
	}

	internal static bool B(float A, int? B)
	{
		if (B.HasValue)
		{
			return A > (float)B.Value;
		}
		return false;
	}

	internal static void B(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, TextRange2 C)
	{
		List<TextRange2> list = new List<TextRange2>();
		try
		{
			if (C.get_Runs(-1, -1).Count == 1)
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
				if (Fonts.A(C.Font.Size))
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
					list.Add(C);
				}
			}
			else
			{
				IEnumerator enumerator2 = default(IEnumerator);
				foreach (TextRange2 item in C.get_Paragraphs(-1, -1))
				{
					if (item.get_Runs(-1, -1).Count == 1)
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
						if (!Fonts.A(item.Font.Size))
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
						list.Add(item);
						continue;
					}
					try
					{
						enumerator2 = item.get_Runs(-1, -1).GetEnumerator();
						while (enumerator2.MoveNext())
						{
							TextRange2 textRange2 = (TextRange2)enumerator2.Current;
							if (Fonts.A(textRange2.Font.Size))
							{
								list.Add(textRange2);
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_0114;
							}
							continue;
							end_IL_0114:
							break;
						}
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (7)
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
				if (Fonts.A(list) == 1)
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
					list.Clear();
					list.Add(C);
				}
			}
			if (list.Count > 0)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					Main.Analysis.Errors.Add(new FractionalFontSize(A, B, list));
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
		list = null;
	}

	private static bool A(float A)
	{
		return (float)checked((int)Math.Round(A)) != A;
	}

	internal static void C(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, TextRange2 C)
	{
		List<TextRange2> list = new List<TextRange2>();
		try
		{
			if (C.get_Runs(-1, -1).Count == 1)
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
				if (Fonts.A(C))
				{
					list.Add(C);
				}
			}
			else
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = C.get_Paragraphs(-1, -1).GetEnumerator();
					IEnumerator enumerator2 = default(IEnumerator);
					while (enumerator.MoveNext())
					{
						TextRange2 textRange = (TextRange2)enumerator.Current;
						if (textRange.get_Runs(-1, -1).Count == 1)
						{
							if (!Fonts.A(textRange))
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
							list.Add(textRange);
							continue;
						}
						try
						{
							enumerator2 = textRange.get_Runs(-1, -1).GetEnumerator();
							while (enumerator2.MoveNext())
							{
								TextRange2 textRange2 = (TextRange2)enumerator2.Current;
								if (!Fonts.A(textRange2))
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
								list.Add(textRange2);
							}
							while (true)
							{
								switch (5)
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
			if (list.Count > 0)
			{
				Main.Analysis.Errors.Add(new Strikethrough(A, B, list));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		list = null;
	}

	private static bool A(TextRange2 A)
	{
		if (A.Font.StrikeThrough != MsoTriState.msoTrue)
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
					return A.Font.DoubleStrikeThrough == MsoTriState.msoTrue;
				}
			}
		}
		return true;
	}

	internal static void A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B, TextRange2 C, List<string> D)
	{
		if (D.Count == 0)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			List<TextRange2> list = new List<TextRange2>();
			try
			{
				if (C.get_Runs(-1, -1).Count == 1)
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
					if (Fonts.A(C, D))
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
						list.Add(C);
					}
				}
				else
				{
					try
					{
						enumerator = C.get_Paragraphs(-1, -1).GetEnumerator();
						while (enumerator.MoveNext())
						{
							TextRange2 textRange = (TextRange2)enumerator.Current;
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
								if (!Fonts.A(textRange, D))
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
								list.Add(textRange);
								continue;
							}
							{
								enumerator2 = textRange.get_Runs(-1, -1).GetEnumerator();
								try
								{
									while (enumerator2.MoveNext())
									{
										TextRange2 textRange2 = (TextRange2)enumerator2.Current;
										if (Fonts.A(textRange2, D))
										{
											list.Add(textRange2);
										}
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_0103;
										}
										continue;
										end_IL_0103:
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
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0130;
							}
							continue;
							end_IL_0130:
							break;
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
				if (list.Count > 0)
				{
					Main.Analysis.Errors.Add(new IllegalFont(A, B, list, D));
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			list = null;
			return;
		}
	}

	private static bool A(TextRange2 A, List<string> B)
	{
		return !B.Contains(A.Font.Name);
	}

	internal static void A(Slide A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		if (B.TextFrame2.AutoSize != MsoAutoSize.msoAutoSizeTextToFitShape)
		{
			return;
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
			Main.Analysis.Errors.Add(new ShrinkTextOnOverflow(A, B));
			return;
		}
	}

	private static int A(List<TextRange2> A)
	{
		if (A.Count > 0)
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
					return A.Select([SpecialName] (TextRange2 textRange) => textRange.Font.Size).Distinct().Count();
				}
			}
		}
		return 0;
	}
}
