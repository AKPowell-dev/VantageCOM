using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.TextOps;

public sealed class Fonts
{
	public static void Replace()
	{
		if (!Licensing.AllowAdvancedTextOperation())
		{
			return;
		}
		Application application = NG.A.Application;
		List<string> list = new List<string>();
		List<string> list2 = new List<string>();
		List<string> list3 = new List<string>();
		ReplaceFontsOptions replaceFontsOptions = null;
		List<string> legalFontTypes;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
		try
		{
			activePresentation = application.ActivePresentation;
			legalFontTypes = new Settings(activePresentation).LegalFontTypes;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			application = null;
			ProjectData.ClearProjectError();
			return;
		}
		FontFamily[] families = new InstalledFontCollection().Families;
		foreach (FontFamily fontFamily in families)
		{
			list3.Add(fontFamily.Name);
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
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
			list2.Add(AH.A(154270));
			list2.AddRange(list3);
			if (legalFontTypes.Count == 0)
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
				list.AddRange(list3);
			}
			else
			{
				list.AddRange(legalFontTypes);
			}
			wpfReplaceFonts wpfReplaceFonts2 = new wpfReplaceFonts();
			wpfReplaceFonts2.cbxFind.ItemsSource = list2;
			wpfReplaceFonts2.cbxFind.SelectedIndex = 0;
			wpfReplaceFonts2.cbxReplace.ItemsSource = list;
			wpfReplaceFonts2.ShowDialog();
			if (wpfReplaceFonts2.DialogResult.HasValue)
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
				if (wpfReplaceFonts2.DialogResult.Value)
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
					object obj;
					if (wpfReplaceFonts2.cbxFind.SelectedIndex != 0)
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
						obj = wpfReplaceFonts2.cbxFind.SelectedValue.ToString();
					}
					else
					{
						obj = "";
					}
					string strFind = (string)obj;
					replaceFontsOptions = new ReplaceFontsOptions(strFind, Conversions.ToString(wpfReplaceFonts2.cbxReplace.SelectedValue), wpfReplaceFonts2.chkBold.IsChecked.Value, wpfReplaceFonts2.chkItalic.IsChecked.Value, wpfReplaceFonts2.chkUnderline.IsChecked.Value);
				}
			}
			wpfReplaceFonts2 = null;
			list = null;
			list2 = null;
			list3 = null;
			legalFontTypes = null;
			if (replaceFontsOptions != null)
			{
				try
				{
					application.StartNewUndoEntry();
					try
					{
						enumerator = activePresentation.Designs.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Design design = (Design)enumerator.Current;
							A(design.SlideMaster.Shapes, replaceFontsOptions);
							try
							{
								enumerator2 = design.SlideMaster.CustomLayouts.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									A(((CustomLayout)enumerator2.Current).Shapes, replaceFontsOptions);
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_0293;
									}
									continue;
									end_IL_0293:
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
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_02cb;
							}
							continue;
							end_IL_02cb:
							break;
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
					enumerator3 = activePresentation.Slides.GetEnumerator();
					try
					{
						while (enumerator3.MoveNext())
						{
							Slide slide = (Slide)enumerator3.Current;
							try
							{
								application.ActiveWindow.View.GotoSlide(slide.SlideIndex);
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							A(slide.Shapes, replaceFontsOptions);
							A(slide.NotesPage.Shapes, replaceFontsOptions);
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0373;
							}
							continue;
							end_IL_0373:
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
					Forms.SuccessMessage(AH.A(154277));
					Base.LogActivity(AH.A(154334));
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					Forms.ErrorMessage(ex6.Message);
					clsReporting.LogException(ex6);
					ProjectData.ClearProjectError();
				}
			}
			application = null;
			activePresentation = null;
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shapes A, ReplaceFontsOptions B)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Fonts.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, B);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, ReplaceFontsOptions B)
	{
		IEnumerator enumerator = default(IEnumerator);
		if (A.Type == MsoShapeType.msoGroup)
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
					{
						enumerator = A.GroupItems.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
								if (shape.HasTextFrame == MsoTriState.msoTrue)
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
									Fonts.A(shape, B);
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
							IDisposable disposable = enumerator as IDisposable;
							if (disposable != null)
							{
								disposable.Dispose();
							}
						}
					}
				}
			}
		}
		if (A.HasTextFrame == MsoTriState.msoTrue)
		{
			Fonts.B(A, B);
			return;
		}
		checked
		{
			if (A.HasTable == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
					{
						Table table = A.Table;
						int count = table.Rows.Count;
						for (int i = 1; i <= count; i++)
						{
							int count2 = table.Columns.Count;
							for (int j = 1; j <= count2; j++)
							{
								if (table.Cell(i, j).Shape.HasTextFrame == MsoTriState.msoTrue)
								{
									Fonts.B(table.Cell(i, j).Shape, B);
								}
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_0121;
								}
								continue;
								end_IL_0121:
								break;
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								table = null;
								return;
							}
						}
					}
					}
				}
			}
			IEnumerator enumerator2 = default(IEnumerator);
			if (A.HasChart == MsoTriState.msoTrue)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						if (B.Find.Length == 0)
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
							Fonts.A(A.Chart.ChartArea.Format.TextFrame2.TextRange, B);
						}
						else
						{
							Fonts.A(A.Chart.ChartArea.Format.TextFrame2.TextRange, B);
						}
						try
						{
							try
							{
								enumerator2 = A.Chart.Shapes.GetEnumerator();
								while (enumerator2.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
									if (shape2.HasTextFrame == MsoTriState.msoTrue)
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
										Fonts.B(shape2, B);
									}
								}
								return;
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
											break;
										default:
											(enumerator2 as IDisposable).Dispose();
											goto end_IL_021d;
										}
										continue;
										end_IL_021d:
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
							return;
						}
					}
				}
			}
			if (A.HasSmartArt != MsoTriState.msoTrue)
			{
				return;
			}
			IEnumerator enumerator3 = default(IEnumerator);
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				enumerator3 = A.SmartArt.AllNodes.GetEnumerator();
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
							switch (7)
							{
							case 0:
								continue;
							}
							break;
						}
						Fonts.A(smartArtNode.TextFrame2.TextRange, B);
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
					IDisposable disposable2 = enumerator3 as IDisposable;
					if (disposable2 != null)
					{
						disposable2.Dispose();
					}
				}
			}
		}
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Shape A, ReplaceFontsOptions B)
	{
		Fonts.A(A.TextFrame2.TextRange, B);
	}

	private static void A(TextRange2 A, ReplaceFontsOptions B)
	{
		if (!B.A)
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
			if (!B.B)
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
				if (!B.C)
				{
					if (B.Find.Length == 0)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								A.Font.Name = B.Replace;
								return;
							}
						}
					}
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = A.get_Runs(-1, -1).GetEnumerator();
						while (enumerator.MoveNext())
						{
							TextRange2 textRange = (TextRange2)enumerator.Current;
							if (Operators.CompareString(textRange.Font.Name, B.Find, TextCompare: false) == 0)
							{
								textRange.Font.Name = B.Replace;
							}
						}
						while (true)
						{
							switch (7)
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
			}
		}
		IEnumerator enumerator2 = default(IEnumerator);
		try
		{
			Font2 font;
			for (enumerator2 = A.get_Runs(-1, -1).GetEnumerator(); enumerator2.MoveNext(); font = null)
			{
				font = ((TextRange2)enumerator2.Current).Font;
				if (B.Find.Length != 0 && Operators.CompareString(font.Name, B.Find, TextCompare: false) != 0)
				{
					continue;
				}
				if (B.A)
				{
					if (font.Bold != MsoTriState.msoTrue)
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
				}
				if (B.B)
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
					if (font.Italic != MsoTriState.msoTrue)
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
				}
				if (B.C)
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
					if (font.UnderlineStyle == MsoTextUnderlineType.msoNoUnderline)
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
				}
				font.Name = B.Replace;
			}
			while (true)
			{
				switch (7)
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
	}
}
