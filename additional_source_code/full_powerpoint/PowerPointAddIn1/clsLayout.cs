using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1;

public sealed class clsLayout
{
	private struct XF
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public float A;

		public float B;

		public float C;

		public float D;

		public bool A;

		public bool B;
	}

	public static readonly string DUMMY_TEXT = AH.A(151924);

	private static readonly string m_A = AH.A(151931);

	private static readonly string B = AH.A(151950);

	private static readonly string C = AH.A(151971);

	private static readonly string D = AH.A(151996);

	public static void ReapplyLayout()
	{
	}

	public static void DoLayout(CustomLayout oLayout = null, ListView lv = null)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Slide slide = null;
		SlideRange slideRange = application.ActiveWindow.Selection.SlideRange;
		int slideIndex = slideRange[1].SlideIndex;
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			Microsoft.Office.Interop.PowerPoint.Shape shape2;
			Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
			Microsoft.Office.Interop.PowerPoint.Shape shape4;
			try
			{
				enumerator = slideRange.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				IEnumerator enumerator3 = default(IEnumerator);
				XF item = default(XF);
				IEnumerator enumerator4 = default(IEnumerator);
				IEnumerator enumerator5 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					Slide slide2 = (Slide)enumerator.Current;
					List<XF> list = new List<XF>();
					application.ActiveWindow.View.GotoSlide(slide2.SlideIndex);
					if (!PB.Settings.ConformLayoutInsertMissingPlaceholders)
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
						slide = slide2.Duplicate()[1];
					}
					Slide slide3 = application.ActivePresentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
					shapeRange = slide2.Shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value));
					{
						enumerator2 = shapeRange.GetEnumerator();
						try
						{
							while (enumerator2.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
								bool a = false;
								if (lv != null)
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
										enumerator3 = lv.Items.GetEnumerator();
										while (enumerator3.MoveNext())
										{
											ListViewItem listViewItem = (ListViewItem)enumerator3.Current;
											if ((Conversions.ToInteger(listViewItem.SubItems[0].Text) == slide2.SlideID) & (Operators.CompareString(listViewItem.SubItems[1].Text, shape.Name, TextCompare: false) == 0))
											{
												shape.Tags.Add(AH.A(151864), AH.A(9078));
												a = true;
												break;
											}
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
								}
								if (shape.Type == MsoShapeType.msoPlaceholder)
								{
									shape2 = TempShape(slide2, slide3, shape);
									shape2.Tags.Add(AH.A(151864), AH.A(9078));
									Microsoft.Office.Interop.PowerPoint.Shape shape3 = slide2.Shapes[slide2.Shapes.Count];
									item.A = shape2;
									item.A = shape3.Top;
									item.B = shape3.Left;
									item.C = shape3.Height;
									item.D = shape3.Width;
									item.A = a;
									list.Add(item);
								}
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_02bc;
								}
								continue;
								end_IL_02bc:
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
					slide3.Delete();
					slide3 = null;
					if (oLayout == null)
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
						slide2.CustomLayout = slide2.CustomLayout;
					}
					else
					{
						slide2.CustomLayout = oLayout;
					}
					if (PB.Settings.ConformLayoutDeleteSuperfluous)
					{
						for (int i = slide2.Shapes.Count; i >= 1; i += -1)
						{
							shape4 = slide2.Shapes[i];
							bool a = false;
							try
							{
								a = Conversions.ToBoolean(shape4.Tags[AH.A(151864)]);
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
							if (a)
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
							try
							{
								enumerator4 = slide2.CustomLayout.Shapes.GetEnumerator();
								while (enumerator4.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape shp = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current;
									if (!Helpers.IsShapeMatch(shape4, shp))
									{
										continue;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_03ca;
										}
										continue;
										end_IL_03ca:
										break;
									}
									goto IL_040a;
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
							shape4.Delete();
							IL_040a:;
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
					int num = list.Count - 1;
					for (int j = 0; j <= num; j++)
					{
						XF xF = list[j];
						shape4 = xF.A;
						float a2 = xF.A;
						float b = xF.B;
						float c = xF.C;
						float d = xF.D;
						bool a = xF.A;
						unchecked
						{
							try
							{
								enumerator5 = slide2.Shapes.Placeholders.GetEnumerator();
								while (true)
								{
									if (enumerator5.MoveNext())
									{
										Microsoft.Office.Interop.PowerPoint.Shape shape5 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator5.Current;
										Microsoft.Office.Interop.PowerPoint.Shape shape6 = shape5;
										if ((shape6.Top == a2) & (shape6.Left == b) & (shape6.Width == d) & (shape6.Height == c))
										{
											while (true)
											{
												switch (5)
												{
												case 0:
													continue;
												}
												shape6.Name = shape4.Name;
												if (!PB.Settings.ConformLayoutSizesAndPosition || a)
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
													shape6.Top = shape4.Top;
													shape6.Left = shape4.Left;
													shape6.Width = shape4.Width;
													shape6.Height = shape4.Height;
												}
												RestoreContent(shape4, shape5);
												if (!PB.Settings.ConformLayoutFormats || a)
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
													A(shape4, shape5);
												}
												shape6.Tags.Add(AH.A(151893), AH.A(9078));
												break;
											}
											break;
										}
										shape6 = null;
										continue;
									}
									while (true)
									{
										switch (4)
										{
										case 0:
											break;
										default:
											goto end_IL_05b8;
										}
										continue;
										end_IL_05b8:
										break;
									}
									break;
								}
							}
							finally
							{
								if (enumerator5 is IDisposable)
								{
									while (true)
									{
										switch (6)
										{
										case 0:
											continue;
										}
										(enumerator5 as IDisposable).Dispose();
										break;
									}
								}
							}
							shape4.Delete();
						}
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
					Placeholders placeholders = slide2.Shapes.Placeholders;
					for (int k = placeholders.Count; k >= 1; k += -1)
					{
						if (!PB.Settings.ConformLayoutInsertMissingPlaceholders)
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
							bool a = false;
							try
							{
								a = Conversions.ToBoolean(placeholders[k].Tags[AH.A(151893)]);
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							if (!a)
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
								DeleteInsertedPlaceholders(slide, placeholders[k]);
							}
						}
						placeholders[k].Tags.Delete(AH.A(151893));
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
					placeholders = null;
					if (slide == null)
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
					slide.Delete();
					slide = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_06f8;
					}
					continue;
					end_IL_06f8:
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
			application.ActiveWindow.View.GotoSlide(slideIndex);
			shape2 = null;
			slideRange = null;
			shapeRange = null;
			shape4 = null;
			application = null;
		}
	}

	public static Microsoft.Office.Interop.PowerPoint.Shape TempShape(Slide sldOrig, Slide sldBlank, Microsoft.Office.Interop.PowerPoint.Shape shpSource)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = default(Microsoft.Office.Interop.PowerPoint.Shape);
		Microsoft.Office.Interop.PowerPoint.Shape shape3 = default(Microsoft.Office.Interop.PowerPoint.Shape);
		Microsoft.Office.Interop.PowerPoint.Shape result = default(Microsoft.Office.Interop.PowerPoint.Shape);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 355:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_000f;
						case 4:
							goto IL_0014;
						case 5:
							goto IL_001d;
						case 6:
							goto IL_003d;
						case 7:
							goto IL_004c;
						case 8:
							goto IL_0055;
						case 9:
							goto IL_0073;
						case 10:
							goto IL_007d;
						case 11:
							goto IL_0087;
						case 12:
							goto IL_008a;
						case 13:
							goto IL_00a9;
						case 14:
							goto IL_00b4;
						case 15:
							goto IL_00c5;
						case 16:
							goto IL_00d8;
						case 17:
							goto IL_00eb;
						case 18:
							goto IL_00f5;
						case 19:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 20:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00f5:
					shape = null;
					break;
					IL_0007:
					num2 = 2;
					DummyText(shpSource);
					goto IL_000f;
					IL_000f:
					num2 = 3;
					shape = shpSource;
					goto IL_0014;
					IL_0014:
					num2 = 4;
					shape.Copy();
					goto IL_001d;
					IL_001d:
					num2 = 5;
					shape2 = sldOrig.Shapes.Paste()[1];
					goto IL_003d;
					IL_003d:
					num2 = 6;
					if (shape2.Type == MsoShapeType.msoPlaceholder)
					{
						goto IL_004c;
					}
					goto IL_00a9;
					IL_004c:
					num2 = 7;
					shape2.Delete();
					goto IL_0055;
					IL_0055:
					num2 = 8;
					shape3 = sldBlank.Shapes.Paste()[1];
					goto IL_0073;
					IL_0073:
					num2 = 9;
					shape3.Copy();
					goto IL_007d;
					IL_007d:
					num2 = 10;
					shape3.Delete();
					goto IL_0087;
					IL_0087:
					shape3 = null;
					goto IL_008a;
					IL_008a:
					num2 = 12;
					shape2 = sldOrig.Shapes.Paste()[1];
					goto IL_00a9;
					IL_00a9:
					num2 = 13;
					A(shpSource, shape2);
					goto IL_00b4;
					IL_00b4:
					num2 = 14;
					shape2.Name = shape.Name;
					goto IL_00c5;
					IL_00c5:
					num2 = 15;
					shape2.Top = shape.Top;
					goto IL_00d8;
					IL_00d8:
					num2 = 16;
					shape2.Left = shape.Left;
					goto IL_00eb;
					IL_00eb:
					num2 = 17;
					shape.Delete();
					goto IL_00f5;
					end_IL_0000_2:
					break;
				}
				num2 = 19;
				result = shape2;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 355;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
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
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static void DummyText(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = shp;
		if ((shape.HasTextFrame == MsoTriState.msoTrue) & (shape.PlaceholderFormat.ContainedType == MsoShapeType.msoAutoShape))
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
			if (shape.TextFrame.HasText == MsoTriState.msoFalse)
			{
				shape.TextFrame.TextRange.Text = DUMMY_TEXT;
			}
		}
		shape = null;
	}

	public static void RestoreTags(Microsoft.Office.Interop.PowerPoint.Shape shpTemp, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		Tags tags = shpTemp.Tags;
		tags.Add(clsLayout.m_A, shp.Top.ToString());
		tags.Add(B, shp.Left.ToString());
		tags.Add(D, shp.Width.ToString());
		tags.Add(C, shp.Height.ToString());
		_ = null;
	}

	public static void RestoreContent(Microsoft.Office.Interop.PowerPoint.Shape oShp, Microsoft.Office.Interop.PowerPoint.Shape shpPlaceholder)
	{
		MsoShapeType type = oShp.Type;
		if (type != MsoShapeType.msoAutoShape)
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
			if (type != MsoShapeType.msoTextBox)
			{
				oShp.Copy();
				shpPlaceholder.Select();
				DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
				activeWindow.View.Paste();
				Microsoft.Office.Interop.PowerPoint.Shape shape = activeWindow.Selection.ShapeRange[1];
				shape.Top = oShp.Top;
				shape.Left = oShp.Left;
				_ = null;
				_ = null;
				return;
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
		oShp.TextFrame.TextRange.Copy();
		TextRange textRange = shpPlaceholder.TextFrame.TextRange;
		textRange.PasteSpecial(PpPasteDataType.ppPasteText);
		if (Operators.CompareString(textRange.Text, DUMMY_TEXT, TextCompare: false) == 0)
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
			textRange.Text = "";
		}
		textRange = null;
	}

	public static void DeleteInsertedPlaceholders(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape oShp)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = sld.Shapes.Placeholders.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shp = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (Helpers.IsShapeMatch(oShp, shp))
				{
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
		if (oShp.PlaceholderFormat.ContainedType != MsoShapeType.msoAutoShape)
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
			oShp.Delete();
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		BulletFormat bulletFormat = default(BulletFormat);
		TextRange textRange = default(TextRange);
		TextRange textRange2 = default(TextRange);
		int num5 = default(int);
		TextRange textRange3 = default(TextRange);
		int count = default(int);
		BulletFormat bulletFormat2 = default(BulletFormat);
		Font font = default(Font);
		int count2 = default(int);
		int num6 = default(int);
		ParagraphFormat paragraphFormat = default(ParagraphFormat);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 1117:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_000f;
						case 5:
							goto IL_0037;
						case 6:
							goto IL_003f;
						case 7:
							goto IL_0052;
						case 8:
							goto IL_006e;
						case 9:
							goto IL_007f;
						case 10:
							goto IL_009c;
						case 11:
							goto IL_00c8;
						case 12:
							goto IL_00e0;
						case 13:
							goto IL_00ed;
						case 14:
							goto IL_0105;
						case 15:
							goto IL_0118;
						case 16:
							goto IL_0129;
						case 17:
							goto IL_014f;
						case 18:
							goto IL_0169;
						case 19:
							goto IL_0183;
						case 20:
							goto IL_019d;
						case 21:
							goto IL_01b5;
						case 22:
							goto IL_01cb;
						case 23:
							goto IL_01e0;
						case 24:
							goto IL_0200;
						case 25:
							goto IL_0210;
						case 26:
							goto IL_0222;
						case 28:
							goto IL_0253;
						case 30:
							goto IL_0267;
						case 31:
							goto IL_0279;
						case 33:
							goto IL_028b;
						case 27:
						case 29:
						case 32:
						case 34:
						case 35:
						case 36:
							goto IL_0297;
						case 37:
							goto IL_02a9;
						case 38:
							goto IL_02b6;
						case 39:
							goto IL_02e2;
						case 40:
							goto IL_02f4;
						case 41:
							goto IL_030b;
						case 42:
							goto IL_032b;
						case 43:
							goto IL_032e;
						case 45:
							goto IL_0333;
						case 44:
						case 46:
							goto IL_0345;
						case 47:
							goto IL_0348;
						case 48:
							goto IL_035a;
						case 49:
							goto IL_0360;
						case 50:
							goto IL_0365;
						case 51:
							goto IL_036b;
						case 52:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 4:
						case 53:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_032b:
					bulletFormat = null;
					goto IL_032e;
					IL_0007:
					num2 = 2;
					A.PickUp();
					goto IL_000f;
					IL_000f:
					num2 = 3;
					if (Information.Err().Number != 0)
					{
						goto end_IL_0000_3;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_0037;
					IL_032e:
					textRange = null;
					goto IL_0345;
					IL_0345:
					textRange2 = null;
					goto IL_0348;
					IL_0348:
					num2 = 47;
					num5 = checked(num5 + 1);
					goto IL_0351;
					IL_0037:
					num2 = 5;
					B.Apply();
					goto IL_003f;
					IL_003f:
					num2 = 6;
					textRange3 = B.TextFrame.TextRange;
					goto IL_0052;
					IL_0052:
					num2 = 7;
					count = textRange3.Paragraphs().Count;
					num5 = 1;
					goto IL_0351;
					IL_0351:
					if (num5 <= count)
					{
						goto IL_006e;
					}
					goto IL_035a;
					IL_035a:
					num2 = 48;
					textRange3 = null;
					goto IL_0360;
					IL_0360:
					num2 = 49;
					bulletFormat2 = null;
					goto IL_0365;
					IL_0365:
					num2 = 50;
					font = null;
					goto IL_036b;
					IL_036b:
					num2 = 51;
					break;
					IL_006e:
					num2 = 8;
					textRange2 = A.TextFrame.TextRange;
					goto IL_007f;
					IL_007f:
					num2 = 9;
					count2 = textRange2.Paragraphs().Count;
					num6 = 1;
					goto IL_033c;
					IL_033c:
					if (num6 <= count2)
					{
						goto IL_009c;
					}
					goto IL_0345;
					IL_009c:
					num2 = 10;
					if (textRange3.Paragraphs(num5).IndentLevel == textRange2.Paragraphs(num6).IndentLevel)
					{
						goto IL_00c8;
					}
					goto IL_0333;
					IL_00c8:
					num2 = 11;
					paragraphFormat = textRange2.Paragraphs(num6).ParagraphFormat;
					goto IL_00e0;
					IL_00e0:
					num2 = 12;
					bulletFormat2 = paragraphFormat.Bullet;
					goto IL_00ed;
					IL_00ed:
					num2 = 13;
					font = ((TextRange)paragraphFormat.Parent).Font;
					goto IL_0105;
					IL_0105:
					num2 = 14;
					_ = A.TextFrame.Ruler;
					goto IL_0118;
					IL_0118:
					num2 = 15;
					textRange = textRange3.Paragraphs(num5);
					goto IL_0129;
					IL_0129:
					num2 = 16;
					textRange.Font.Color.RGB = font.Color.RGB;
					goto IL_014f;
					IL_014f:
					num2 = 17;
					textRange.Font.Size = font.Size;
					goto IL_0169;
					IL_0169:
					num2 = 18;
					textRange.Font.Name = font.Name;
					goto IL_0183;
					IL_0183:
					num2 = 19;
					textRange.Font.Bold = font.Bold;
					goto IL_019d;
					IL_019d:
					num2 = 20;
					textRange.Font.Italic = font.Italic;
					goto IL_01b5;
					IL_01b5:
					num2 = 21;
					textRange.Font.Underline = font.Underline;
					goto IL_01cb;
					IL_01cb:
					num2 = 22;
					bulletFormat = textRange.ParagraphFormat.Bullet;
					goto IL_01e0;
					IL_01e0:
					num2 = 23;
					bulletFormat.Font.Name = bulletFormat2.Font.Name;
					goto IL_0200;
					IL_0200:
					num2 = 24;
					bulletFormat.RelativeSize = bulletFormat2.RelativeSize;
					goto IL_0210;
					IL_0210:
					num2 = 25;
					bulletFormat.Type = bulletFormat2.Type;
					goto IL_0222;
					IL_0222:
					num2 = 26;
					switch (bulletFormat2.Type)
					{
					case PpBulletType.ppBulletUnnumbered:
						break;
					case PpBulletType.ppBulletNumbered:
						goto IL_0267;
					case PpBulletType.ppBulletMixed:
						goto IL_028b;
					default:
						goto IL_0297;
					}
					goto IL_0253;
					IL_028b:
					num2 = 33;
					bulletFormat.Style = PpNumberedBulletStyle.ppBulletStyleMixed;
					goto IL_0297;
					IL_0267:
					num2 = 30;
					bulletFormat.StartValue = bulletFormat2.StartValue;
					goto IL_0279;
					IL_0279:
					num2 = 31;
					bulletFormat.Style = bulletFormat2.Style;
					goto IL_0297;
					IL_0253:
					num2 = 28;
					bulletFormat.Character = bulletFormat2.Character;
					goto IL_0297;
					IL_0297:
					num2 = 36;
					bulletFormat.UseTextColor = bulletFormat2.UseTextColor;
					goto IL_02a9;
					IL_02a9:
					num2 = 37;
					if (bulletFormat2.UseTextColor == MsoTriState.msoFalse)
					{
						goto IL_02b6;
					}
					goto IL_02e2;
					IL_02b6:
					num2 = 38;
					bulletFormat.Font.Color.RGB = bulletFormat2.Font.Color.RGB;
					goto IL_02e2;
					IL_02e2:
					num2 = 39;
					bulletFormat.UseTextFont = bulletFormat2.UseTextFont;
					goto IL_02f4;
					IL_02f4:
					num2 = 40;
					if (bulletFormat2.UseTextFont == MsoTriState.msoFalse)
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
						goto IL_030b;
					}
					goto IL_032b;
					IL_0333:
					num2 = 45;
					num6 = checked(num6 + 1);
					goto IL_033c;
					IL_030b:
					num2 = 41;
					bulletFormat.Font.Name = bulletFormat2.Font.Name;
					goto IL_032b;
					end_IL_0000_2:
					break;
				}
				num2 = 52;
				paragraphFormat = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1117;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num != 0)
		{
			ProjectData.ClearProjectError();
		}
	}
}
