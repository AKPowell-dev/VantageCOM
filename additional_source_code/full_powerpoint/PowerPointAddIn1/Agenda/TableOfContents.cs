using System;
using System.Collections;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Agenda;

public sealed class TableOfContents
{
	internal static Slide A(bool A)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		Slide slide = null;
		Slide slide2 = null;
		CustomLayout customLayout = null;
		bool flag = true;
		try
		{
			presentation = application.ActivePresentation;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (presentation != null)
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
			slide2 = Slide(presentation);
			if (slide2 != null)
			{
				if (MessageBox.Show(AH.A(5734), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
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
					customLayout = slide2.CustomLayout;
				}
				else
				{
					flag = false;
				}
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
				if (customLayout == null)
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
					customLayout = Helpers.GetLayout(presentation, SlideType.TableOfContents);
					if (customLayout == null)
					{
						customLayout = Helpers.GetLayout(presentation, SlideType.Agenda);
					}
				}
				if (customLayout != null)
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
					int num;
					if (Helpers.GetSlideType(presentation.Slides[1]) != SlideType.Title)
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
						num = 1;
					}
					else
					{
						num = 2;
					}
					int index = num;
					if (!A)
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
						application.StartNewUndoEntry();
					}
					slide = presentation.Slides.AddSlide(index, customLayout);
					slide.Select();
					if (slide2 != null)
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
						slide2.Delete();
					}
					try
					{
						Update.A(A: true, presentation);
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						Forms.ErrorMessage(AH.A(5893) + ex4.Message);
						ProjectData.ClearProjectError();
					}
					customLayout = null;
				}
				else
				{
					Forms.WarningMessage(AH.A(6010));
				}
			}
			slide2 = null;
			presentation = null;
		}
		application = null;
		return slide;
	}

	public static Slide Slide(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		Slide result = null;
		try
		{
			int num;
			if (pres.SectionProperties.Count > 1)
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
				num = Math.Min(10, pres.SectionProperties.SlidesCount(1));
			}
			else
			{
				num = 10;
			}
			int num2 = num;
			int num3 = 1;
			while (true)
			{
				if (num3 <= num2)
				{
					Slide slide = pres.Slides[num3];
					SlideType slideType = Helpers.GetSlideType(slide);
					if (slideType != SlideType.TableOfContents)
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
						if (slideType != SlideType.Agenda)
						{
							num3 = checked(num3 + 1);
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
					}
					result = slide;
					break;
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_0095;
					}
					continue;
					end_IL_0095:
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
		return result;
	}

	internal static void A(int A)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		bool flag = false;
		if (clsRibbon.CallbackSlideView(ShowWarning: true))
		{
			try
			{
				Selection selection = application.ActiveWindow.Selection;
				if (selection.ShapeRange.Count != 1)
				{
					Forms.WarningMessage(AH.A(6621));
					goto IL_0382;
				}
				if (Helpers.IsSpecialSlide(selection.SlideRange[1]))
				{
					Forms.WarningMessage(AH.A(6532));
					goto IL_0382;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
				MsoShapeType type = selection.ShapeRange[1].Type;
				if (type == MsoShapeType.msoAutoShape)
				{
					goto IL_00ca;
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
				if (type == MsoShapeType.msoPlaceholder)
				{
					goto IL_00ca;
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
				if (type == MsoShapeType.msoTextBox)
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
					goto IL_00ca;
				}
				Forms.WarningMessage(AH.A(6428));
				goto end_IL_001a;
				IL_0382:
				selection = null;
				goto end_IL_001a;
				IL_00ca:
				application.StartNewUndoEntry();
				if (MessageBox.Show(AH.A(6294), AH.A(5874), MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
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
					Microsoft.Office.Interop.PowerPoint.Shape shape = selection.ShapeRange[1];
					float top = shape.Top;
					float left = shape.Left;
					float height = shape.Height;
					float width = shape.Width;
					_ = null;
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = activePresentation.Slides.GetEnumerator();
						IEnumerator enumerator2 = default(IEnumerator);
						while (enumerator.MoveNext())
						{
							Slide slide = (Slide)enumerator.Current;
							if (Helpers.GetSlideType(slide) == SlideType.TableOfContents)
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
								flag = true;
							}
							else
							{
								if (Helpers.IsSpecialSlide(slide))
								{
									continue;
								}
								try
								{
									enumerator2 = slide.Shapes.GetEnumerator();
									while (enumerator2.MoveNext())
									{
										Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
										Microsoft.Office.Interop.PowerPoint.Shape shape3 = shape2;
										if (shape3.Top == top)
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
											if (shape3.Left == left && shape3.Width == width)
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
												if (shape3.Height == height)
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
													TableOfContents.A(shape2, A);
												}
											}
										}
										shape3 = null;
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_021a;
										}
										continue;
										end_IL_021a:
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
							}
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0254;
							}
							continue;
							end_IL_0254:
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
				else
				{
					foreach (Slide slide2 in activePresentation.Slides)
					{
						SlideType slideType = Helpers.GetSlideType(slide2);
						if (slideType == SlideType.TableOfContents)
						{
							flag = true;
							break;
						}
					}
					TableOfContents.A(selection.ShapeRange[1], A);
				}
				if (!flag)
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
					TableOfContents.A(A: true);
				}
				else
				{
					try
					{
						Update.RefreshTopicStyle(activePresentation);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				activePresentation = null;
				TableOfContents.A(AH.A(6467) + A);
				goto IL_0382;
				end_IL_001a:;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		application = null;
	}

	internal static void A()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		if (clsRibbon.CallbackSlideView(ShowWarning: true))
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
				Selection selection = application.ActiveWindow.Selection;
				if (selection.ShapeRange.Count == 1)
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
					if (!Helpers.IsSpecialSlide(selection.SlideRange[1]))
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
						Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
						application.StartNewUndoEntry();
						if (MessageBox.Show(AH.A(6294), AH.A(5874), MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
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
							Microsoft.Office.Interop.PowerPoint.Shape shape = selection.ShapeRange[1];
							float top = shape.Top;
							float left = shape.Left;
							float height = shape.Height;
							float width = shape.Width;
							_ = null;
							IEnumerator enumerator2 = default(IEnumerator);
							foreach (Slide slide in activePresentation.Slides)
							{
								if (Helpers.IsSpecialSlide(slide))
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
								try
								{
									enumerator2 = slide.Shapes.GetEnumerator();
									while (enumerator2.MoveNext())
									{
										Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
										Microsoft.Office.Interop.PowerPoint.Shape shape3 = shape2;
										if (shape3.Top == top)
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
											if (shape3.Left == left)
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
												if (shape3.Width == width)
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
													if (shape3.Height == height)
													{
														A(shape2);
													}
												}
											}
										}
										shape3 = null;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_01c4;
										}
										continue;
										end_IL_01c4:
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
							}
						}
						else
						{
							A(selection.ShapeRange[1]);
						}
						try
						{
							Update.RefreshTopicStyle(activePresentation);
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						activePresentation = null;
						A(AH.A(6738));
					}
					else
					{
						Forms.WarningMessage(AH.A(6811));
					}
				}
				else
				{
					Forms.WarningMessage(AH.A(6900));
				}
				selection = null;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		application = null;
	}

	internal static void B()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
		application.StartNewUndoEntry();
		_ = null;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = activePresentation.Slides.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Slide slide = (Slide)enumerator.Current;
				try
				{
					enumerator2 = slide.Shapes.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current);
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_009d;
				}
				continue;
				end_IL_009d:
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
		Update.RefreshTopicStyle(activePresentation);
		activePresentation = null;
		A(AH.A(7021));
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, int B)
	{
		A.Tags.Add(Constants.TAG_TOC_LEVEL, B.ToString());
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		try
		{
			A.Tags.Delete(Constants.TAG_TOC_LEVEL);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A(string A)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)11, A);
	}
}
