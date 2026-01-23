using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.TextOps;

public sealed class TextBox
{
	internal static void A()
	{
		if (!Licensing.AllowAdvancedTextOperation())
		{
			return;
		}
		checked
		{
			Selection selection = default(Selection);
			Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
			List<string> list = default(List<string>);
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator enumerator2 = default(IEnumerator);
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
				Application application = NG.A.Application;
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
				try
				{
					selection = application.ActiveWindow.Selection;
					if (selection.Type == PpSelectionType.ppSelectionShapes)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							shapeRange = PowerPointAddIn1.Shapes.Base.SelectedShapes(selection);
							if (shapeRange.Count > 1)
							{
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									shape = shapeRange[1];
									if (shape.HasTextFrame == MsoTriState.msoTrue)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											list = new List<string>();
											application.StartNewUndoEntry();
											TextRange2 textRange = shape.TextFrame2.TextRange;
											textRange.InsertAfter(AH.A(7894));
											int count = shapeRange.Count;
											for (int i = 2; i <= count; i++)
											{
												int length = textRange.Text.Length;
												try
												{
													Microsoft.Office.Interop.PowerPoint.Shape shape2 = shapeRange[i];
													if (shape2.HasTextFrame != MsoTriState.msoTrue)
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
														if (shape2.TextFrame2.HasText != MsoTriState.msoTrue)
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
															try
															{
																enumerator = shape2.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
																while (enumerator.MoveNext())
																{
																	TextRange2 textRange2 = (TextRange2)enumerator.Current;
																	textRange.InsertAfter(textRange2.Text);
																	textRange.get_Paragraphs(textRange.get_Paragraphs(-1, -1).Count, -1).ParagraphFormat.IndentLevel = textRange2.ParagraphFormat.IndentLevel;
																	_ = null;
																	try
																	{
																		enumerator2 = textRange2.get_Runs(-1, -1).GetEnumerator();
																		while (enumerator2.MoveNext())
																		{
																			TextRange2 textRange3 = (TextRange2)enumerator2.Current;
																			Font2 font = textRange3.Font;
																			try
																			{
																				Font2 font2 = textRange.get_Characters(length + textRange3.Start, textRange3.Length).Font;
																				if (font.Highlight.RGB != 0 && font.Highlight.Type == MsoColorType.msoColorTypeRGB)
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
																					font2.Highlight.RGB = font.Highlight.RGB;
																				}
																				font2.Name = font.Name;
																				font2.Size = font.Size;
																				font2.Bold = font.Bold;
																				font2.Italic = font.Italic;
																				font2.UnderlineStyle = font.UnderlineStyle;
																				font2.Fill.ForeColor.RGB = font.Fill.ForeColor.RGB;
																				font2.Superscript = font.Superscript;
																				font2.Subscript = font.Subscript;
																				font2 = null;
																			}
																			catch (Exception ex)
																			{
																				ProjectData.SetProjectError(ex);
																				Exception ex2 = ex;
																				ProjectData.ClearProjectError();
																			}
																			font = null;
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
																	switch (3)
																	{
																	case 0:
																		break;
																	default:
																		goto end_IL_0345;
																	}
																	continue;
																	end_IL_0345:
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
															list.Add(shape2.Name);
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
											}
											textRange = null;
											if (list.Count > 0)
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
												application.ActiveWindow.Selection.SlideRange[1].Shapes.Range(list.ToArray()).Delete();
											}
											Base.LogActivity(AH.A(157008));
											break;
										}
									}
									else
									{
										Forms.WarningMessage(AH.A(157029));
									}
									break;
								}
							}
							else
							{
								Helpers.TwoOrMoreShapesRequiredError();
							}
							break;
						}
					}
					else
					{
						Helpers.TwoOrMoreShapesRequiredError();
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
				application = null;
				shapeRange = null;
				JG.A(selection);
				JG.A(shape);
				JG.A(list);
				return;
			}
		}
	}

	internal static void B()
	{
		if (!Licensing.AllowAdvancedTextOperation())
		{
			return;
		}
		checked
		{
			Selection selection = default(Selection);
			List<Microsoft.Office.Interop.PowerPoint.Shape> list = default(List<Microsoft.Office.Interop.PowerPoint.Shape>);
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
				Application application = NG.A.Application;
				Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
				try
				{
					selection = application.ActiveWindow.Selection;
					Microsoft.Office.Interop.PowerPoint.Shape shape2;
					Microsoft.Office.Interop.PowerPoint.Shape shape;
					if (selection.Type == PpSelectionType.ppSelectionShapes)
					{
						shapeRange = PowerPointAddIn1.Shapes.Base.SelectedShapes(selection);
						if (shapeRange.Count == 1)
						{
							shape = shapeRange[1];
							shape2 = shape;
							if (shape2.HasTextFrame == MsoTriState.msoTrue)
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
								if (shape2.TextFrame2.HasText == MsoTriState.msoTrue)
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
									if (shape2.TextFrame2.TextRange.get_Paragraphs(-1, -1).Count > 1)
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
										if (shape2.Type != MsoShapeType.msoPlaceholder)
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
											float left = shape2.Left;
											list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
											list.Add(shape);
											application.StartNewUndoEntry();
											Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = shape2.TextFrame2;
											float marginTop = textFrame.MarginTop;
											MsoVerticalAnchor verticalAnchor = textFrame.VerticalAnchor;
											textFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
											TextRange2 textRange = textFrame.TextRange.get_Paragraphs(-1, -1);
											for (int i = textRange.Count; i >= 2; i += -1)
											{
												TextRange2 textRange2 = textRange.Item(i);
												Microsoft.Office.Interop.PowerPoint.Shape shape3 = shape.Duplicate()[1];
												Microsoft.Office.Interop.PowerPoint.Shape shape4 = shape3;
												TextRange2 textRange3 = shape4.TextFrame2.TextRange.get_Paragraphs(-1, -1);
												for (int j = textRange3.Count - 1; j >= 1; j += -1)
												{
													textRange3.Item(j).Delete();
												}
												while (true)
												{
													switch (6)
													{
													case 0:
														break;
													default:
														goto end_IL_01c5;
													}
													continue;
													end_IL_01c5:
													break;
												}
												textRange3 = null;
												shape4.Top = textRange2.BoundTop + textRange2.ParagraphFormat.SpaceBefore - marginTop;
												shape4.Left = left;
												shape4 = null;
												A(shape3, verticalAnchor);
												list.Add(shape3);
												shape3 = null;
												textRange2.Delete();
											}
											textRange = null;
											_ = null;
											A(shape, verticalAnchor);
											selection.SlideRange[1].Shapes.Range(list.Select([SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape A) => A.ZOrderPosition).ToArray()).Select();
											_ = shape.ZOrderPosition;
											for (int num = list.Count - 1; num >= 1; num += -1)
											{
												int num2 = num;
												for (int num3 = 2; num3 <= num2; num3++)
												{
													list[num].ZOrder(MsoZOrderCmd.msoSendBackward);
												}
												while (true)
												{
													switch (5)
													{
													case 0:
														break;
													default:
														goto end_IL_02cc;
													}
													continue;
													end_IL_02cc:
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
											Base.LogActivity(AH.A(157126));
										}
										else
										{
											Forms.WarningMessage(AH.A(157147));
										}
									}
									else
									{
										Forms.WarningMessage(AH.A(157242));
									}
									goto IL_0333;
								}
							}
							Forms.WarningMessage(AH.A(157337));
							goto IL_0333;
						}
						Forms.WarningMessage(AH.A(156221));
					}
					goto end_IL_002a;
					IL_0333:
					shape2 = null;
					shape = null;
					end_IL_002a:;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				application = null;
				shapeRange = null;
				JG.A(selection);
				JG.A(list);
				return;
			}
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, MsoVerticalAnchor B)
	{
		Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = A.TextFrame2;
		TextRange2 textRange = textFrame.TextRange;
		if (textRange.Text.Last() == '\r')
		{
			textRange.get_Characters(textRange.Text.Length, -1).Delete();
		}
		textRange = null;
		if (textFrame.AutoSize != MsoAutoSize.msoAutoSizeShapeToFitText)
		{
			MsoAutoSize autoSize = textFrame.AutoSize;
			textFrame.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
			textFrame.AutoSize = autoSize;
		}
		textFrame.VerticalAnchor = B;
		textFrame = null;
	}

	internal static void C()
	{
		if (!Licensing.AllowAdvancedTextOperation())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Application application = NG.A.Application;
			Settings settings;
			Selection selection;
			try
			{
				selection = application.ActiveWindow.Selection;
				if (selection.Type != PpSelectionType.ppSelectionShapes)
				{
					throw new Exception();
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					Settings.Margins? b = null;
					settings = new Settings(application.ActivePresentation);
					if (settings.TextboxMargins.HasValue)
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
						b = settings.TextboxMargins.Value;
					}
					else
					{
						try
						{
							Microsoft.Office.Interop.PowerPoint.Shape bodyPlaceholder = Helpers.GetBodyPlaceholder(application.ActivePresentation);
							if (bodyPlaceholder != null)
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
								b = new Settings.Margins
								{
									Top = bodyPlaceholder.TextFrame2.MarginTop,
									Bottom = bodyPlaceholder.TextFrame2.MarginBottom,
									Left = bodyPlaceholder.TextFrame2.MarginLeft,
									Right = bodyPlaceholder.TextFrame2.MarginRight
								};
							}
							bodyPlaceholder = null;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
					}
					enumerator = PowerPointAddIn1.Shapes.Base.SelectedShapes(selection).GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, b);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_0176;
							}
							continue;
							end_IL_0176:
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
					Base.LogActivity(AH.A(157410));
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.WarningMessage(AH.A(81244));
				ProjectData.ClearProjectError();
			}
			settings = null;
			selection = null;
			application = null;
			return;
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Settings.Margins? B)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			if (A.Type != MsoShapeType.msoGroup)
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
						if (A.HasTextFrame == MsoTriState.msoTrue)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									TextBox.B(A, B);
									return;
								}
							}
						}
						if (A.HasTable == MsoTriState.msoTrue)
						{
							while (true)
							{
								switch (3)
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
											Cell cell = table.Cell(i, j);
											if (cell.Selected)
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
												TextBox.B(cell.Shape, B);
											}
											cell = null;
										}
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_00c3;
											}
											continue;
											end_IL_00c3:
											break;
										}
									}
									while (true)
									{
										switch (5)
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
						if (A.HasSmartArt == MsoTriState.msoTrue)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									try
									{
										enumerator = A.SmartArt.AllNodes.GetEnumerator();
										while (enumerator.MoveNext())
										{
											SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
											try
											{
												Microsoft.Office.Core.TextFrame2 textFrame = smartArtNode.TextFrame2;
												if (B.HasValue)
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
													Settings.Margins value = B.Value;
													if (textFrame.MarginBottom == value.Bottom)
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
														if (textFrame.MarginLeft == value.Left)
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
															if (textFrame.MarginRight == value.Right)
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
																if (textFrame.MarginTop == value.Top)
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
																	textFrame.MarginBottom = 0f;
																	textFrame.MarginLeft = 0f;
																	textFrame.MarginRight = 0f;
																	textFrame.MarginTop = 0f;
																	goto IL_0260;
																}
															}
														}
													}
													textFrame.MarginBottom = value.Bottom;
													textFrame.MarginLeft = value.Left;
													textFrame.MarginRight = value.Right;
													textFrame.MarginTop = value.Top;
												}
												else
												{
													textFrame.MarginBottom = 0f;
													textFrame.MarginLeft = 0f;
													textFrame.MarginRight = 0f;
													textFrame.MarginTop = 0f;
												}
												goto IL_0260;
												IL_0260:
												textFrame = null;
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
													break;
												default:
													(enumerator as IDisposable).Dispose();
													goto end_IL_029a;
												}
												continue;
												end_IL_029a:
												break;
											}
										}
									}
								}
							}
						}
						return;
					}
				}
			}
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = A.GroupItems.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					TextBox.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, B);
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
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Shape A, Settings.Margins? B)
	{
		Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = A.TextFrame2;
		if (B.HasValue)
		{
			Settings.Margins value = B.Value;
			if (textFrame.MarginBottom == value.Bottom)
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
				if (textFrame.MarginLeft == value.Left)
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
					if (textFrame.MarginRight == value.Right && textFrame.MarginTop == value.Top)
					{
						textFrame.MarginBottom = 0f;
						textFrame.MarginLeft = 0f;
						textFrame.MarginRight = 0f;
						textFrame.MarginTop = 0f;
						goto IL_0107;
					}
				}
			}
			textFrame.MarginBottom = value.Bottom;
			textFrame.MarginLeft = value.Left;
			textFrame.MarginRight = value.Right;
			textFrame.MarginTop = value.Top;
		}
		else
		{
			textFrame.MarginBottom = 0f;
			textFrame.MarginLeft = 0f;
			textFrame.MarginRight = 0f;
			textFrame.MarginTop = 0f;
		}
		goto IL_0107;
		IL_0107:
		textFrame = null;
	}
}
