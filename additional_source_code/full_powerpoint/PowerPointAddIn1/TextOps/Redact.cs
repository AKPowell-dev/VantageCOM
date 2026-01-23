using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Shapes;

namespace PowerPointAddIn1.TextOps;

public sealed class Redact
{
	public static void RedactSelection()
	{
		if (!Licensing.AllowAdvancedTextOperation())
		{
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Selection selection = application.ActiveWindow.Selection;
		if (selection.Type == PpSelectionType.ppSelectionText)
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
			application.StartNewUndoEntry();
			TextRange2 textRange;
			try
			{
				textRange = selection.TextRange2;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = selection.TextRange2.get_Words(-1, -1).GetEnumerator();
					while (enumerator.MoveNext())
					{
						RedactWord((TextRange2)enumerator.Current);
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_0090;
						}
						continue;
						end_IL_0090:
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
				textRange.Select();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, AH.A(155652));
			textRange = null;
		}
		else if (selection.Type == PpSelectionType.ppSelectionShapes)
		{
			try
			{
				if (selection.HasChildShapeRange)
				{
					IEnumerator enumerator2 = default(IEnumerator);
					try
					{
						enumerator2 = selection.ChildShapeRange.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, application, selection);
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0149;
							}
							continue;
							end_IL_0149:
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
				else
				{
					IEnumerator enumerator3 = default(IEnumerator);
					try
					{
						enumerator3 = selection.ShapeRange.GetEnumerator();
						while (enumerator3.MoveNext())
						{
							A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current, application, selection);
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_01a8;
							}
							continue;
							end_IL_01a8:
							break;
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
				}
			}
			catch (NotImplementedException ex3)
			{
				ProjectData.SetProjectError(ex3);
				NotImplementedException ex4 = ex3;
				Forms.ErrorMessage(ex4.Message);
				ProjectData.ClearProjectError();
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				Forms.ErrorMessage(ex6.Message);
				ProjectData.ClearProjectError();
			}
		}
		selection = null;
		application = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Application B, Selection C)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator enumerator2 = default(IEnumerator);
			IEnumerator enumerator3 = default(IEnumerator);
			IEnumerator enumerator4 = default(IEnumerator);
			IEnumerator enumerator5 = default(IEnumerator);
			IEnumerator enumerator6 = default(IEnumerator);
			if (A.Type != MsoShapeType.msoGroup)
			{
				while (true)
				{
					float top;
					float left;
					int zOrderPosition;
					Microsoft.Office.Interop.PowerPoint.Shape shape;
					Microsoft.Office.Interop.PowerPoint.Shape shape3;
					switch (2)
					{
					case 0:
						break;
					default:
						{
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							shape = A;
							if (Images.HasPictureOrGraphic(A))
							{
								top = shape.Top;
								left = shape.Left;
								zOrderPosition = shape.ZOrderPosition;
								shape.Fill.ForeColor.RGB = 0;
								shape.Fill.BackColor.RGB = 0;
								shape.PictureFormat.Brightness = 0f;
								shape.Copy();
								if (Images.HasGraphic(A))
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
									if (A.Type == MsoShapeType.msoPlaceholder)
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
										C.SlideRange[1].Shapes.PasteSpecial(PpPasteDataType.ppPastePNG).Select();
										goto IL_011d;
									}
								}
								C.SlideRange[1].Shapes.PasteSpecial(PpPasteDataType.ppPasteEnhancedMetafile).Select();
								goto IL_011d;
							}
							if (shape.HasChart == MsoTriState.msoTrue)
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
								if (Forms.OkCancelMessage2(AH.A(155685)) == DialogResult.OK)
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
									Microsoft.Office.Interop.PowerPoint.ChartArea chartArea = shape.Chart.ChartArea;
									chartArea.ClearContents();
									chartArea.Format.Fill.ForeColor.RGB = 0;
									chartArea.Format.Fill.BackColor.RGB = 0;
									_ = null;
									ChartData chartData = shape.Chart.ChartData;
									Workbook workbook = (Workbook)chartData.Workbook;
									{
										enumerator = workbook.Worksheets.GetEnumerator();
										try
										{
											while (enumerator.MoveNext())
											{
												((Worksheet)enumerator.Current).Cells.Clear();
											}
											while (true)
											{
												switch (3)
												{
												case 0:
													break;
												default:
													goto end_IL_025d;
												}
												continue;
												end_IL_025d:
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
									workbook.Close(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									workbook = null;
									if (chartData.IsLinked)
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
										chartData.BreakLink();
									}
									chartData = null;
								}
							}
							else if (shape.HasTable == MsoTriState.msoTrue)
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
								int count = shape.Table.Rows.Count;
								for (int i = 1; i <= count; i++)
								{
									int count2 = shape.Table.Columns.Count;
									for (int j = 1; j <= count2; j++)
									{
										Cell cell = shape.Table.Cell(i, j);
										if (cell.Selected)
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
											{
												enumerator2 = cell.Shape.TextFrame2.TextRange.get_Words(-1, -1).GetEnumerator();
												try
												{
													while (enumerator2.MoveNext())
													{
														RedactWord((TextRange2)enumerator2.Current);
													}
													while (true)
													{
														switch (1)
														{
														case 0:
															break;
														default:
															goto end_IL_0390;
														}
														continue;
														end_IL_0390:
														break;
													}
												}
												finally
												{
													IDisposable disposable2 = enumerator2 as IDisposable;
													if (disposable2 != null)
													{
														disposable2.Dispose();
													}
												}
											}
										}
										cell = null;
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											break;
										default:
											goto end_IL_03c4;
										}
										continue;
										end_IL_03c4:
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
							}
							else if (shape.HasSmartArt == MsoTriState.msoTrue)
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
								{
									enumerator3 = shape.SmartArt.AllNodes.GetEnumerator();
									try
									{
										while (enumerator3.MoveNext())
										{
											SmartArtNode smartArtNode = (SmartArtNode)enumerator3.Current;
											try
											{
												enumerator4 = smartArtNode.Shapes.GetEnumerator();
												while (enumerator4.MoveNext())
												{
													Microsoft.Office.Core.Shape shape2 = (Microsoft.Office.Core.Shape)enumerator4.Current;
													if (shape2.TextFrame2.HasText == MsoTriState.msoTrue)
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
														{
															enumerator5 = shape2.TextFrame2.TextRange.get_Words(-1, -1).GetEnumerator();
															try
															{
																while (enumerator5.MoveNext())
																{
																	RedactWord((TextRange2)enumerator5.Current);
																}
																while (true)
																{
																	switch (6)
																	{
																	case 0:
																		break;
																	default:
																		goto end_IL_04b3;
																	}
																	continue;
																	end_IL_04b3:
																	break;
																}
															}
															finally
															{
																IDisposable disposable4 = enumerator5 as IDisposable;
																if (disposable4 != null)
																{
																	disposable4.Dispose();
																}
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
														goto end_IL_04e3;
													}
													continue;
													end_IL_04e3:
													break;
												}
											}
											finally
											{
												if (enumerator4 is IDisposable)
												{
													while (true)
													{
														switch (6)
														{
														case 0:
															break;
														default:
															(enumerator4 as IDisposable).Dispose();
															goto end_IL_04f8;
														}
														continue;
														end_IL_04f8:
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
												goto end_IL_051d;
											}
											continue;
											end_IL_051d:
											break;
										}
									}
									finally
									{
										IDisposable disposable3 = enumerator3 as IDisposable;
										if (disposable3 != null)
										{
											disposable3.Dispose();
										}
									}
								}
							}
							else if (shape.HasTextFrame == MsoTriState.msoTrue)
							{
								try
								{
									enumerator6 = shape.TextFrame2.TextRange.get_Words(-1, -1).GetEnumerator();
									while (enumerator6.MoveNext())
									{
										RedactWord((TextRange2)enumerator6.Current);
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
												break;
											default:
												(enumerator6 as IDisposable).Dispose();
												goto end_IL_0591;
											}
											continue;
											end_IL_0591:
											break;
										}
									}
								}
							}
							goto IL_05a8;
						}
						IL_05a8:
						shape = null;
						return;
						IL_011d:
						shape.Delete();
						shape3 = C.ShapeRange[1];
						shape3.Top = top;
						shape3.Left = left;
						while (shape3.ZOrderPosition < zOrderPosition)
						{
							shape3.ZOrder(MsoZOrderCmd.msoBringForward);
						}
						while (shape3.ZOrderPosition > zOrderPosition)
						{
							shape3.ZOrder(MsoZOrderCmd.msoSendBackward);
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
						shape3 = null;
						goto IL_05a8;
					}
				}
			}
			IEnumerator enumerator7 = default(IEnumerator);
			try
			{
				enumerator7 = A.GroupItems.GetEnumerator();
				while (enumerator7.MoveNext())
				{
					Redact.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator7.Current, B, C);
				}
			}
			finally
			{
				if (enumerator7 is IDisposable)
				{
					while (true)
					{
						switch (4)
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
	}

	public static void FindAndRedact()
	{
		if (!Licensing.AllowAdvancedTextOperation())
		{
			return;
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
			Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
			string text = "";
			if (application.Presentations.Count > 0)
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
					Selection selection = application.ActiveWindow.Selection;
					if (selection.Type == PpSelectionType.ppSelectionText)
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
						if (selection.TextRange2.get_Words(-1, -1).Count == 1)
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
							text = selection.TextRange2.Text.Trim();
						}
					}
					selection = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				wpfRedact obj = new wpfRedact();
				obj.txtFind.Text = text;
				obj.txtFind.Focus();
				obj.txtFind.SelectAll();
				obj.Show();
				_ = null;
			}
			application = null;
			return;
		}
	}

	public static void RedactWord(TextRange2 rngWord)
	{
		rngWord = rngWord.TrimText();
		if (rngWord.Text.Length == 0)
		{
			return;
		}
		checked
		{
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
				float boundWidth = rngWord.BoundWidth;
				float boundTop = rngWord.BoundTop;
				rngWord.Text = Strings.StrDup(rngWord.Text.Length, AH.A(58688));
				Font2 font = rngWord.Font;
				float boundWidth2 = rngWord.BoundWidth;
				float boundTop2 = rngWord.BoundTop;
				int num = 0;
				if (boundWidth2 > boundWidth || boundTop2 > boundTop)
				{
					while (true)
					{
						font.Spacing -= 0.5f;
						boundWidth2 = rngWord.BoundWidth;
						boundTop2 = rngWord.BoundTop;
						num++;
						if (boundWidth2 <= boundWidth)
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
							if (boundTop2 == boundTop)
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
								break;
							}
						}
						if (num <= 20)
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
						break;
					}
					num = 0;
					while (boundWidth2 < boundWidth)
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
						if (num <= 10)
						{
							font.Spacing += 0.05f;
							boundWidth2 = rngWord.BoundWidth;
							num++;
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
						break;
					}
				}
				else if (boundWidth2 < boundWidth || boundTop2 < boundTop)
				{
					while (true)
					{
						font.Spacing += 0.5f;
						boundWidth2 = rngWord.BoundWidth;
						boundTop2 = rngWord.BoundTop;
						num++;
						if (boundWidth2 >= boundWidth)
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
							if (boundTop2 == boundTop)
							{
								break;
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
						if (num <= 20)
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
						break;
					}
					num = 0;
					while (boundWidth2 > boundWidth)
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
						if (num <= 10)
						{
							font.Spacing -= 0.05f;
							boundWidth2 = rngWord.BoundWidth;
							num++;
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
						break;
					}
				}
				Font2 font2 = font;
				float size = font2.Size;
				string name = font2.Name;
				int rGB = font2.Fill.ForeColor.RGB;
				font2.Highlight.RGB = rGB;
				font2.Size = size;
				font2.Name = name;
				font2.Fill.ForeColor.RGB = rGB;
				_ = null;
				font = null;
				return;
			}
		}
	}
}
