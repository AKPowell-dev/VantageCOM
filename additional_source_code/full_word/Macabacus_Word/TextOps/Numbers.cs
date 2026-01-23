using System;
using System.Globalization;
using System.Text.RegularExpressions;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.TextOps;

public sealed class Numbers
{
	public static void NumberToPlainWords()
	{
		A(A: false);
	}

	public static void NumberToLegalWords()
	{
		A(A: true);
	}

	private static void A(bool A)
	{
		Application application = PC.A.Application;
		UndoRecord undoRecord = null;
		application.ScreenUpdating = false;
		try
		{
			Selection selection = application.Selection;
			object Extend;
			object Count;
			object Unit;
			if (selection.Type == WdSelectionType.wdSelectionIP)
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
				Selection selection2 = selection;
				Unit = WdUnits.wdWord;
				Count = 1;
				Extend = WdMovementType.wdMove;
				selection2.MoveLeft(ref Unit, ref Count, ref Extend);
				Selection selection3 = selection;
				Extend = WdUnits.wdWord;
				Count = 1;
				Unit = WdMovementType.wdExtend;
				selection3.MoveRight(ref Extend, ref Count, ref Unit);
			}
			Selection selection4 = selection;
			Unit = XC.A(18458);
			Count = WdConstants.wdBackward;
			selection4.MoveEndWhile(ref Unit, ref Count);
			Selection selection5 = selection;
			Count = ' ';
			Unit = WdConstants.wdBackward;
			selection5.MoveEndWhile(ref Count, ref Unit);
			Selection selection6 = selection;
			Unit = XC.A(20062);
			Count = WdConstants.wdBackward;
			selection6.MoveStartWhile(ref Unit, ref Count);
			Selection selection7 = selection;
			Count = XC.A(20065);
			Unit = WdConstants.wdBackward;
			selection7.MoveStartWhile(ref Count, ref Unit);
			Selection selection8 = selection;
			Unit = XC.A(20068);
			Count = WdConstants.wdBackward;
			selection8.MoveStartWhile(ref Unit, ref Count);
			Selection selection9 = selection;
			Count = XC.A(20071);
			Unit = WdConstants.wdBackward;
			selection9.MoveStartWhile(ref Count, ref Unit);
			Selection selection10 = selection;
			Unit = XC.A(20074);
			Count = WdConstants.wdBackward;
			selection10.MoveStartWhile(ref Unit, ref Count);
			Selection selection11 = selection;
			Count = XC.A(20077);
			Unit = WdConstants.wdBackward;
			selection11.MoveStartWhile(ref Count, ref Unit);
			Selection selection12 = selection;
			Unit = XC.A(20080);
			Count = WdConstants.wdForward;
			selection12.MoveEndWhile(ref Unit, ref Count);
			Selection selection13 = selection;
			Count = XC.A(20083);
			Unit = WdConstants.wdForward;
			selection13.MoveEndWhile(ref Count, ref Unit);
			Selection selection14 = selection;
			Unit = XC.A(20086);
			Count = WdConstants.wdForward;
			selection14.MoveEndWhile(ref Unit, ref Count);
			Selection selection15 = selection;
			Count = XC.A(20089);
			Unit = WdConstants.wdForward;
			selection15.MoveEndWhile(ref Count, ref Unit);
			string text;
			double result;
			long num;
			string text2;
			bool flag;
			bool flag2;
			bool flag3;
			bool flag4;
			bool flag5;
			bool flag6;
			string text3;
			if (selection.FormattedText.Font.Superscript != 0)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_0270;
					}
					continue;
					end_IL_0270:
					break;
				}
			}
			else
			{
				text = selection.Text.Trim();
				selection = null;
				if (double.TryParse(B(text.Replace(CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator, "")).Replace(XC.A(20080), "").Replace(XC.A(20083), ""), out result))
				{
					num = checked((long)Conversion.Int(result));
					text2 = "";
					flag = false;
					flag2 = false;
					flag3 = false;
					flag4 = false;
					flag5 = false;
					flag6 = false;
					text3 = text;
					if (!text3.StartsWith(XC.A(20062)))
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
						if (!text3.StartsWith(XC.A(20171)))
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
							if (!text3.StartsWith(XC.A(20178)))
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
								if (!text3.StartsWith(XC.A(20185)))
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
									if (!text3.StartsWith(XC.A(20192)))
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
										if (!text3.StartsWith(XC.A(20199)))
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
											if (!text3.StartsWith(XC.A(20206)))
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
												if (!text3.StartsWith(XC.A(20213)))
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
													if (!text3.StartsWith(XC.A(20218)))
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
														if (!text3.StartsWith(XC.A(20223)))
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
															if (!text3.StartsWith(XC.A(20230)))
															{
																if (!text3.StartsWith(XC.A(20065)))
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
																	if (!text3.StartsWith(XC.A(20254)))
																	{
																		if (!text3.StartsWith(XC.A(20068)))
																		{
																			if (!text3.StartsWith(XC.A(20274)))
																			{
																				if (!text3.StartsWith(XC.A(20071)))
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
																					if (!text3.StartsWith(XC.A(20296)))
																					{
																						if (text3.StartsWith(XC.A(20312)) || text3.StartsWith(XC.A(20319)))
																						{
																							text2 = XC.A(20326);
																						}
																						else
																						{
																							if (!text3.StartsWith(XC.A(20074)))
																							{
																								if (!text3.StartsWith(XC.A(20337)))
																								{
																									if (text3.StartsWith(XC.A(20077), StringComparison.OrdinalIgnoreCase))
																									{
																										text2 = XC.A(20353);
																									}
																									else if (text3.StartsWith(XC.A(20370)))
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
																										text2 = XC.A(20377);
																									}
																									else
																									{
																										if (!text3.StartsWith(XC.A(20392)))
																										{
																											if (!text3.StartsWith(XC.A(20399)))
																											{
																												if (text3.StartsWith(XC.A(20421)))
																												{
																													text2 = XC.A(20428);
																												}
																												else
																												{
																													if (!text3.StartsWith(XC.A(20441)))
																													{
																														if (!text3.StartsWith(XC.A(20446)))
																														{
																															if (text3.EndsWith(XC.A(20086)))
																															{
																																text2 = XC.A(20464);
																																flag5 = true;
																															}
																															else if (text3.EndsWith(XC.A(20089)))
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
																																text2 = XC.A(20477);
																																flag6 = true;
																															}
																															goto IL_07a7;
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
																													text2 = XC.A(20453);
																												}
																												goto IL_07a7;
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
																										text2 = XC.A(20406);
																										flag4 = true;
																									}
																									goto IL_07a7;
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
																							text2 = XC.A(20344);
																						}
																						goto IL_07a7;
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
																				text2 = XC.A(20303);
																				goto IL_07a7;
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
																		text2 = XC.A(20281);
																		flag3 = true;
																		goto IL_07a7;
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
																text2 = XC.A(20261);
																flag2 = true;
																goto IL_07a7;
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
												}
											}
										}
									}
								}
							}
						}
					}
					text2 = XC.A(20237);
					flag = true;
					goto IL_07a7;
				}
				Forms.WarningMessage(XC.A(20092));
			}
			goto end_IL_0018;
			IL_12ea:
			string text4;
			Selection selection16;
			selection16.TypeText(text4);
			goto IL_12f2;
			IL_0b23:
			if (num == 1)
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
				if (text2.Length > 0 && !text4.Contains(XC.A(20585)))
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
					if (!text4.Contains(XC.A(20600)))
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
						if (flag)
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
							text2 = XC.A(20615);
						}
						else if (flag2)
						{
							text2 = XC.A(20630);
						}
						else if (flag3)
						{
							text2 = XC.A(20641);
						}
						else if (flag4)
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
							text2 = XC.A(20654);
						}
						else if (flag5)
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
							text2 = XC.A(20667);
						}
						else if (flag6)
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
							text2 = XC.A(20678);
						}
					}
				}
			}
			double num2 = result % 1.0;
			bool flag7;
			bool flag8;
			if (num2 == 0.0)
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
				if (flag7)
				{
					string text5 = Conversions.ToString(Operators.ConcatenateObject(XC.A(18458), Numbers.A()));
					if (A)
					{
						text4 = text4 + XC.A(20691) + text.Replace(XC.A(20080), "") + XC.A(20696);
					}
					text4 += text5;
				}
				else if (flag8)
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
					string text6 = Conversions.ToString(Operators.ConcatenateObject(XC.A(18458), B()));
					if (A)
					{
						text4 = text4 + XC.A(20691) + text.Replace(XC.A(20083), "") + XC.A(20696);
					}
					text4 += text6;
				}
				else
				{
					if (A)
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
						text4 = text4 + XC.A(20691) + B(text).Trim() + XC.A(20696);
					}
					text4 += text2;
				}
				selection16.TypeText(text4);
			}
			else
			{
				if (text2.Length != 0)
				{
					if (flag || flag2)
					{
						if (num == 0L)
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
							if (text4.ToLower().StartsWith(XC.A(20865)))
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
								text4 = "";
								goto IL_113d;
							}
						}
						if (A)
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
							text4 = text4 + XC.A(20691) + num + XC.A(20696);
						}
						text4 = text4 + text2 + XC.A(20854);
					}
					else if (flag3)
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
						if (num == 0L)
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
							if (text4.ToLower().StartsWith(XC.A(20865)))
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
								text4 = "";
								goto IL_113d;
							}
						}
						if (A)
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
							text4 = text4 + XC.A(20691) + num + XC.A(20696);
						}
						text4 = text4 + text2 + XC.A(18458);
					}
					else
					{
						text4 += XC.A(18458);
					}
					goto IL_113d;
				}
				string input = Math.Round(num2, 6).ToString();
				string text7 = "";
				input = Regex.Replace(input, XC.A(20699) + Regex.Escape(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator), "");
				switch (input.Length)
				{
				case 1:
					text7 = XC.A(20706);
					break;
				case 2:
					text7 = XC.A(20719);
					break;
				case 3:
					text7 = XC.A(20740);
					break;
				case 4:
					text7 = XC.A(20763);
					break;
				case 5:
					text7 = XC.A(20794);
					break;
				case 6:
					text7 = XC.A(20833);
					break;
				}
				Fields fields = selection16.Fields;
				Range range = selection16.Range;
				Unit = WdFieldType.wdFieldEmpty;
				Count = XC.A(20521) + input + XC.A(20526);
				Extend = true;
				fields.Add(range, ref Unit, ref Count, ref Extend);
				Selection selection17 = selection16;
				Extend = WdUnits.wdWord;
				Count = 1;
				Unit = WdMovementType.wdExtend;
				selection17.MoveLeft(ref Extend, ref Count, ref Unit);
				text4 = text4 + XC.A(20854) + selection16.Text + XC.A(18458) + text7;
				if (A)
				{
					text4 = text4 + XC.A(20691) + text.Replace(XC.A(20080), "").Replace(XC.A(20083), "") + XC.A(20696);
				}
				if (flag7)
				{
					text4 = Conversions.ToString(Operators.ConcatenateObject(text4, Operators.ConcatenateObject(XC.A(18458), Numbers.A())));
				}
				if (flag8)
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
					text4 = Conversions.ToString(Operators.ConcatenateObject(text4, Operators.ConcatenateObject(XC.A(18458), B())));
				}
				selection16.TypeText(text4);
			}
			goto IL_12f2;
			IL_12f2:
			selection16 = null;
			if (A)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, XC.A(20944));
					break;
				}
			}
			else
			{
				clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, XC.A(20490));
			}
			goto end_IL_0018;
			IL_07a7:
			text3 = null;
			flag7 = text.EndsWith(XC.A(20080));
			flag8 = text.EndsWith(XC.A(20083));
			undoRecord = application.UndoRecord;
			undoRecord.StartCustomRecord(XC.A(20490));
			text4 = "";
			if (num > 999999999)
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
				if (num <= 999999999999L)
				{
					Selection selection18 = application.Selection;
					Fields fields2 = selection18.Fields;
					Range range2 = application.Selection.Range;
					Unit = WdFieldType.wdFieldEmpty;
					Count = XC.A(20521) + Strings.Trim(Conversions.ToString(Conversion.Int(((double)num / 1000000000.0).ToString()))) + XC.A(20526);
					Extend = true;
					fields2.Add(range2, ref Unit, ref Count, ref Extend);
					Extend = WdUnits.wdWord;
					Count = 1;
					Unit = WdMovementType.wdExtend;
					selection18.MoveLeft(ref Extend, ref Count, ref Unit);
					text4 = selection18.Text + XC.A(20551);
					_ = null;
					num = Conversions.ToLong(Strings.Right(num.ToString(), 9));
				}
			}
			if (num > 999999)
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
				if (num <= 999999999)
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
					Selection selection19 = application.Selection;
					Fields fields3 = selection19.Fields;
					Range range3 = application.Selection.Range;
					Unit = WdFieldType.wdFieldEmpty;
					Count = XC.A(20521) + Strings.Trim(Conversions.ToString(Conversion.Int(((double)num / 1000000.0).ToString()))) + XC.A(20526);
					Extend = true;
					fields3.Add(range3, ref Unit, ref Count, ref Extend);
					Selection selection20 = selection19;
					Extend = WdUnits.wdWord;
					Count = 1;
					Unit = WdMovementType.wdExtend;
					selection20.MoveLeft(ref Extend, ref Count, ref Unit);
					if (text4.Length > 0)
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
						text4 += XC.A(18458);
					}
					text4 = text4 + selection19.Text + XC.A(20568);
					selection19 = null;
					num = Conversions.ToLong(Strings.Right(num.ToString(), 6));
				}
			}
			if (num <= 999999)
			{
				selection16 = application.Selection;
				if (num == 0L)
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
					if (text4.Length > 0)
					{
						goto IL_0b23;
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
				Fields fields4 = selection16.Fields;
				Range range4 = application.Selection.Range;
				Unit = WdFieldType.wdFieldEmpty;
				Count = XC.A(20521) + num + XC.A(20526);
				Extend = true;
				fields4.Add(range4, ref Unit, ref Count, ref Extend);
				Selection selection21 = selection16;
				Extend = WdUnits.wdWord;
				Count = 1;
				Unit = WdMovementType.wdExtend;
				selection21.MoveLeft(ref Extend, ref Count, ref Unit);
				text4 = ((text4.Length <= 0) ? (text4 + selection16.Text) : (text4 + XC.A(18458) + selection16.Text));
				goto IL_0b23;
			}
			Forms.WarningMessage(XC.A(20991));
			goto end_IL_0018;
			IL_113d:
			Fields fields5 = selection16.Fields;
			Range range5 = selection16.Range;
			Unit = WdFieldType.wdFieldEmpty;
			Count = XC.A(20874) + result + XC.A(20899);
			Extend = true;
			fields5.Add(range5, ref Unit, ref Count, ref Extend);
			Selection selection22 = selection16;
			Extend = WdUnits.wdWord;
			Count = 1;
			Unit = WdMovementType.wdExtend;
			selection22.MoveLeft(ref Extend, ref Count, ref Unit);
			text4 += selection16.Text;
			double num3 = Numbers.A(num2);
			if (!flag)
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
				if (!flag2)
				{
					if (flag3)
					{
						if (A)
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
							text4 = text4 + XC.A(20691) + num3 + XC.A(20696);
						}
						if (num3 == 1.0)
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
							text4 += XC.A(20678);
						}
						else
						{
							text4 += XC.A(20477);
						}
					}
					goto IL_12ea;
				}
			}
			if (A)
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
				text4 = text4 + XC.A(20691) + num3 + XC.A(20696);
			}
			if (num3 == 1.0)
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
				text4 += XC.A(20667);
			}
			else
			{
				text4 += XC.A(20464);
			}
			goto IL_12ea;
			end_IL_0018:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		if (undoRecord != null)
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
			undoRecord.EndCustomRecord();
		}
		application = null;
		undoRecord = null;
	}

	private static double A(double A)
	{
		return Math.Round(A, 2) * 100.0;
	}

	private static string B(string A)
	{
		return A.Replace(XC.A(20171), "").Replace(XC.A(20178), "").Replace(XC.A(20185), "")
			.Replace(XC.A(20192), "")
			.Replace(XC.A(20199), "")
			.Replace(XC.A(20254), "")
			.Replace(XC.A(20274), "")
			.Replace(XC.A(20296), "")
			.Replace(XC.A(20312), "")
			.Replace(XC.A(20319), "")
			.Replace(XC.A(20337), "")
			.Replace(XC.A(20370), "")
			.Replace(XC.A(20392), "")
			.Replace(XC.A(20399), "")
			.Replace(XC.A(20421), "")
			.Replace(XC.A(20446), "")
			.Replace(XC.A(20206), "")
			.Replace(XC.A(20230), "")
			.Replace(XC.A(20223), "")
			.Replace(XC.A(20213), "")
			.Replace(XC.A(20218), "")
			.Replace(XC.A(20441), "")
			.Replace(XC.A(20062), "")
			.Replace(XC.A(20065), "")
			.Replace(XC.A(20068), "")
			.Replace(XC.A(20071), "")
			.Replace(XC.A(20074), "")
			.Replace(XC.A(20077), "")
			.Replace(XC.A(20086), "")
			.Replace(XC.A(20089), "");
	}

	private static object A()
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0005: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_0036: Expected I4, but got Unknown
		Language applicationLanguage = clsEnvironment.ApplicationLanguage;
		switch (applicationLanguage - 1)
		{
		case 0:
			return XC.A(21092);
		case 2:
			return XC.A(21107);
		case 1:
			return XC.A(21128);
		case 3:
			return XC.A(21147);
		case 5:
		case 6:
		case 7:
		case 9:
			return XC.A(21162);
		case 4:
			return XC.A(21177);
		case 8:
			return XC.A(21196);
		default:
			return XC.A(21092);
		}
	}

	private static object B()
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0005: Unknown result type (might be due to invalid IL or missing references)
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_000a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected I4, but got Unknown
		Language applicationLanguage = clsEnvironment.ApplicationLanguage;
		return (applicationLanguage - 1) switch
		{
			0 => XC.A(21217), 
			2 => XC.A(21228), 
			1 => XC.A(21239), 
			3 => XC.A(21248), 
			5 => XC.A(21255), 
			6 => XC.A(21268), 
			7 => XC.A(21279), 
			9 => XC.A(21288), 
			4 => XC.A(21299), 
			8 => XC.A(21310), 
			_ => XC.A(21217), 
		};
	}
}
