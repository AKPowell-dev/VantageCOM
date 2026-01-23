using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using ExcelAddIn1.Formulas;
using ExcelAddIn1.Library2;
using ExcelAddIn1.Library2.UI;
using MacabacusMacros;
using MacabacusMacros.Libraries.Pane.UI;
using MacabacusMacros.UI;
using MacabacusMacros.UI.FormsExtensions;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal sealed class VE
{
	private struct RE
	{
		internal int A;

		internal int B;

		internal int C;

		internal int D;

		internal RE(Range A)
		{
			this = default(RE);
			checked
			{
				try
				{
					Range rows = A.Rows;
					this.A = A.Row;
					B = this.A - 1 + rows.Count;
					Range columns = A.Columns;
					C = A.Column;
					D = C - 1 + columns.Count;
				}
				finally
				{
					Range columns = null;
					Range rows = null;
				}
			}
		}
	}

	[CompilerGenerated]
	internal sealed class SE
	{
		public Range A;

		public Range B;

		public UE A;

		public SE(SE A)
		{
			if (A == null)
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
				this.A = A.A;
				B = A.B;
				return;
			}
		}

		[SpecialName]
		internal Range A()
		{
			return this.A.A.A.m_A.Intersect(this.A, B, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
	}

	[CompilerGenerated]
	internal sealed class TE
	{
		public List<Range> A;

		public VE A;

		public TE(TE A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class UE
	{
		public int A;

		public TE A;

		public Func<Range> A;

		public UE(UE A)
		{
			if (A == null)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal Range A()
		{
			return this.A.A[checked(this.A - 1)];
		}
	}

	private Microsoft.Office.Interop.Excel.Application m_A;

	private readonly Func<ContentItem, Workbook> m_A;

	internal VE(Microsoft.Office.Interop.Excel.Application A, Func<ContentItem, Workbook> B)
	{
		this.m_A = A;
		this.m_A = B;
	}

	internal void A()
	{
		this.m_A = null;
	}

	internal bool A(IList<ContentItem> A)
	{
		TE tE = new TE(tE);
		tE.A = this;
		List<Workbook> list = new List<Workbook>();
		Range range = null;
		Workbook workbook = null;
		WE wE = null;
		try
		{
			workbook = this.m_A.ActiveWorkbook;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		this.m_A.ScreenUpdating = false;
		List<Range> list2 = new List<Range>();
		List<Range> list3 = new List<Range>();
		tE.A = new List<Range>();
		Range range2 = null;
		try
		{
			if (workbook == null)
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
					break;
				}
			}
			else
			{
				wE = new WE(this.m_A);
				Range range3 = this.m_A.ActiveCell;
				Range c = range3;
				Worksheet worksheet = range3.Worksheet;
				string text = null;
				try
				{
					UE uE = new UE(uE);
					uE.A = tE;
					if (workbook.Path.Length == 0)
					{
						workbook.Saved = false;
					}
					bool? flag = null;
					bool? flag2 = null;
					bool? flag3 = null;
					bool? flag4 = ((A.Count > 1) ? VE.A() : new bool?(false));
					if (!flag4.HasValue)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_010c;
							}
							continue;
							end_IL_010c:
							break;
						}
					}
					else
					{
						bool value = flag4.Value;
						int num = 0;
						bool flag5 = A.Count > 1;
						checked
						{
							IEnumerator<ContentItem> enumerator = default(IEnumerator<ContentItem>);
							try
							{
								enumerator = A.GetEnumerator();
								int val = default(int);
								while (enumerator.MoveNext())
								{
									ContentItem current = enumerator.Current;
									Range range6;
									try
									{
										if (num > 0)
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
											range3 = JH.A(range3, !value, num + 1);
										}
										Workbook workbook2 = this.m_A(current);
										if (!list.Contains(workbook2))
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
											list.Add(workbook2);
										}
										((Worksheet)workbook2.Sheets[((TableItem)(object)current).SheetIndex]).Activate();
										Range range4 = (Range)this.m_A.Selection;
										list2.Add(range4);
										int C;
										int D;
										Range range5 = JH.A(range3, range4, out C, out D);
										list3.Add(range5);
										range2 = ((range2 != null) ? this.m_A.Union(range2, range5, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) : range5);
										int num2;
										if (!value)
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
											num2 = D;
										}
										else
										{
											num2 = C;
										}
										int num3 = num2;
										if (num > 0)
										{
											int num4 = Math.Min(num3, val);
											try
											{
												if (value)
												{
													while (true)
													{
														switch (1)
														{
														case 0:
															continue;
														}
														range6 = ((Range)worksheet.Cells[range5.Row, range5.Column - 1]).get_Resize((object)num4, (object)1);
														break;
													}
												}
												else
												{
													range6 = ((Range)worksheet.Cells[range5.Row - 1, range5.Column]).get_Resize((object)1, (object)num4);
												}
											}
											finally
											{
											}
											uE.A.A.Add(range6);
											range2 = this.m_A.Union(range2, range6, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										}
										val = num3;
										int num5;
										if (!value)
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
											num5 = C;
										}
										else
										{
											num5 = D;
										}
										num = num5;
									}
									finally
									{
										range6 = null;
										Range range5 = null;
										Range range4 = null;
										Workbook workbook2 = null;
									}
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_05c0;
									}
									continue;
									end_IL_05c0:
									break;
								}
							}
							finally
							{
								if (enumerator != null)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											continue;
										}
										enumerator.Dispose();
										break;
									}
								}
							}
						}
						if (this.A(range2, worksheet, c))
						{
							uE.A = -1;
							IEnumerator<ContentItem> enumerator2 = default(IEnumerator<ContentItem>);
							try
							{
								enumerator2 = A.GetEnumerator();
								SE sE = default(SE);
								while (enumerator2.MoveNext())
								{
									ContentItem current2 = enumerator2.Current;
									sE = new SE(sE);
									sE.A = uE;
									sE.A.A = checked(sE.A.A + 1);
									text = string.Format(VH.A(86485), current2.GetItemTitle());
									try
									{
										sE.A = list3[sE.A.A];
										Range range7 = list2[sE.A.A];
										range3 = (Range)sE.A.Cells[1, 1];
										int count = range7.Rows.Count;
										int count2 = range7.Columns.Count;
										sE.B = worksheet.UsedRange;
										RE c2 = new RE(sE.B);
										List<object> D2 = null;
										List<object> E = null;
										bool flag6 = this.A(range7, range3, c2, out D2, out E);
										bool? C2 = null;
										int num6;
										if (!object.Equals(flag, true))
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
											if (!this.A(worksheet, sE.B, ref C2))
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
												if (this.m_A.Intersect(sE.A, sE.B.EntireRow, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
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
													num6 = ((!flag6) ? 1 : 0);
													goto IL_08d4;
												}
											}
										}
										num6 = 1;
										goto IL_08d4;
										IL_0a91:
										int num7;
										bool flag7 = (byte)num7 != 0;
										bool flag8;
										if (!flag8)
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
											if (!flag.HasValue)
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
												int value2;
												if (!flag5)
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
													value2 = ((!UIFormsExtensions.AskYesNo((System.Windows.Window)null, VH.A(86508), false, true)) ? 1 : 0);
												}
												else
												{
													value2 = 1;
												}
												flag = (byte)value2 != 0;
											}
											if (flag == true)
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
												flag8 = true;
											}
										}
										if (!flag7)
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
											if (!flag2.HasValue)
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
												int value3;
												if (!flag5)
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
													value3 = ((!UIFormsExtensions.AskYesNo((System.Windows.Window)null, VH.A(86575), false, true)) ? 1 : 0);
												}
												else
												{
													value3 = 1;
												}
												flag2 = (byte)value3 != 0;
											}
											if (flag2 == true)
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
												flag7 = true;
											}
										}
										if (!flag3.HasValue)
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
											if (Helpers.A(range7))
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
												flag3 = true;
											}
										}
										VE.A(sE.A);
										if (sE.A.A > 0)
										{
											Func<Range> a;
											if (sE.A.A != null)
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
												a = sE.A.A;
											}
											else
											{
												a = (sE.A.A = [SpecialName] () => sE.A.A.A[checked(sE.A.A - 1)]);
											}
											VE.A(a);
										}
										bool flag9 = false;
										if (object.Equals(flag3, true))
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
											wE.A(range7, range3, sE.A, workbook);
										}
										else
										{
											range7.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
											flag9 = true;
											range3.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										}
										if (flag7)
										{
											if (!flag9)
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
												range7.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
											}
											range3.PasteSpecial(XlPasteType.xlPasteColumnWidths, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										}
										Range a2 = range7;
										Range b = range3;
										List<object> c3;
										if (!flag8)
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
											c3 = E;
										}
										else
										{
											c3 = D2;
										}
										this.A(a2, b, c3);
										Range range8;
										if (range != null)
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
											range8 = this.m_A.Union(range, sE.A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										}
										else
										{
											range8 = sE.A;
										}
										range = range8;
										int num8;
										if (!value)
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
											num8 = count;
										}
										else
										{
											num8 = count2;
										}
										num = num8;
										goto end_IL_0668;
										IL_08d4:
										flag8 = (byte)num6 != 0;
										if (!object.Equals(flag2, true))
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
											if (!this.A(worksheet, sE.B, ref C2))
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
												if (this.m_A.Intersect(sE.A, sE.B.EntireColumn, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
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
													num7 = ((!B(range7, range3, c2)) ? 1 : 0);
													goto IL_0a91;
												}
											}
										}
										num7 = 1;
										goto IL_0a91;
										end_IL_0668:;
									}
									finally
									{
										range3 = null;
										sE.B = null;
										sE.A = null;
										Range range7 = null;
									}
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_0e90;
									}
									continue;
									end_IL_0e90:
									break;
								}
							}
							finally
							{
								if (enumerator2 != null)
								{
									while (true)
									{
										switch (1)
										{
										case 0:
											continue;
										}
										enumerator2.Dispose();
										break;
									}
								}
							}
							text = null;
							return true;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_05f1;
							}
							continue;
							end_IL_05f1:
							break;
						}
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					string format = VH.A(86646);
					object[] array = new object[4];
					string text2 = text;
					if (text2 == null)
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
						text2 = VH.A(86739);
					}
					array[0] = text2;
					array[1] = VH.A(7803);
					array[2] = VH.A(7803);
					array[3] = ex4.Message;
					Forms.ErrorMessage(string.Format(format, array));
					clsReporting.LogException(ex4);
					ProjectData.ClearProjectError();
				}
			}
		}
		finally
		{
			if (wE != null)
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
				wE.A();
			}
			wE = null;
			this.A(list);
			if (list.Count > 1)
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
				VE.A(range);
			}
			this.m_A.ScreenUpdating = true;
			range2 = null;
			tE.A.Clear();
			tE.A = null;
			list3.Clear();
			list3 = null;
			list2.Clear();
			list2 = null;
			range = null;
			workbook = null;
			list.Clear();
			list = null;
			Worksheet worksheet = null;
			Range range3 = null;
			Range c = null;
		}
		return false;
	}

	private bool A(Range A, Worksheet B, Range C)
	{
		if (A == null)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		bool flag = false;
		bool? flag2 = null;
		bool flag3 = false;
		try
		{
			Range usedRange = B.UsedRange;
			A = this.m_A.Intersect(A, usedRange, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			if (A != null)
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
				if (this.m_A.WorksheetFunction.CountA(A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != 0.0)
				{
					B.Activate();
					VE.A(A);
					flag = true;
					flag2 = this.m_A.ScreenUpdating;
					this.m_A.ScreenUpdating = true;
					flag3 = UIFormsExtensions.AskOkCancel((System.Windows.Window)null, VH.A(86764));
					return flag3;
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
			return true;
		}
		finally
		{
			if (flag && !flag3)
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
				VE.A(C);
			}
			if (flag2.HasValue)
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
				this.m_A.ScreenUpdating = flag2.Value;
			}
			A = null;
			Range usedRange = null;
		}
	}

	private static void A(Range A)
	{
		try
		{
			if (A == null)
			{
				return;
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
				A.Select();
				return;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private bool A(Worksheet A, Range B, ref bool? C)
	{
		if (C.HasValue)
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
					return C.Value;
				}
			}
		}
		C = false;
		if (object.Equals(RuntimeHelpers.GetObjectValue(A.Columns.ColumnWidth), A.StandardWidth) && object.Equals(RuntimeHelpers.GetObjectValue(A.Rows.RowHeight), A.StandardHeight))
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
			C = true;
		}
		else if (this.m_A.WorksheetFunction.CountA(B, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) == 0.0)
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
			C = true;
		}
		return C.Value;
	}

	private static bool? A()
	{
		wpfInsertTables wpfInsertTables = new wpfInsertTables();
		bool? result;
		try
		{
			if (wpfInsertTables.ShowDialog() == true)
			{
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
					result = object.Equals(wpfInsertTables.optHorz.IsChecked, true);
					return result;
				}
			}
			result = null;
		}
		finally
		{
			try
			{
				wpfInsertTables.Close();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	private void A(List<Workbook> A)
	{
		if (A.Count < 1)
		{
			return;
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
			try
			{
				Mouse.OverrideCursor = Cursors.AppStarting;
				this.m_A.DisplayAlerts = false;
				try
				{
					this.m_A.CutCopyMode = (XlCutCopyMode)0;
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					ProjectData.ClearProjectError();
				}
				using List<Workbook>.Enumerator enumerator = A.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Workbook current = enumerator.Current;
					try
					{
						current.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					finally
					{
						current = null;
					}
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
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				clsReporting.LogException(new Exception(string.Format(VH.A(86904), ex4.Message), ex4));
				ProjectData.ClearProjectError();
				return;
			}
			finally
			{
				this.m_A.DisplayAlerts = true;
				Mouse.OverrideCursor = null;
			}
		}
	}

	private static void A(Func<Range> A)
	{
		Range range = A();
		try
		{
			if (range != null)
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
				range.UnMerge();
			}
			range?.Clear();
		}
		finally
		{
			range = null;
		}
	}

	private bool A(Range A, Range B, RE C, out List<object> D, out List<object> E)
	{
		D = new List<object>();
		E = new List<object>();
		Range range = JH.A(A, this.m_A);
		if (range == null)
		{
			return false;
		}
		bool flag = false;
		checked
		{
			try
			{
				int num = B.Row;
				int num2 = range.Row - A.Row;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = range.Rows.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Range range2 = (Range)enumerator.Current;
						Range range3 = JH.A(B, B: true, num2);
						try
						{
							D.Add(RuntimeHelpers.GetObjectValue(range2.RowHeight));
							if (num < C.A || num > C.B)
							{
								E.Add(DBNull.Value);
							}
							else
							{
								E.Add(RuntimeHelpers.GetObjectValue(range3.RowHeight));
								int num3;
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
									if (1 == 0)
									{
										/*OpCode not supported: LdMemberToken*/;
									}
									num3 = ((!object.Equals(RuntimeHelpers.GetObjectValue(range3.RowHeight), RuntimeHelpers.GetObjectValue(range2.RowHeight))) ? 1 : 0);
								}
								else
								{
									num3 = 1;
								}
								flag = unchecked((byte)num3) != 0;
							}
						}
						finally
						{
							range3 = null;
						}
						num2++;
						num++;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_011c;
						}
						continue;
						end_IL_011c:
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
			}
			finally
			{
				Range range3 = null;
				range = null;
			}
			return flag;
		}
	}

	private bool B(Range A, Range B, RE C)
	{
		Range range = JH.A(A, this.m_A);
		if (range == null)
		{
			return false;
		}
		checked
		{
			try
			{
				int num = B.Column;
				int num2 = range.Column - A.Column;
				IEnumerator enumerator = range.Columns.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						Range range2 = (Range)enumerator.Current;
						try
						{
							if (num >= C.C)
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
									if (num > C.D || object.Equals(RuntimeHelpers.GetObjectValue(JH.A(B, B: false, num2).ColumnWidth), RuntimeHelpers.GetObjectValue(range2.ColumnWidth)))
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
										return true;
									}
								}
							}
						}
						finally
						{
						}
						num2++;
						num++;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_00d8;
						}
						continue;
						end_IL_00d8:
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
			finally
			{
				range = null;
			}
			return false;
		}
	}

	private void A(Range A, Range B, List<object> C)
	{
		if (C == null)
		{
			return;
		}
		checked
		{
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
				Range range = JH.A(A, this.m_A);
				if (range == null)
				{
					return;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					try
					{
						int num = range.Row - A.Row;
						int num2 = 0;
						try
						{
							enumerator = range.Rows.GetEnumerator();
							while (enumerator.MoveNext())
							{
								_ = (Range)enumerator.Current;
								try
								{
									object objectValue = RuntimeHelpers.GetObjectValue(C[num2]);
									if (object.Equals(RuntimeHelpers.GetObjectValue(objectValue), DBNull.Value))
									{
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_0093;
											}
											continue;
											end_IL_0093:
											break;
										}
									}
									else
									{
										Range range2 = JH.A(B, B: true, num);
										if (!object.Equals(RuntimeHelpers.GetObjectValue(range2.RowHeight), RuntimeHelpers.GetObjectValue(objectValue)))
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												range2.RowHeight = RuntimeHelpers.GetObjectValue(objectValue);
												break;
											}
										}
									}
								}
								finally
								{
									Range range2 = null;
								}
								num++;
								num2++;
							}
							return;
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
					finally
					{
						Range range2 = null;
						range = null;
					}
				}
			}
		}
	}
}
