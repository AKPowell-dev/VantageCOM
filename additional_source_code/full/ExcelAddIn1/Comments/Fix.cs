using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using A;
using ExcelAddIn1.Sheets;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Comments;

public sealed class Fix
{
	public static void NotesInSelection()
	{
		Application application = MH.A.Application;
		Range range = null;
		int num = 0;
		int num2 = 0;
		bool blnResize = default(bool);
		bool blnAuto = default(bool);
		try
		{
			XmlNode xmlNode = KH.A.SettingsXml.DocumentElement.SelectSingleNode(VH.A(117297));
			blnResize = Conversions.ToBoolean(xmlNode.Attributes[VH.A(117320)].Value);
			blnAuto = Conversions.ToBoolean(xmlNode.Attributes[VH.A(117333)].Value);
			_ = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		checked
		{
			try
			{
				Window activeWindow = application.ActiveWindow;
				if (activeWindow.SelectedSheets.Count > 1)
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
					{
						IEnumerator enumerator = activeWindow.SelectedSheets.GetEnumerator();
						try
						{
							IEnumerator enumerator2 = default(IEnumerator);
							while (enumerator.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
								if (!(objectValue is Worksheet))
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
								Worksheet worksheet = (Worksheet)objectValue;
								try
								{
									try
									{
										enumerator2 = worksheet.Comments.GetEnumerator();
										while (enumerator2.MoveNext())
										{
											Note((Comment)enumerator2.Current, blnResize, blnAuto);
										}
										while (true)
										{
											switch (2)
											{
											case 0:
												break;
											default:
												goto end_IL_014c;
											}
											continue;
											end_IL_014c:
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
									num += worksheet.Comments.Count;
									num2++;
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									clsReporting.LogException(ex4);
									ProjectData.ClearProjectError();
								}
								worksheet = null;
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_01b8;
								}
								continue;
								end_IL_01b8:
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
					Forms.InfoMessage(VH.A(142561) + num + VH.A(142154) + num2 + VH.A(142175));
				}
				else if (application.Selection is Range)
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
					Range range2 = (Range)application.Selection;
					ExcelAddIn1.Sheets.Protection.Unprotect(range2.Worksheet);
					try
					{
						Range range3;
						if (!Operators.ConditionalCompareObjectEqual(range2.Cells.CountLarge, 1, TextCompare: false))
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
							range3 = range2.SpecialCells(XlCellType.xlCellTypeComments, RuntimeHelpers.GetObjectValue(Missing.Value));
						}
						else
						{
							range3 = range2;
						}
						range = range3;
					}
					catch (Exception ex5)
					{
						ProjectData.SetProjectError(ex5);
						Exception ex6 = ex5;
						ProjectData.ClearProjectError();
					}
					if (range != null)
					{
						try
						{
							IEnumerator enumerator3 = default(IEnumerator);
							try
							{
								enumerator3 = range.GetEnumerator();
								while (enumerator3.MoveNext())
								{
									Range range4 = (Range)enumerator3.Current;
									if (range4.Comment == null)
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
									Note(range4.Comment, blnResize, blnAuto);
								}
							}
							finally
							{
								if (enumerator3 is IDisposable)
								{
									while (true)
									{
										switch (3)
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
						catch (Exception ex7)
						{
							ProjectData.SetProjectError(ex7);
							Exception ex8 = ex7;
							clsReporting.LogException(ex8);
							ProjectData.ClearProjectError();
						}
						range = null;
					}
					range2 = null;
				}
				activeWindow = null;
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				clsReporting.LogException(ex10);
				ProjectData.ClearProjectError();
			}
			application = null;
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)1, VH.A(142574));
		}
	}

	public static void Note(Comment c, bool blnResize, bool blnAuto)
	{
		if (blnResize)
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
			if (blnAuto)
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
				C(c);
			}
			else
			{
				B(c);
			}
		}
		A(c);
	}

	private static void A(Comment A)
	{
		Comment comment = A;
		comment.Shape.Top = Conversions.ToSingle(Operators.SubtractObject(NewLateBinding.LateGet(comment.Parent, null, VH.A(57409), new object[0], null, null, null), 7));
		comment.Shape.Left = Conversions.ToSingle(Operators.AddObject(NewLateBinding.LateGet(NewLateBinding.LateGet(comment.Parent, null, VH.A(60565), new object[2] { 0, 1 }, null, null, null), null, VH.A(56582), new object[0], null, null, null), 11));
		comment = null;
	}

	private static void B(Comment A)
	{
		A.Shape.Width = 108f;
		A.Shape.Height = 59.04f;
	}

	private static void C(Comment A)
	{
		int num = 170;
		string expression = A.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		string[] array = Strings.Split(expression, VH.A(41382));
		int num2 = 0;
		int num3 = Information.UBound(array);
		checked
		{
			for (int i = 0; i <= num3; i++)
			{
				if (Strings.Len(array[i]) <= num2)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				num2 = Strings.Len(array[i]);
			}
			A.Shape.TextFrame.AutoSize = true;
			int num4 = (int)Math.Round(Conversion.Fix((float)num / A.Shape.Width * (float)num2) - 1f);
			int num5 = 0;
			string[] array2 = Strings.Split(expression, VH.A(41385));
			int num6 = Information.UBound(array2);
			int num7 = default(int);
			for (int i = 0; i <= num6; i++)
			{
				if (Strings.Len(array2[i]) > num7)
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
					num7 = Strings.Len(array2[i]);
				}
				if (Strings.Len(array2[i]) > num4)
				{
					num5++;
				}
			}
			if (num7 > num4)
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
				num4 = num7 + 1;
			}
			int num8 = 0;
			int num9 = Information.UBound(array);
			for (int i = 0; i <= num9; i++)
			{
				num8++;
				string text;
				do
				{
					text = Strings.Left(array[i], num4);
					if (Strings.Len(text) == num4)
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
						text = Strings.Left(text, Strings.InStrRev(text, VH.A(41385)));
					}
					array[i] = Strings.Replace(array[i], text, "");
					if (Strings.Len(array[i]) > 0)
					{
						num8++;
					}
				}
				while (!((Strings.Len(array[i]) == 0) | (text.Length == 0)));
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_020b;
					}
					continue;
					end_IL_020b:
					break;
				}
			}
			num8 += num5;
			Shape shape = A.Shape;
			TextFrame textFrame = shape.TextFrame;
			textFrame.AutoMargins = false;
			textFrame.MarginTop = 0f;
			textFrame.MarginBottom = 0f;
			_ = null;
			float num10 = A.Shape.Height / (float)array.Length;
			shape.TextFrame.AutoMargins = true;
			int num11 = (int)Math.Round(num10 * (float)num8);
			shape.Width = num;
			shape.Height = num11;
			_ = null;
		}
	}
}
