using System;
using System.Collections;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using A;
using ExcelAddIn1.ExcelApp;
using ExcelAddIn1.Publishing.Share;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.View;

public sealed class PrintAreas
{
	public static void HidePageBreaks()
	{
		if (!Licensing.AllowAdvancedViewOperation())
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
			try
			{
				try
				{
					enumerator = MH.A.Application.ActiveWindow.SelectedSheets.GetEnumerator();
					while (enumerator.MoveNext())
					{
						object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
						if (objectValue is Worksheet)
						{
							((Worksheet)objectValue).DisplayPageBreaks = false;
						}
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_006e;
						}
						continue;
						end_IL_006e:
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
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, VH.A(175095));
			return;
		}
	}

	public static void SmartPrintArea()
	{
		if (!Licensing.AllowAdvancedViewOperation())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (EditMode.IsEditMode(application))
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
					application = null;
					return;
				}
			}
		}
		string left = default(string);
		string text2 = default(string);
		string text = default(string);
		bool flag = default(bool);
		string left2 = default(string);
		if (application.Selection is Range)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
				{
					Range range = (Range)application.Selection;
					Worksheet worksheet = range.Worksheet;
					try
					{
						string listSeparator = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
						if (Operators.CompareString(worksheet.PageSetup.PrintArea, "", TextCompare: false) == 0)
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
							text = A(range);
							left = text;
						}
						else
						{
							string[] array = Strings.Split(worksheet.PageSetup.PrintArea, listSeparator, -1, CompareMethod.Text);
							string[] array2 = array;
							int num = 0;
							while (true)
							{
								if (num >= array2.Length)
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
									break;
								}
								object obj = array2[num];
								Range arg = ((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(obj), RuntimeHelpers.GetObjectValue(Missing.Value));
								if (application.Intersect(range, arg, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) != null)
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
									text2 = Conversions.ToString(obj);
									break;
								}
								num = checked(num + 1);
							}
							if (Operators.CompareString(text2, "", TextCompare: false) != 0)
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
								if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
									text = worksheet.PageSetup.PrintArea;
									text = Strings.Replace(text, text2, "");
									text = Strings.Replace(text, VH.A(175128), VH.A(2378));
									try
									{
										text = Regex.Replace(text, VH.A(175133), "");
										text = Regex.Replace(text, VH.A(175138), "");
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										ProjectData.ClearProjectError();
									}
									if (Operators.CompareString(text, "", TextCompare: false) == 0)
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
										flag = true;
									}
								}
								else if (MessageBox.Show(VH.A(175143), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
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
									text = range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									left = text;
								}
							}
							else
							{
								wpfPrintArea wpfPrintArea2 = new wpfPrintArea();
								wpfPrintArea2.ShowDialog();
								if (wpfPrintArea2.DialogResult.HasValue && wpfPrintArea2.DialogResult.Value)
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
									if (wpfPrintArea2.SetPrintArea)
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
										text = A(range);
										left = text;
									}
									else
									{
										text = worksheet.PageSetup.PrintArea.Replace(VH.A(77635), VH.A(2378)) + VH.A(2378) + range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
										left = range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
									}
								}
								wpfPrintArea2 = null;
							}
						}
						if (Operators.CompareString(text, "", TextCompare: false) != 0 || flag)
						{
							try
							{
								application.ScreenUpdating = false;
								application.DisplayAlerts = false;
								application.PrintCommunication = false;
								int num2 = Conversions.ToInteger(application.ActiveWindow.Zoom);
								worksheet.PageSetup.PrintArea = text;
								if (!flag)
								{
									string[] array = Strings.Split(text, VH.A(2378), -1, CompareMethod.Text);
									try
									{
										left2 = array[1];
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										ProjectData.ClearProjectError();
									}
									if (Operators.CompareString(left2, "", TextCompare: false) == 0)
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
										PageSetup pageSetup = worksheet.PageSetup;
										pageSetup.Zoom = false;
										pageSetup.FitToPagesTall = 1;
										pageSetup.FitToPagesWide = 1;
										range = ((_Worksheet)worksheet).get_Range((object)array[0], RuntimeHelpers.GetObjectValue(Missing.Value));
										if (Operators.ConditionalCompareObjectGreater(range.Width, range.Height, TextCompare: false))
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
											pageSetup.Orientation = XlPageOrientation.xlLandscape;
										}
										else
										{
											pageSetup.Orientation = XlPageOrientation.xlPortrait;
										}
										pageSetup = null;
									}
								}
								if (Operators.CompareString(left, "", TextCompare: false) != 0)
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
									_ = null;
								}
								Window activeWindow = application.ActiveWindow;
								if (activeWindow.View == XlWindowView.xlNormalView)
								{
									activeWindow.View = XlWindowView.xlPageBreakPreview;
									activeWindow.Zoom = num2;
								}
								activeWindow = null;
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								Forms.ErrorMessage(ex6.Message);
								ProjectData.ClearProjectError();
							}
							finally
							{
								application.DisplayAlerts = true;
								application.ScreenUpdating = true;
								application.PrintCommunication = true;
							}
						}
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						Forms.ErrorMessage(ex8.Message);
						ProjectData.ClearProjectError();
					}
					application = null;
					range = null;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, VH.A(175415));
					return;
				}
				}
			}
		}
		application = null;
	}

	private static string A(Range A)
	{
		if (Operators.ConditionalCompareObjectEqual(A.Cells.CountLarge, 1, TextCompare: false))
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return A.CurrentRegion.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
			}
		}
		return A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	public static void SetPrintAreas()
	{
		if (!Licensing.AllowAdvancedViewOperation())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			bool flag = false;
			if (application.Selection is Range)
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
				Range range = (Range)application.Selection;
				if (Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false))
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
					if (MessageBox.Show(VH.A(175448), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
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
				}
				if (!flag)
				{
					application.ScreenUpdating = false;
					application.EnableEvents = false;
					application.DisplayAlerts = false;
					application.PrintCommunication = false;
					try
					{
						Worksheet worksheet = range.Worksheet;
						string cell = range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						try
						{
							enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
							while (enumerator.MoveNext())
							{
								object objectValue = RuntimeHelpers.GetObjectValue(enumerator.Current);
								if (!(objectValue is Worksheet))
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
								Worksheet worksheet2 = (Worksheet)objectValue;
								worksheet2.Activate();
								Window activeWindow = application.ActiveWindow;
								int num = Conversions.ToInteger(activeWindow.Zoom);
								activeWindow.View = XlWindowView.xlPageBreakPreview;
								activeWindow.Zoom = num;
								_ = null;
								worksheet2.PageSetup.PrintArea = "";
								worksheet2.PageSetup.PrintArea = ((_Worksheet)worksheet2).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
								worksheet2 = null;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_0201;
								}
								continue;
								end_IL_0201:
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
						worksheet.Activate();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					finally
					{
						Worksheet worksheet = null;
					}
					range = null;
					application.DisplayAlerts = true;
					application.ScreenUpdating = true;
					application.EnableEvents = true;
					application.PrintCommunication = true;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)7, VH.A(175588));
				}
			}
			application = null;
			return;
		}
	}

	public static void RemovePrintAreas()
	{
		if (!Licensing.AllowAdvancedViewOperation())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			application.ScreenUpdating = false;
			application.EnableEvents = false;
			try
			{
				object objectValue = RuntimeHelpers.GetObjectValue(application.ActiveSheet);
				try
				{
					enumerator = application.ActiveWindow.SelectedSheets.GetEnumerator();
					while (enumerator.MoveNext())
					{
						object objectValue2 = RuntimeHelpers.GetObjectValue(enumerator.Current);
						if (!(objectValue2 is Worksheet))
						{
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
						Base.K((Worksheet)objectValue2);
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0098;
						}
						continue;
						end_IL_0098:
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
				NewLateBinding.LateCall(objectValue, null, VH.A(39985), new object[0], null, null, null, IgnoreReturn: true);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			finally
			{
				object objectValue = null;
			}
			application.ScreenUpdating = true;
			application.EnableEvents = true;
			application = null;
			return;
		}
	}
}
