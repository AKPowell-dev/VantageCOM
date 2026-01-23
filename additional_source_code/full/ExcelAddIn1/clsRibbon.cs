using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Xml;
using A;
using ExcelAddIn1.Format;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

public sealed class clsRibbon
{
	private static bool A()
	{
		return Conversions.ToBoolean(KH.A.SettingsXml.DocumentElement.SelectSingleNode(VH.A(186629)).InnerText);
	}

	public static string MacabacusTabKeyTipExcel()
	{
		return clsRibbon.MacabacusTabKeyTip(VH.A(169659));
	}

	public static void DynamicAutoColorToggle()
	{
		XmlDocument settingsXml = KH.A.SettingsXml;
		XmlNode xmlNode = settingsXml.DocumentElement.SelectSingleNode(VH.A(186672));
		xmlNode.InnerText = Conversions.ToString(!Conversions.ToBoolean(xmlNode.InnerText));
		_ = null;
		KH.A.SaveSettings(settingsXml);
		settingsXml = null;
		KH.A.AutoColorOnEntry = !KH.A.AutoColorOnEntry;
	}

	public static bool GetMaximizedState()
	{
		return !MH.A.Application.DisplayStatusBar;
	}

	public static int GetItemCount(IRibbonControl control)
	{
		int count = default(int);
		try
		{
			count = A(control).Value.Count;
			return count;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return count;
	}

	public static string GetCustomLabel(IRibbonControl control)
	{
		int num = Conversions.ToInteger(control.Tag);
		string result = default(string);
		try
		{
			result = VH.A(186705) + checked(num + 1) + VH.A(41385) + A(control).Key;
			return result;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static string GetCycleScreentip(IRibbonControl control)
	{
		string key = default(string);
		try
		{
			key = A(control).Key;
			return key;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return key;
	}

	public static bool CustomMenuVisible(IRibbonControl control)
	{
		bool result = default(bool);
		try
		{
			if (A(control).Value.Count > 0)
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
						result = !A();
						return result;
					}
				}
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

	public static bool CustomGalleryVisible(IRibbonControl control)
	{
		bool result = default(bool);
		try
		{
			if (GetItemCount(control) > 0)
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
						result = A();
						return result;
					}
				}
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

	public static string GetStyleScreentip(IRibbonControl control, int j)
	{
		string innerText = default(string);
		try
		{
			innerText = A(control).Value[j].SelectSingleNode(VH.A(19019)).InnerText;
			return innerText;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return innerText;
	}

	private static KeyValuePair<string, List<XmlNode>> A(IRibbonControl A)
	{
		return KH.A.CustomCycles.ElementAt(Conversions.ToInteger(A.Tag));
	}

	public static Bitmap GetItemImage(IRibbonControl control, int i)
	{
		Bitmap result = null;
		Worksheet worksheet = null;
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Range range;
		XmlNode xmlNode;
		if (!Workbooks.IsShared(application.ActiveWorkbook, true, (System.Windows.Window)null))
		{
			try
			{
				application.EnableEvents = false;
				application.ScreenUpdating = false;
				if (application.ActiveWindow.SelectedSheets.Count > 1)
				{
					NewLateBinding.LateCall(application.ActiveSheet, null, VH.A(51162), new object[0], null, null, null, IgnoreReturn: true);
				}
				worksheet = (Worksheet)application.ActiveWorkbook.Worksheets.Add(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				application.ActiveWindow.DisplayGridlines = false;
				worksheet.Visible = XlSheetVisibility.xlSheetHidden;
				range = ((_Worksheet)worksheet).get_Range((object)VH.A(76945), RuntimeHelpers.GetObjectValue(Missing.Value));
				xmlNode = A(control).Value[i];
				range.Value2 = xmlNode.SelectSingleNode(VH.A(19019)).InnerText;
				ExcelAddIn1.Format.Styles.ApplyStyle(range, xmlNode);
				range.VerticalAlignment = XlVAlign.xlVAlignCenter;
				range.ColumnWidth = 18;
				range.RowHeight = 18;
				range = ((_Application)application).get_Range((object)range.get_Offset((object)(-1), (object)(-1)), (object)range.get_Offset((object)1, (object)1));
				A((Range)range.Cells[1, 1]);
				A((Range)range.Cells[3, 3]);
				application.ScreenUpdating = true;
				range.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);
				if (System.Windows.Forms.Clipboard.ContainsImage())
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
					result = new Bitmap(System.Windows.Forms.Clipboard.GetImage());
				}
				application.ScreenUpdating = false;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			if (worksheet != null)
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
				application.DisplayAlerts = false;
				try
				{
					worksheet.Delete();
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				application.DisplayAlerts = true;
			}
			application.EnableEvents = true;
			application.ScreenUpdating = true;
		}
		else
		{
			KH.A.Invalidate();
		}
		application = null;
		worksheet = null;
		range = null;
		xmlNode = null;
		return result;
	}

	private static void A(Range A)
	{
		A.RowHeight = 0.8;
		A.ColumnWidth = 0.1;
		_ = null;
	}

	public static void ApplyGalleryStyle(IRibbonControl control, int i)
	{
		try
		{
			ExcelAddIn1.Format.Styles.ApplyStyle(KH.A.CustomCycles.ElementAt(Conversions.ToInteger(control.Tag)).Value[i]);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	public static string MenuCustom(IRibbonControl control)
	{
		int num = 1;
		int num2 = Conversions.ToInteger(control.Tag);
		StringBuilder stringBuilder = A();
		KeyValuePair<string, List<XmlNode>> keyValuePair = KH.A.CustomCycles.ElementAt(num2);
		string key = keyValuePair.Key;
		List<XmlNode> value = keyValuePair.Value;
		keyValuePair = default(KeyValuePair<string, List<XmlNode>>);
		checked
		{
			stringBuilder.Append(VH.A(186708) + (num2 + 1) + VH.A(186753) + num2 + VH.A(186866) + key + VH.A(186893));
			stringBuilder.Append(VH.A(186954) + num2 + VH.A(187019));
			foreach (XmlNode item in value)
			{
				key = item.SelectSingleNode(VH.A(19019)).InnerText;
				string text = VH.A(187028) + num + VH.A(41385) + key;
				string text2 = num2 + VH.A(2378) + (num - 1);
				stringBuilder.Append(VH.A(187039) + num2 + VH.A(187074) + num + VH.A(187085) + text + VH.A(187104) + text2 + VH.A(187157) + key + VH.A(187019));
				num++;
			}
			stringBuilder.Append(VH.A(187206));
			return stringBuilder.ToString();
		}
	}

	private static void A(ref StringBuilder A, NumberFormatCycle B, string C)
	{
		int num = 1;
		checked
		{
			foreach (NumberFormatCycle.NumberFormat item in B.Items)
			{
				string text = item.Name.Replace(VH.A(186705), VH.A(187221));
				string text2 = VH.A(187028) + num + VH.A(41385) + text;
				A.Append(string.Format(C, num.ToString(), text2, (num - 1).ToString(), text));
				num++;
			}
			A.Append(VH.A(187206));
		}
	}

	public static string MenuGeneral()
	{
		StringBuilder A = clsRibbon.A();
		clsRibbon.A(ref A, KH.A.CycleNumber, VH.A(187242));
		return A.ToString();
	}

	public static string MenuCurrency()
	{
		StringBuilder A = clsRibbon.A();
		clsRibbon.A(ref A, KH.A.CycleCurrency, VH.A(187426));
		return A.ToString();
	}

	public static string MenuPercent()
	{
		StringBuilder A = clsRibbon.A();
		clsRibbon.A(ref A, KH.A.CyclePercent, VH.A(187612));
		return A.ToString();
	}

	public static string MenuMultiple()
	{
		StringBuilder A = clsRibbon.A();
		clsRibbon.A(ref A, KH.A.CycleMultiple, VH.A(187784));
		return A.ToString();
	}

	public static string MenuDate()
	{
		StringBuilder A = clsRibbon.A();
		clsRibbon.A(ref A, KH.A.CycleDate, VH.A(187960));
		return A.ToString();
	}

	public static string MenuBinary()
	{
		StringBuilder A = clsRibbon.A();
		clsRibbon.A(ref A, KH.A.CycleBinary, VH.A(188120));
		return A.ToString();
	}

	public static string MenuRatio()
	{
		StringBuilder A = clsRibbon.A();
		clsRibbon.A(ref A, KH.A.CycleRatio, VH.A(188288));
		return A.ToString();
	}

	public static string MenuFontStyle()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		using (List<string>.Enumerator enumerator = KH.A.FontStyleCycle.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				string current = enumerator.Current;
				string text = VH.A(187028) + num + VH.A(41385) + current;
				stringBuilder.Append(VH.A(188452) + num + VH.A(187085) + text + VH.A(188495) + current + VH.A(186866) + current + VH.A(187019));
				num = checked(num + 1);
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
				break;
			}
		}
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	public static string MenuFontSize()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		using (List<float>.Enumerator enumerator = KH.A.FontSizeCycle.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				string text = enumerator.Current.ToString();
				string text2 = VH.A(187028) + num + VH.A(188556) + text;
				stringBuilder.Append(VH.A(188579) + num + VH.A(187085) + text2 + VH.A(188620) + text + VH.A(188679) + text + VH.A(187019));
				num = checked(num + 1);
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
				break;
			}
		}
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	public static string MenuHeight()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		using (List<float>.Enumerator enumerator = KH.A.RowHeightCycle.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				string text = enumerator.Current.ToString();
				string text2 = VH.A(187028) + num + VH.A(188726) + text;
				stringBuilder.Append(VH.A(188751) + num + VH.A(187085) + text2 + VH.A(188794) + text + VH.A(187157) + text + VH.A(187019));
				num = checked(num + 1);
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
				break;
			}
		}
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	public static string MenuWidth()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		using (List<float>.Enumerator enumerator = KH.A.ColumnWidthCycle.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				string text = enumerator.Current.ToString();
				string text2 = VH.A(187028) + num + VH.A(188855) + text;
				stringBuilder.Append(VH.A(188884) + num + VH.A(187085) + text2 + VH.A(188925) + text + VH.A(188990) + text + VH.A(187019));
				num = checked(num + 1);
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
				break;
			}
		}
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	public static string BorderColorMenu()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		using (List<ColorCycle.Color>.Enumerator enumerator = KH.A.BorderColorCycle.Colors.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				ColorCycle.Color current = enumerator.Current;
				stringBuilder.Append(VH.A(189043) + Conversions.ToString(num) + VH.A(187085) + A(num, current.RGB) + VH.A(189090) + current.RGB + VH.A(186866) + A(current.RGB) + VH.A(189155));
				num = checked(num + 1);
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
				break;
			}
		}
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	public static string ChartColorMenu()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		foreach (ColorCycle.Color color in KH.A.ChartColorCycle.Colors)
		{
			stringBuilder.Append(VH.A(189222) + Conversions.ToString(num) + VH.A(187085) + A(num, color.RGB) + VH.A(189267) + color.RGB + VH.A(186866) + A(color.RGB) + VH.A(189155));
			num = checked(num + 1);
		}
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	public static string FontColorMenu()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		using (List<ColorCycle.Color>.Enumerator enumerator = KH.A.FontColorCycle.Colors.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				ColorCycle.Color current = enumerator.Current;
				stringBuilder.Append(VH.A(189330) + Conversions.ToString(num) + VH.A(187085) + A(num, current.RGB) + VH.A(189373) + current.RGB + VH.A(186866) + A(current.RGB) + VH.A(189155));
				num = checked(num + 1);
			}
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
				break;
			}
		}
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	public static string FillColorMenu()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		using (List<ColorCycle.Color>.Enumerator enumerator = KH.A.FillColorCycle.Colors.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				ColorCycle.Color current = enumerator.Current;
				if (current.RGB.Length > 0)
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
					stringBuilder.Append(VH.A(189434) + Conversions.ToString(num) + VH.A(187085) + A(num, current.RGB) + VH.A(189477) + current.RGB + VH.A(186866) + A(current.RGB) + VH.A(189155));
				}
				else
				{
					stringBuilder.Append(VH.A(189434) + Conversions.ToString(num) + VH.A(187085) + A(num, current.RGB) + VH.A(189538));
				}
				num = checked(num + 1);
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_0156;
				}
				continue;
				end_IL_0156:
				break;
			}
		}
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	public static string AutoColorMenu()
	{
		int num = 1;
		int num2 = 0;
		StringBuilder stringBuilder = A();
		checked
		{
			string text = default(string);
			foreach (string autoColor in KH.A.AutoColors)
			{
				if (Operators.CompareString(autoColor, "", TextCompare: false) != 0)
				{
					switch (num2)
					{
					case 0:
						text = VH.A(189623);
						break;
					case 1:
						text = VH.A(189636);
						break;
					case 2:
						text = VH.A(189665);
						break;
					case 3:
						text = VH.A(189682);
						break;
					case 4:
						text = VH.A(189713);
						break;
					case 5:
						text = VH.A(189742);
						break;
					case 6:
						text = VH.A(189763);
						break;
					case 7:
						text = VH.A(189790);
						break;
					}
					text = VH.A(187028) + num + VH.A(41385) + text;
					stringBuilder.Append(VH.A(189803) + Conversions.ToString(num) + VH.A(187085) + text + VH.A(189373) + autoColor + VH.A(186866) + A(autoColor) + VH.A(189155));
					num++;
				}
				num2++;
			}
			stringBuilder.Append(VH.A(187206));
			return stringBuilder.ToString();
		}
	}

	private static string A(int A, string B)
	{
		if (B.Length > 0)
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
					if (A < 10)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								return VH.A(187028) + Conversions.ToString(A) + VH.A(189846) + clsRibbon.A(B);
							}
						}
					}
					return Conversions.ToString(A) + VH.A(189846) + clsRibbon.A(B);
				}
			}
		}
		if (A < 10)
		{
			return VH.A(187028) + Conversions.ToString(A) + VH.A(189851);
		}
		return Conversions.ToString(A) + VH.A(189851);
	}

	private static string A(string A)
	{
		string[] array = Strings.Split(A, VH.A(2378));
		return VH.A(10515) + array[0] + VH.A(10524) + array[1] + VH.A(10524) + array[2] + VH.A(39904);
	}

	public static string CellsToStandardSizeMenu()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = clsPublish.StandardSizeNodes().GetEnumerator();
			while (enumerator.MoveNext())
			{
				string text = ((XmlNode)enumerator.Current).Attributes[VH.A(67336)].Value.Replace(VH.A(186705), VH.A(187221));
				string text2 = num + VH.A(189846) + text;
				if (num < 10)
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
					text2 = VH.A(187028) + text2;
				}
				stringBuilder.Append(VH.A(189870) + num + VH.A(187085) + text2 + VH.A(189933) + num + VH.A(186866) + text + VH.A(189994));
				num = checked(num + 1);
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
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	public static string ChartResizeMenu()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		stringBuilder.Append(VH.A(191050));
		stringBuilder.Append(VH.A(191149));
		stringBuilder.Append(VH.A(191771));
		stringBuilder.Append(VH.A(192261));
		stringBuilder.Append(VH.A(192711));
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = clsPublish.StandardSizeNodes().GetEnumerator();
			while (enumerator.MoveNext())
			{
				string text = ((XmlNode)enumerator.Current).Attributes[VH.A(67336)].Value.Replace(VH.A(186705), VH.A(187221));
				string text2 = num + VH.A(189846) + text;
				if (num < 10)
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
					text2 = VH.A(187028) + text2;
				}
				stringBuilder.Append(VH.A(192826) + num + VH.A(187085) + text2 + VH.A(192873) + num + VH.A(186866) + text + VH.A(187019));
				num = checked(num + 1);
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_01a2;
				}
				continue;
				end_IL_01a2:
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
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	public static string ShowGuideMenu()
	{
		int num = 1;
		StringBuilder stringBuilder = A();
		stringBuilder.Append(VH.A(192711));
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = clsPublish.StandardSizeNodes().GetEnumerator();
			while (enumerator.MoveNext())
			{
				string text = ((XmlNode)enumerator.Current).Attributes[VH.A(67336)].Value.Replace(VH.A(186705), VH.A(187221));
				string text2 = num + VH.A(189846) + text;
				if (num < 10)
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
					text2 = VH.A(187028) + num + VH.A(189846) + text;
				}
				else
				{
					text2 = ((num != 10) ? (num + VH.A(189846) + text) : (VH.A(192930) + text));
				}
				stringBuilder.Append(VH.A(192949) + num + VH.A(187085) + text2 + VH.A(192992) + num + VH.A(186866) + text + VH.A(187019));
				num = checked(num + 1);
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_019b;
				}
				continue;
				end_IL_019b:
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
		stringBuilder.Append(VH.A(193049));
		stringBuilder.Append(VH.A(193148));
		stringBuilder.Append(VH.A(193556));
		stringBuilder.Append(VH.A(193976));
		stringBuilder.Append(VH.A(194368));
		stringBuilder.Append(VH.A(194431));
		stringBuilder.Append(VH.A(187206));
		return stringBuilder.ToString();
	}

	private static StringBuilder A()
	{
		return new StringBuilder(VH.A(194715));
	}

	public static Bitmap RecolorColorButton(string id)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Graphics graphics = default(Graphics);
		Bitmap bitmap = default(Bitmap);
		string left = default(string);
		Pen pen = default(Pen);
		Rectangle rect = default(Rectangle);
		SolidBrush solidBrush = default(SolidBrush);
		Color color = default(Color);
		Bitmap result = default(Bitmap);
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
				case 620:
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
						case 4:
							goto IL_000c;
						case 5:
							goto IL_0032;
						case 6:
							goto IL_0042;
						case 8:
							goto IL_0055;
						case 9:
							goto IL_0072;
						case 10:
							goto IL_009b;
						case 12:
							goto IL_00a9;
						case 11:
						case 13:
							goto IL_00ba;
						case 15:
							goto IL_00cf;
						case 16:
							goto IL_00e3;
						case 17:
							goto IL_0102;
						case 19:
							goto IL_0110;
						case 18:
						case 20:
							goto IL_011f;
						case 3:
						case 7:
						case 14:
						case 21:
							goto IL_0132;
						case 22:
							goto IL_0152;
						case 23:
							goto IL_015e;
						case 24:
							goto IL_016a;
						case 25:
							goto IL_017a;
						case 26:
							goto IL_0188;
						case 27:
							goto IL_0196;
						case 28:
							goto IL_01a4;
						case 29:
							goto IL_01ae;
						case 30:
							goto IL_01b8;
						case 31:
							goto IL_01c2;
						case 32:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 33:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_017a:
					num2 = 25;
					graphics = Graphics.FromImage(bitmap);
					goto IL_0188;
					IL_0007:
					num2 = 2;
					left = id;
					goto IL_000c;
					IL_000c:
					num2 = 4;
					if (Operators.CompareString(left, clsColors.LAST_FONT_COLOR_BUTTON, TextCompare: false) == 0)
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
						goto IL_0032;
					}
					goto IL_0055;
					IL_0188:
					num2 = 26;
					graphics.DrawRectangle(pen, rect);
					goto IL_0196;
					IL_0196:
					num2 = 27;
					graphics.FillRectangle(solidBrush, rect);
					goto IL_01a4;
					IL_01a4:
					num2 = 28;
					graphics.Dispose();
					goto IL_01ae;
					IL_0032:
					num2 = 5;
					bitmap = new Bitmap(J.FontColorPicker);
					goto IL_0042;
					IL_0042:
					num2 = 6;
					color = K.Settings.LastFontColor;
					goto IL_0132;
					IL_0055:
					num2 = 8;
					if (Operators.CompareString(left, clsColors.LAST_FILL_COLOR_BUTTON, TextCompare: false) == 0)
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
						goto IL_0072;
					}
					goto IL_00cf;
					IL_01ae:
					num2 = 29;
					pen.Dispose();
					goto IL_01b8;
					IL_0072:
					num2 = 9;
					if (K.Settings.LastFillColor == Color.Transparent)
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
						goto IL_009b;
					}
					goto IL_00a9;
					IL_01b8:
					num2 = 30;
					solidBrush.Dispose();
					goto IL_01c2;
					IL_009b:
					num2 = 10;
					bitmap = J.NoFill;
					goto IL_00ba;
					IL_00a9:
					num2 = 12;
					bitmap = new Bitmap(J.FillColorPicker);
					goto IL_00ba;
					IL_00ba:
					num2 = 13;
					color = K.Settings.LastFillColor;
					goto IL_0132;
					IL_00cf:
					num2 = 15;
					if (Operators.CompareString(left, clsColors.LAST_BORDER_COLOR_BUTTON, TextCompare: false) == 0)
					{
						goto IL_00e3;
					}
					goto IL_0132;
					IL_00e3:
					num2 = 16;
					if (K.Settings.LastBorderColor == Color.Transparent)
					{
						goto IL_0102;
					}
					goto IL_0110;
					IL_0102:
					num2 = 17;
					bitmap = J.NoBorder;
					goto IL_011f;
					IL_0110:
					num2 = 19;
					bitmap = new Bitmap(J.BorderColorPicker);
					goto IL_011f;
					IL_011f:
					num2 = 20;
					color = K.Settings.LastBorderColor;
					goto IL_0132;
					IL_0132:
					num2 = 21;
					if (!(color != Color.Transparent))
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
						break;
					}
					goto IL_0152;
					IL_01c2:
					num2 = 31;
					rect = default(Rectangle);
					break;
					IL_0152:
					num2 = 22;
					pen = new Pen(color);
					goto IL_015e;
					IL_015e:
					num2 = 23;
					solidBrush = new SolidBrush(color);
					goto IL_016a;
					IL_016a:
					num2 = 24;
					rect = new Rectangle(0, 12, 16, 4);
					goto IL_017a;
					end_IL_0000_2:
					break;
				}
				num2 = 32;
				result = bitmap;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 620;
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
			ProjectData.ClearProjectError();
		}
		return result;
	}
}
