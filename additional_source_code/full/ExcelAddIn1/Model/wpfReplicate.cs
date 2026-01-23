using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Markup;
using System.Windows.Media;
using System.Xml;
using A;
using ExcelAddIn1.Workbook;
using Foo.Controls;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Model;

[DesignerGenerated]
public sealed class wpfReplicate : System.Windows.Window, IComponentConnector, IStyleConnector
{
	private Microsoft.Office.Interop.Excel.Application m_A;

	private Range m_A;

	private int m_A;

	private int m_B;

	private ObservableCollection<SumRow> m_A;

	[AccessedThroughProperty("numCopies")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("optSame")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lblSame")]
	private TextBlock m_A;

	[AccessedThroughProperty("numSpacer")]
	[CompilerGenerated]
	private MacNumericUpDown m_B;

	[AccessedThroughProperty("optAbove")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("optBelow")]
	private System.Windows.Controls.RadioButton m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("optMultiple")]
	private System.Windows.Controls.RadioButton m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("lblMultiple")]
	private TextBlock m_B;

	[AccessedThroughProperty("cbxBaseName")]
	[CompilerGenerated]
	private System.Windows.Controls.ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSum")]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lblRows")]
	private TextBlock m_C;

	[AccessedThroughProperty("lbxRows")]
	[CompilerGenerated]
	private System.Windows.Controls.ListBox m_A;

	[AccessedThroughProperty("btnReplicate")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private System.Windows.Controls.Button m_B;

	private bool m_A;

	internal virtual MacNumericUpDown numCopies
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton optSame
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual TextBlock lblSame
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual MacNumericUpDown numSpacer
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton optAbove
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton optBelow
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton optMultiple
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal virtual TextBlock lblMultiple
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	internal virtual System.Windows.Controls.ComboBox cbxBaseName
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSum
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual TextBlock lblRows
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_C = value;
		}
	}

	internal virtual System.Windows.Controls.ListBox lbxRows
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnReplicate
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnReplicate_Click;
			System.Windows.Controls.Button button = this.m_A;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnCancel
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnCancel_Click;
			System.Windows.Controls.Button button = this.m_B;
			if (button != null)
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
				button.Click -= value2;
			}
			this.m_B = value;
			button = this.m_B;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	public wpfReplicate()
	{
		base.Loaded += wpfReplicate_Loaded;
		base.Closing += CleanUp;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = MH.A.Application;
		this.m_A = (Range)this.m_A.Selection;
		this.m_A = Conversions.ToInteger(this.m_A.Rows.CountLarge);
		this.m_B = Conversions.ToInteger(this.m_A.Columns.CountLarge);
	}

	private void wpfReplicate_Loaded(object sender, RoutedEventArgs e)
	{
		base.MinHeight = base.ActualHeight;
		base.MaxHeight = base.ActualHeight;
		optSame.Checked += optSame_CheckedChanged;
		optSame.Unchecked += optSame_CheckedChanged;
		optMultiple.Checked += optMultiple_CheckedChanged;
		optMultiple.Unchecked += optMultiple_CheckedChanged;
		chkSum.Checked += chkSum_CheckedChanged;
		chkSum.Unchecked += chkSum_CheckedChanged;
		try
		{
			XmlNode xmlNode = KH.A.SettingsXml.DocumentElement.SelectSingleNode(VH.A(96265));
			numCopies.Value = Conversions.ToDouble(xmlNode.SelectSingleNode(VH.A(96296)).InnerText);
			bool flag = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(96321)).InnerText);
			optSame.IsChecked = flag;
			optMultiple.IsChecked = !flag;
			chkSum.IsChecked = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(96340)).InnerText);
			numSpacer.Value = Conversions.ToDouble(xmlNode.SelectSingleNode(VH.A(96359)).InnerText);
			optBelow.IsChecked = Conversions.ToBoolean(xmlNode.SelectSingleNode(VH.A(96380)).InnerText);
			System.Windows.Controls.RadioButton radioButton = optAbove;
			bool? isChecked = optBelow.IsChecked;
			bool? isChecked2;
			if (!isChecked.HasValue)
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
				isChecked2 = isChecked;
			}
			else
			{
				isChecked2 = isChecked != true;
			}
			radioButton.IsChecked = isChecked2;
			xmlNode = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			optSame.IsChecked = true;
			optMultiple.IsChecked = false;
			optBelow.IsChecked = true;
			optAbove.IsChecked = false;
			ProjectData.ClearProjectError();
		}
		this.m_A = new ObservableCollection<SumRow>();
		int a = this.m_A;
		bool blnChecked;
		string strDisplay;
		SolidColorBrush brush;
		Range range;
		for (int i = 1; i <= a; this.m_A.Add(new SumRow(strDisplay, brush, blnChecked)), brush = null, range = null, i = checked(i + 1))
		{
			blnChecked = false;
			string text = Conversions.ToString(NewLateBinding.LateGet(this.m_A.Cells[i, 1], null, VH.A(96399), new object[0], null, null, null));
			range = (Range)this.m_A.Rows[i, RuntimeHelpers.GetObjectValue(Missing.Value)];
			if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
				strDisplay = VH.A(96408) + Conversions.ToString(range.Row) + VH.A(96417) + text;
				brush = A();
				if (RangeHelpers.CellsWithNumbers(range) == null)
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
				string text2 = text.ToUpper();
				if (!text2.StartsWith(VH.A(96424)))
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
					if (!text2.StartsWith(VH.A(96437)))
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
						if (text2.StartsWith(VH.A(96446)))
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
							if (!text2.StartsWith(VH.A(96459)))
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
								if (!text2.StartsWith(VH.A(96486)))
								{
									goto IL_04c8;
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
						}
						if (!text2.StartsWith(VH.A(96509)) && !text2.StartsWith(VH.A(78355)))
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
							if (!text2.StartsWith(VH.A(96528)))
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
								if (!text2.Contains(VH.A(96547)))
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
									if (!text2.Contains(VH.A(96560)))
									{
										blnChecked = true;
										continue;
									}
								}
							}
						}
					}
				}
				goto IL_04c8;
			}
			strDisplay = VH.A(96408) + Conversions.ToString(range.Row);
			brush = B();
			continue;
			IL_04c8:
			try
			{
				range = range.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				blnChecked = true;
				ProjectData.ClearProjectError();
			}
		}
		while (true)
		{
			switch (7)
			{
			case 0:
				continue;
			}
			lbxRows.ItemsSource = this.m_A;
			return;
		}
	}

	private void optSame_CheckedChanged(object sender, RoutedEventArgs e)
	{
		if (optSame.IsChecked == true)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					lblMultiple.IsEnabled = false;
					cbxBaseName.IsEnabled = false;
					optMultiple.IsChecked = false;
					return;
				}
			}
		}
		lblMultiple.IsEnabled = true;
		cbxBaseName.IsEnabled = true;
	}

	private void optMultiple_CheckedChanged(object sender, RoutedEventArgs e)
	{
		if (optMultiple.IsChecked == true)
		{
			lblSame.IsEnabled = false;
			((UIElement)(object)numSpacer).IsEnabled = false;
			optAbove.IsEnabled = false;
			optBelow.IsEnabled = false;
			optSame.IsChecked = false;
		}
		else
		{
			lblSame.IsEnabled = true;
			((UIElement)(object)numSpacer).IsEnabled = true;
			optAbove.IsEnabled = true;
			optBelow.IsEnabled = true;
		}
	}

	private void chkSum_CheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkSum.IsChecked.Value;
		lblRows.IsEnabled = value;
		lbxRows.IsEnabled = value;
		if (!value)
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
			lbxRows.Focus();
			return;
		}
	}

	private void btnReplicate_Click(object sender, RoutedEventArgs e)
	{
		checked
		{
			int num = (int)Math.Round(numCopies.Value.Value);
			XlCalculation calculation = default(XlCalculation);
			List<Range> D;
			Range range;
			Worksheet worksheet;
			Microsoft.Office.Interop.Excel.Workbook workbook;
			try
			{
				worksheet = this.m_A.Worksheet;
				workbook = (Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent;
				if (!workbook.Saved)
				{
					DialogResult dialogResult = Forms.YesNoCancelMessage(VH.A(96581));
					if (dialogResult != System.Windows.Forms.DialogResult.OK)
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
						if (dialogResult == System.Windows.Forms.DialogResult.Cancel)
						{
							worksheet = null;
							workbook = null;
							return;
						}
					}
					else
					{
						QuickSave.Save(workbook);
					}
				}
				D = new List<Range>();
				Microsoft.Office.Interop.Excel.Application a = this.m_A;
				a.ScreenUpdating = false;
				a.EnableEvents = false;
				calculation = a.Calculation;
				a.Calculation = XlCalculation.xlCalculationManual;
				_ = null;
				if (optSame.IsChecked == true)
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
					int num2 = (int)Math.Round(numSpacer.Value.Value);
					int num3 = this.m_A + num2;
					double standardHeight = worksheet.StandardHeight;
					Worksheet worksheet2 = (Worksheet)workbook.Worksheets.Add(RuntimeHelpers.GetObjectValue(Missing.Value), worksheet, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					if (optBelow.IsChecked == true)
					{
						object instance = this.m_A.Cells[this.m_A, 1];
						((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, VH.A(60565), new object[2] { 1, 0 }, null, null, null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, VH.A(60565), new object[2]
						{
							num3 * num,
							0
						}, null, null, null))).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
						instance = null;
						int num4 = num;
						for (int i = 1; i <= num4; i++)
						{
							A(worksheet2);
							range = (Range)NewLateBinding.LateGet(this.m_A.Cells[1, 1], null, VH.A(60565), new object[2]
							{
								num3 * i,
								0
							}, null, null, null);
							string cell = range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							worksheet.Paste(RuntimeHelpers.GetObjectValue(worksheet.Cells[range.Row, 1]), RuntimeHelpers.GetObjectValue(Missing.Value));
							range = ((_Worksheet)worksheet).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value));
							A(worksheet2, worksheet, range, ref D);
							int num5 = num2;
							for (int j = 1; j <= num5; j++)
							{
								Range range2 = range.get_Offset((object)(-j), (object)0);
								range2.RowHeight = standardHeight;
								range2.EntireRow.ClearFormats();
								_ = null;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_038d;
								}
								continue;
								end_IL_038d:
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
					else
					{
						object instance2 = this.m_A.Cells[1, 1];
						((_Worksheet)worksheet).get_Range(RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance2, null, VH.A(60565), new object[2] { 0, 0 }, null, null, null)), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance2, null, VH.A(60565), new object[2]
						{
							num3 * num - 1,
							0
						}, null, null, null))).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
						instance2 = null;
						string cell2 = this.m_A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						int num6 = num;
						for (int k = 1; k <= num6; k++)
						{
							A(worksheet2);
							range = (Range)NewLateBinding.LateGet(this.m_A.Cells[1, 1], null, VH.A(60565), new object[2]
							{
								-num3 * k,
								0
							}, null, null, null);
							string cell = range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							worksheet.Paste(RuntimeHelpers.GetObjectValue(worksheet.Cells[range.Row, 1]), RuntimeHelpers.GetObjectValue(Missing.Value));
							range = ((_Worksheet)worksheet).get_Range((object)cell, RuntimeHelpers.GetObjectValue(Missing.Value));
							A(worksheet2, worksheet, range, ref D);
							int num7 = num2;
							for (int l = 1; l <= num7; l++)
							{
								Range range3 = range.get_Offset((object)(l + this.m_A), (object)0);
								range3.RowHeight = standardHeight;
								range3.EntireRow.ClearFormats();
								_ = null;
							}
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
									goto end_IL_0611;
								}
								continue;
								end_IL_0611:
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
						this.m_A = ((_Worksheet)worksheet).get_Range((object)cell2, RuntimeHelpers.GetObjectValue(Missing.Value));
					}
					range = null;
					this.m_A.DisplayAlerts = false;
					try
					{
						worksheet2.Delete();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					this.m_A.DisplayAlerts = true;
					worksheet2 = null;
					if (chkSum.IsChecked == true)
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
						int num8 = clsColors.RGB2Ole(KH.A.AutoColors[2]);
						IEnumerator<SumRow> enumerator = default(IEnumerator<SumRow>);
						try
						{
							enumerator = this.m_A.GetEnumerator();
							IEnumerator enumerator2 = default(IEnumerator);
							while (enumerator.MoveNext())
							{
								SumRow current = enumerator.Current;
								if (!current.IsChecked)
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
								int num9 = this.m_A.IndexOf(current) + 1;
								Range range4 = RangeHelpers.CellsWithNumbers((Range)this.m_A.Rows[num9, RuntimeHelpers.GetObjectValue(Missing.Value)]);
								if (range4 == null)
								{
									continue;
								}
								try
								{
									enumerator2 = range4.GetEnumerator();
									while (enumerator2.MoveNext())
									{
										range = (Range)enumerator2.Current;
										if (Information.IsDate(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
										{
											continue;
										}
										int num10 = Conversions.ToInteger(Operators.AddObject(Operators.SubtractObject(range.Column, NewLateBinding.LateGet(this.m_A.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(41354), new object[0], null, null, null)), 1));
										string text = VH.A(48936);
										using (List<Range>.Enumerator enumerator3 = D.GetEnumerator())
										{
											while (enumerator3.MoveNext())
											{
												Range current2 = enumerator3.Current;
												text = Conversions.ToString(Operators.ConcatenateObject(text, Operators.ConcatenateObject(NewLateBinding.LateGet(current2.Cells[num9, num10], null, VH.A(5814), new object[2] { 0, 0 }, null, null, null), VH.A(54459))));
											}
											while (true)
											{
												switch (1)
												{
												case 0:
													break;
												default:
													goto end_IL_0895;
												}
												continue;
												end_IL_0895:
												break;
											}
										}
										text = Strings.Left(text, text.Length - 1);
										range.Formula = text;
										range.Font.Color = num8;
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_08ec;
										}
										continue;
										end_IL_08ec:
										break;
									}
								}
								finally
								{
									if (enumerator2 is IDisposable)
									{
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											(enumerator2 as IDisposable).Dispose();
											break;
										}
									}
								}
								range4 = null;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_0929;
								}
								continue;
								end_IL_0929:
								break;
							}
						}
						finally
						{
							if (enumerator != null)
							{
								while (true)
								{
									switch (6)
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
					this.m_A.Worksheet.Activate();
					if (optAbove.IsChecked == true)
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
						this.m_A.Select();
					}
				}
				else if (optMultiple.IsChecked == true)
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
					bool displayGridlines = this.m_A.ActiveWindow.DisplayGridlines;
					Range visibleRange = this.m_A.ActiveWindow.VisibleRange;
					int num11 = num;
					for (int m = 1; m <= num11; m++)
					{
						worksheet = (Worksheet)workbook.Worksheets.Add(RuntimeHelpers.GetObjectValue(Missing.Value), worksheet, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						worksheet.Name = cbxBaseName.Text + VH.A(41385) + Conversions.ToString(m);
						Microsoft.Office.Interop.Excel.Window activeWindow = this.m_A.ActiveWindow;
						if (activeWindow.DisplayGridlines != displayGridlines)
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
							activeWindow.DisplayGridlines = displayGridlines;
						}
						activeWindow = null;
						B(worksheet);
						C(worksheet);
						this.m_A.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
						range = (Range)worksheet.Cells[this.m_A.Row, this.m_A.Column];
						range.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						if (chkSum.IsChecked == true)
						{
							D.Add(((_Worksheet)worksheet).get_Range((object)range, (object)range.get_Offset((object)(this.m_A - 1), (object)(this.m_B - 1))));
						}
						Microsoft.Office.Interop.Excel.Window activeWindow2 = this.m_A.ActiveWindow;
						activeWindow2.ScrollRow = visibleRange.Row;
						activeWindow2.ScrollColumn = visibleRange.Column;
						_ = null;
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
					range = null;
					if (chkSum.IsChecked == true)
					{
						int num8 = clsColors.RGB2Ole(KH.A.AutoColors[3]);
						IEnumerator<SumRow> enumerator4 = default(IEnumerator<SumRow>);
						try
						{
							enumerator4 = this.m_A.GetEnumerator();
							IEnumerator enumerator5 = default(IEnumerator);
							while (enumerator4.MoveNext())
							{
								SumRow current3 = enumerator4.Current;
								Range range4 = RangeHelpers.CellsWithNumbers((Range)this.m_A.Rows[this.m_A.IndexOf(current3) + 1, RuntimeHelpers.GetObjectValue(Missing.Value)]);
								if (range4 == null)
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
								try
								{
									enumerator5 = range4.GetEnumerator();
									while (enumerator5.MoveNext())
									{
										range = (Range)enumerator5.Current;
										if (Information.IsDate(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
										string text = VH.A(48936);
										using (List<Range>.Enumerator enumerator6 = D.GetEnumerator())
										{
											while (enumerator6.MoveNext())
											{
												Range current4 = enumerator6.Current;
												text = text + ((_Worksheet)current4.Worksheet).get_Range((object)range.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)), RuntimeHelpers.GetObjectValue(Missing.Value)).get_Address((object)0, (object)0, XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(54459);
											}
											while (true)
											{
												switch (2)
												{
												case 0:
													break;
												default:
													goto end_IL_0d6f;
												}
												continue;
												end_IL_0d6f:
												break;
											}
										}
										text = Strings.Left(text, text.Length - 1);
										range.Formula = text;
										range.Font.Color = num8;
									}
								}
								finally
								{
									if (enumerator5 is IDisposable)
									{
										while (true)
										{
											switch (4)
											{
											case 0:
												continue;
											}
											(enumerator5 as IDisposable).Dispose();
											break;
										}
									}
								}
								range4 = null;
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_0df7;
								}
								continue;
								end_IL_0df7:
								break;
							}
						}
						finally
						{
							if (enumerator4 != null)
							{
								while (true)
								{
									switch (7)
									{
									case 0:
										continue;
									}
									enumerator4.Dispose();
									break;
								}
							}
						}
					}
					this.m_A.Worksheet.Activate();
				}
				XmlDocument settingsXml = KH.A.SettingsXml;
				XmlNode xmlNode = settingsXml.DocumentElement.SelectSingleNode(VH.A(96265));
				xmlNode.SelectSingleNode(VH.A(96296)).InnerText = Conversions.ToString(numCopies.Value.Value);
				xmlNode.SelectSingleNode(VH.A(96321)).InnerText = Conversions.ToString(optSame.IsChecked.Value);
				xmlNode.SelectSingleNode(VH.A(96340)).InnerText = Conversions.ToString(chkSum.IsChecked.Value);
				xmlNode.SelectSingleNode(VH.A(96359)).InnerText = Conversions.ToString(numSpacer.Value.Value);
				xmlNode.SelectSingleNode(VH.A(96380)).InnerText = Conversions.ToString(optBelow.IsChecked.Value);
				_ = null;
				_ = null;
				KH.A.SaveSettings(settingsXml);
				settingsXml = null;
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.ErrorMessage(ex4.Message);
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
			}
			this.m_A.Calculation = calculation;
			this.m_A.ScreenUpdating = true;
			this.m_A.EnableEvents = true;
			this.m_A.CutCopyMode = (XlCutCopyMode)0;
			D = null;
			range = null;
			worksheet = null;
			workbook = null;
			base.DialogResult = true;
		}
	}

	private void A(Worksheet A)
	{
		B(A);
		this.m_A.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
		Range range = (Range)A.Cells[this.m_A.Row, this.m_A.Column];
		range.PasteSpecial(XlPasteType.xlPasteAll, XlPasteSpecialOperation.xlPasteSpecialOperationNone, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		range = ((_Worksheet)A).get_Range((object)range, (object)checked(range.get_Offset((object)(this.m_A - 1), (object)(this.m_B - 1))));
		range.EntireRow.Cut(RuntimeHelpers.GetObjectValue(Missing.Value));
		range = null;
	}

	private void B(Worksheet A)
	{
		this.m_A.EntireRow.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
		A.Paste(RuntimeHelpers.GetObjectValue(A.Cells[this.m_A.Row, 1]), RuntimeHelpers.GetObjectValue(Missing.Value));
		D(A);
	}

	private void C(Worksheet A)
	{
		this.m_A.EntireColumn.Copy(RuntimeHelpers.GetObjectValue(Missing.Value));
		A.Paste(RuntimeHelpers.GetObjectValue(A.Cells[1, this.m_A.Column]), RuntimeHelpers.GetObjectValue(Missing.Value));
		D(A);
	}

	private void D(Worksheet A)
	{
		Worksheet worksheet = A;
		worksheet.Cells.Clear();
		try
		{
			((ChartObjects)worksheet.ChartObjects(RuntimeHelpers.GetObjectValue(Missing.Value))).Delete();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		for (int i = worksheet.Shapes.Count; i >= 1; i = checked(i + -1))
		{
			worksheet.Shapes.Item(i).Delete();
		}
		worksheet = null;
	}

	private void A(Worksheet A, Worksheet B, Range C, ref List<Range> D)
	{
		Range range = ((_Worksheet)B).get_Range((object)C, (object)checked(C.get_Offset((object)(this.m_A - 1), (object)(this.m_B - 1))));
		if (chkSum.IsChecked == true)
		{
			D.Add(range);
		}
		this.A(range, A.Name);
		range = null;
	}

	private void A(Range A, string B)
	{
		Range range = null;
		try
		{
			range = A.SpecialCells(XlCellType.xlCellTypeFormulas, RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (range == null)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			B = B.Replace(VH.A(39851), VH.A(39854));
			try
			{
				enumerator = range.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range obj = (Range)enumerator.Current;
					string text = Conversions.ToString(obj.Formula);
					if (text.Contains(VH.A(39851) + B + VH.A(43343)))
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
						text = text.Replace(VH.A(39851) + B + VH.A(43343), "");
					}
					else if (text.Contains(B + VH.A(7827)))
					{
						text = text.Replace(B + VH.A(7827), "");
					}
					obj.Formula = text;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_013f;
					}
					continue;
					end_IL_013f:
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
			range = null;
			return;
		}
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private void CleanUp(object sender, CancelEventArgs e)
	{
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
	}

	private void SumRowChecked(object sender, RoutedEventArgs e)
	{
		((SumRow)((System.Windows.Controls.CheckBox)sender).DataContext).TextColor = A();
	}

	private void SumRowUnhecked(object sender, RoutedEventArgs e)
	{
		((SumRow)((System.Windows.Controls.CheckBox)sender).DataContext).TextColor = B();
	}

	private SolidColorBrush A()
	{
		return new SolidColorBrush(SystemColors.ControlTextColor);
	}

	private SolidColorBrush B()
	{
		return new SolidColorBrush(SystemColors.ControlDarkColor);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_A)
		{
			return;
		}
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
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(96755), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0010: Expected O, but got Unknown
		//IL_0060: Unknown result type (might be due to invalid IL or missing references)
		//IL_006a: Expected O, but got Unknown
		if (connectionId == 1)
		{
			numCopies = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					optSame = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					lblSame = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					numSpacer = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					optAbove = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					optBelow = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					optMultiple = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			lblMultiple = (TextBlock)target;
			return;
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					cbxBaseName = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			chkSum = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					lblRows = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					lbxRows = (System.Windows.Controls.ListBox)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnReplicate = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnCancel = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		this.m_A = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId != 13)
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
			((System.Windows.Controls.CheckBox)target).Checked += SumRowChecked;
			((System.Windows.Controls.CheckBox)target).Unchecked += SumRowUnhecked;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
