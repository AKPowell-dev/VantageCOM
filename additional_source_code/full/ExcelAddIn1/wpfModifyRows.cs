using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Xml;
using A;
using ExcelAddIn1.Format;
using Foo.Controls;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[DesignerGenerated]
public sealed class wpfModifyRows : System.Windows.Window, IComponentConnector
{
	private Microsoft.Office.Interop.Excel.Application m_A;

	private Range m_A;

	private bool m_A;

	private List<XmlNode> m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("optFormat")]
	private RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lbl1")]
	private TextBlock m_A;

	[AccessedThroughProperty("numInsert")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lbl2")]
	private TextBlock B;

	[AccessedThroughProperty("optInsert")]
	[CompilerGenerated]
	private RadioButton B;

	[AccessedThroughProperty("lbl3")]
	[CompilerGenerated]
	private TextBlock C;

	[CompilerGenerated]
	[AccessedThroughProperty("numEvery")]
	private MacNumericUpDown B;

	[AccessedThroughProperty("lbl4")]
	[CompilerGenerated]
	private TextBlock D;

	[AccessedThroughProperty("optDelete")]
	[CompilerGenerated]
	private RadioButton C;

	[AccessedThroughProperty("lbl5")]
	[CompilerGenerated]
	private TextBlock E;

	[CompilerGenerated]
	[AccessedThroughProperty("numRows")]
	private MacNumericUpDown C;

	[CompilerGenerated]
	[AccessedThroughProperty("lbl6")]
	private TextBlock F;

	[CompilerGenerated]
	[AccessedThroughProperty("grpFormat")]
	private GroupBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("optInherit")]
	private RadioButton D;

	[AccessedThroughProperty("optStyle")]
	[CompilerGenerated]
	private RadioButton E;

	[AccessedThroughProperty("cbxStyles")]
	[CompilerGenerated]
	private ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("optClear")]
	private RadioButton F;

	[CompilerGenerated]
	[AccessedThroughProperty("grpHeight")]
	private GroupBox B;

	[AccessedThroughProperty("optNoChange")]
	[CompilerGenerated]
	private RadioButton G;

	[CompilerGenerated]
	[AccessedThroughProperty("optSetTo")]
	private RadioButton H;

	[AccessedThroughProperty("numRowHeight")]
	[CompilerGenerated]
	private MacNumericUpDown D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSave")]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button m_A;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private Button B;

	private bool B;

	internal virtual RadioButton optFormat
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

	internal virtual TextBlock lbl1
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

	internal virtual MacNumericUpDown numInsert
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

	internal virtual TextBlock lbl2
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual RadioButton optInsert
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual TextBlock lbl3
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.C = value;
		}
	}

	internal virtual MacNumericUpDown numEvery
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual TextBlock lbl4
	{
		[CompilerGenerated]
		get
		{
			return this.D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.D = value;
		}
	}

	internal virtual RadioButton optDelete
	{
		[CompilerGenerated]
		get
		{
			return this.C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = optDelete_CheckedChanged;
			RoutedEventHandler value3 = optDelete_CheckedChanged;
			RadioButton radioButton = this.C;
			if (radioButton != null)
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
				radioButton.Checked -= value2;
				radioButton.Unchecked -= value3;
			}
			this.C = value;
			radioButton = this.C;
			if (radioButton == null)
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
				radioButton.Checked += value2;
				radioButton.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual TextBlock lbl5
	{
		[CompilerGenerated]
		get
		{
			return this.E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.E = value;
		}
	}

	internal virtual MacNumericUpDown numRows
	{
		[CompilerGenerated]
		get
		{
			return C;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			//IL_0007: Unknown result type (might be due to invalid IL or missing references)
			//IL_000d: Expected O, but got Unknown
			MacRangeBaseValueChangedHandler val = new MacRangeBaseValueChangedHandler(numRows_ValueChanged);
			MacNumericUpDown val2 = C;
			if (val2 != null)
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
				val2.ValueChanged -= val;
			}
			C = value;
			val2 = C;
			if (val2 == null)
			{
				return;
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				val2.ValueChanged += val;
				return;
			}
		}
	}

	internal virtual TextBlock lbl6
	{
		[CompilerGenerated]
		get
		{
			return this.F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.F = value;
		}
	}

	internal virtual GroupBox grpFormat
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

	internal virtual RadioButton optInherit
	{
		[CompilerGenerated]
		get
		{
			return this.D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.D = value;
		}
	}

	internal virtual RadioButton optStyle
	{
		[CompilerGenerated]
		get
		{
			return E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = optCustom_CheckedChanged;
			RoutedEventHandler value3 = optCustom_CheckedChanged;
			RadioButton radioButton = E;
			if (radioButton != null)
			{
				radioButton.Checked -= value2;
				radioButton.Unchecked -= value3;
			}
			E = value;
			radioButton = E;
			if (radioButton != null)
			{
				radioButton.Checked += value2;
				radioButton.Unchecked += value3;
			}
		}
	}

	internal virtual ComboBox cbxStyles
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

	internal virtual RadioButton optClear
	{
		[CompilerGenerated]
		get
		{
			return F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			F = value;
		}
	}

	internal virtual GroupBox grpHeight
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.B = value;
		}
	}

	internal virtual RadioButton optNoChange
	{
		[CompilerGenerated]
		get
		{
			return G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			G = value;
		}
	}

	internal virtual RadioButton optSetTo
	{
		[CompilerGenerated]
		get
		{
			return H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = optSetTo_CheckedChanged;
			RoutedEventHandler value3 = optSetTo_CheckedChanged;
			RadioButton radioButton = H;
			if (radioButton != null)
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
				radioButton.Checked -= value2;
				radioButton.Unchecked -= value3;
			}
			H = value;
			radioButton = H;
			if (radioButton == null)
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
				radioButton.Checked += value2;
				radioButton.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual MacNumericUpDown numRowHeight
	{
		[CompilerGenerated]
		get
		{
			return D;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			D = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSave
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

	internal virtual Button btnCancel
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
			RoutedEventHandler value2 = btnCancel_Click;
			Button button = this.m_A;
			if (button != null)
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

	internal virtual Button btnOk
	{
		[CompilerGenerated]
		get
		{
			return this.B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnOk_Click;
			Button button = this.B;
			if (button != null)
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
				button.Click -= value2;
			}
			this.B = value;
			button = this.B;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	public wpfModifyRows()
	{
		base.Loaded += wpfModifyRows_Loaded;
		base.Closing += wpfModifyRows_Closing;
		this.m_A = true;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = MH.A.Application;
		this.m_A = (Range)this.m_A.Selection;
	}

	private void wpfModifyRows_Loaded(object sender, RoutedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		List<XmlNode>.Enumerator enumerator = default(List<XmlNode>.Enumerator);
		KeyValuePair<string, List<XmlNode>> current = default(KeyValuePair<string, List<XmlNode>>);
		List<string> list = default(List<string>);
		Dictionary<string, List<XmlNode>>.Enumerator enumerator2 = default(Dictionary<string, List<XmlNode>>.Enumerator);
		string key = default(string);
		XmlNode xmlNode = default(XmlNode);
		string value = default(string);
		XmlNode current2 = default(XmlNode);
		string value2 = default(string);
		string value3 = default(string);
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
				case 1550:
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
							goto IL_0037;
						case 5:
							goto IL_0044;
						case 6:
							goto IL_0067;
						case 7:
							goto IL_0072;
						case 8:
							goto IL_0093;
						case 9:
							goto IL_00c8;
						case 10:
							goto IL_00d8;
						case 11:
							goto IL_00f7;
						case 12:
							goto IL_0107;
						case 13:
							goto IL_0122;
						case 14:
							goto IL_0132;
						case 15:
							goto IL_0143;
						case 16:
							goto IL_0148;
						case 17:
							goto IL_0164;
						case 18:
							goto IL_0180;
						case 19:
							goto IL_019c;
						case 20:
							goto IL_01c3;
						case 22:
							goto IL_01e9;
						case 23:
							goto IL_020a;
						case 25:
							goto IL_0220;
						case 26:
							goto IL_0235;
						case 28:
							goto IL_024d;
						case 29:
							goto IL_0264;
						case 21:
						case 24:
						case 27:
						case 30:
							goto IL_027a;
						case 31:
							goto IL_02b3;
						case 32:
							goto IL_02f0;
						case 34:
							goto IL_0318;
						case 35:
							goto IL_0339;
						case 37:
							goto IL_0351;
						case 38:
							goto IL_0372;
						case 40:
							goto IL_038a;
						case 41:
							goto IL_03ab;
						case 33:
						case 36:
						case 39:
						case 42:
							goto IL_03c1;
						case 43:
							goto IL_03fb;
						case 45:
							goto IL_041f;
						case 46:
							goto IL_0440;
						case 48:
							goto IL_0458;
						case 49:
							goto IL_0477;
						case 44:
						case 47:
						case 50:
							goto IL_048d;
						case 51:
							goto IL_04ce;
						case 52:
							goto IL_050d;
						case 53:
							goto IL_0510;
						case 54:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 55:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0072:
					num2 = 7;
					enumerator = current.Value.GetEnumerator();
					goto IL_00db;
					IL_0007:
					num2 = 2;
					list = new List<string>();
					goto IL_000f;
					IL_000f:
					num2 = 3;
					numRows.Value = (double?)this.m_A.Rows.CountLarge;
					goto IL_0037;
					IL_0037:
					num2 = 4;
					this.m_A = new List<XmlNode>();
					goto IL_0044;
					IL_0044:
					num2 = 5;
					enumerator2 = KH.A.CustomCycles.GetEnumerator();
					goto IL_010a;
					IL_010a:
					if (enumerator2.MoveNext())
					{
						current = enumerator2.Current;
						goto IL_0067;
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
					goto IL_0122;
					IL_0067:
					num2 = 6;
					key = current.Key;
					goto IL_0072;
					IL_0122:
					num2 = 13;
					((IDisposable)enumerator2/*cast due to .constrained prefix*/).Dispose();
					goto IL_0132;
					IL_0132:
					num2 = 14;
					cbxStyles.ItemsSource = list;
					goto IL_0143;
					IL_0143:
					num2 = 15;
					list = null;
					goto IL_0148;
					IL_0148:
					num2 = 16;
					optFormat.Checked += UpdateLabels;
					goto IL_0164;
					IL_0164:
					num2 = 17;
					optInsert.Checked += UpdateLabels;
					goto IL_0180;
					IL_0180:
					num2 = 18;
					optDelete.Checked += UpdateLabels;
					goto IL_019c;
					IL_019c:
					num2 = 19;
					xmlNode = KH.A.SettingsXml.DocumentElement.SelectSingleNode(VH.A(207397));
					goto IL_01c3;
					IL_01c3:
					num2 = 20;
					value = xmlNode.Attributes[VH.A(206896)].Value;
					goto IL_01e9;
					IL_01e9:
					num2 = 22;
					if (Operators.CompareString(value, Conversions.ToString(0), TextCompare: false) == 0)
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
						goto IL_020a;
					}
					goto IL_0220;
					IL_00db:
					if (enumerator.MoveNext())
					{
						current2 = enumerator.Current;
						goto IL_0093;
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
					goto IL_00f7;
					IL_020a:
					num2 = 23;
					optFormat.IsChecked = true;
					goto IL_027a;
					IL_0220:
					num2 = 25;
					if (Operators.CompareString(value, Conversions.ToString(1), TextCompare: false) == 0)
					{
						goto IL_0235;
					}
					goto IL_024d;
					IL_0235:
					num2 = 26;
					optInsert.IsChecked = true;
					goto IL_027a;
					IL_024d:
					num2 = 28;
					if (Operators.CompareString(value, Conversions.ToString(2), TextCompare: false) == 0)
					{
						goto IL_0264;
					}
					goto IL_027a;
					IL_0264:
					num2 = 29;
					optDelete.IsChecked = true;
					goto IL_027a;
					IL_027a:
					num2 = 30;
					numInsert.Value = Conversions.ToDouble(xmlNode.Attributes[VH.A(206909)].Value);
					goto IL_02b3;
					IL_02b3:
					num2 = 31;
					numEvery.Value = Conversions.ToDouble(xmlNode.Attributes[VH.A(206924)].Value);
					goto IL_02f0;
					IL_02f0:
					num2 = 32;
					value2 = xmlNode.Attributes[VH.A(206939)].Value;
					goto IL_0318;
					IL_0318:
					num2 = 34;
					if (Operators.CompareString(value2, Conversions.ToString(0), TextCompare: false) == 0)
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
						goto IL_0339;
					}
					goto IL_0351;
					IL_0093:
					num2 = 8;
					list.Add(key + VH.A(206862) + current2.SelectSingleNode(VH.A(19019)).InnerText);
					goto IL_00c8;
					IL_0339:
					num2 = 35;
					optInherit.IsChecked = true;
					goto IL_03c1;
					IL_0351:
					num2 = 37;
					if (Operators.CompareString(value2, Conversions.ToString(1), TextCompare: false) == 0)
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
						goto IL_0372;
					}
					goto IL_038a;
					IL_00c8:
					num2 = 9;
					this.m_A.Add(current2);
					goto IL_00d8;
					IL_0372:
					num2 = 38;
					optStyle.IsChecked = true;
					goto IL_03c1;
					IL_038a:
					num2 = 40;
					if (Operators.CompareString(value2, Conversions.ToString(2), TextCompare: false) == 0)
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
						goto IL_03ab;
					}
					goto IL_03c1;
					IL_00d8:
					num2 = 10;
					goto IL_00db;
					IL_03ab:
					num2 = 41;
					optClear.IsChecked = true;
					goto IL_03c1;
					IL_03c1:
					num2 = 42;
					cbxStyles.SelectedIndex = Conversions.ToInteger(xmlNode.Attributes[VH.A(206952)].Value);
					goto IL_03fb;
					IL_03fb:
					num2 = 43;
					value3 = xmlNode.Attributes[VH.A(207418)].Value;
					goto IL_041f;
					IL_041f:
					num2 = 45;
					if (Operators.CompareString(value3, Conversions.ToString(0), TextCompare: false) == 0)
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
						goto IL_0440;
					}
					goto IL_0458;
					IL_00f7:
					num2 = 11;
					((IDisposable)enumerator/*cast due to .constrained prefix*/).Dispose();
					goto IL_0107;
					IL_0440:
					num2 = 46;
					optNoChange.IsChecked = true;
					goto IL_048d;
					IL_0458:
					num2 = 48;
					if (Operators.CompareString(value3, Conversions.ToString(1), TextCompare: false) == 0)
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
						goto IL_0477;
					}
					goto IL_048d;
					IL_0107:
					num2 = 12;
					goto IL_010a;
					IL_0477:
					num2 = 49;
					optSetTo.IsChecked = true;
					goto IL_048d;
					IL_048d:
					num2 = 50;
					numRowHeight.Value = float.Parse(xmlNode.Attributes[VH.A(206974)].Value, CultureInfo.InvariantCulture);
					goto IL_04ce;
					IL_04ce:
					num2 = 51;
					chkSave.IsChecked = Conversions.ToBoolean(xmlNode.Attributes[VH.A(206983)].Value);
					goto IL_050d;
					IL_050d:
					xmlNode = null;
					goto IL_0510;
					IL_0510:
					num2 = 53;
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 54;
				this.m_A = false;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1550;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void UpdateLabels(object sender, RoutedEventArgs e)
	{
		RadioButton radioButton = (RadioButton)sender;
		if (radioButton.IsChecked == true)
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
			string name = radioButton.Name;
			if (Operators.CompareString(name, optFormat.Name, TextCompare: false) == 0)
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
				lbl1.Text = VH.A(207000);
				lbl2.Text = VH.A(207031);
			}
			else if (Operators.CompareString(name, optInsert.Name, TextCompare: false) == 0)
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
				lbl1.Text = VH.A(57289);
				lbl2.Text = VH.A(2877);
			}
			else if (Operators.CompareString(name, optDelete.Name, TextCompare: false) == 0)
			{
				lbl1.Text = VH.A(207036);
				lbl2.Text = VH.A(207031);
			}
		}
		radioButton = null;
	}

	private void numRows_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Range range = default(Range);
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
				case 354:
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
							goto IL_0027;
						case 5:
							goto IL_0030;
						case 6:
							goto IL_00f6;
						case 7:
							goto IL_00f8;
						case 8:
							goto IL_0108;
						case 9:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
						case 10:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0030:
					num2 = 5;
					this.m_A = ((_Worksheet)range.Worksheet).get_Range(RuntimeHelpers.GetObjectValue(range.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)]), RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(range.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, VH.A(60565), new object[1] { numRows.Value - 1.0 }, null, null, null)));
					goto IL_00f6;
					IL_0007:
					num2 = 2;
					if (this.m_A)
					{
						goto end_IL_0000_3;
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
					goto IL_0027;
					IL_00f8:
					num2 = 7;
					this.m_A.Select();
					goto IL_0108;
					IL_0108:
					num2 = 8;
					this.m_A.ScreenUpdating = false;
					break;
					IL_00f6:
					range = null;
					goto IL_00f8;
					IL_0027:
					num2 = 4;
					range = this.m_A;
					goto IL_0030;
					end_IL_0000_2:
					break;
				}
				num2 = 9;
				this.m_A.ScreenUpdating = true;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 354;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void optCustom_CheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = optStyle.IsChecked.Value;
		cbxStyles.IsEnabled = value;
		if (!value)
		{
			cbxStyles.SelectedIndex = -1;
		}
	}

	private void optSetTo_CheckedChanged(object sender, RoutedEventArgs e)
	{
		((UIElement)(object)numRowHeight).IsEnabled = optSetTo.IsChecked.Value;
	}

	private void optDelete_CheckedChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked = optDelete.IsChecked;
		bool value = ((!isChecked) ?? isChecked).Value;
		grpFormat.IsEnabled = value;
		grpHeight.IsEnabled = value;
	}

	private void wpfModifyRows_Closing(object sender, CancelEventArgs e)
	{
		optFormat.Checked -= UpdateLabels;
		optInsert.Checked -= UpdateLabels;
		optDelete.Checked -= UpdateLabels;
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		XmlNode xmlNode = default(XmlNode);
		int num5 = default(int);
		int num6 = default(int);
		int num7 = default(int);
		bool value = default(bool);
		bool value2 = default(bool);
		bool value3 = default(bool);
		float num8 = default(float);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		Range range = default(Range);
		Range range2 = default(Range);
		XmlDocument xmlDocument = default(XmlDocument);
		XmlAttribute xmlAttribute = default(XmlAttribute);
		Range range3 = default(Range);
		XmlAttribute xmlAttribute2 = default(XmlAttribute);
		int num9 = default(int);
		int num10 = default(int);
		int num11 = default(int);
		Range range4 = default(Range);
		int num12 = default(int);
		int num13 = default(int);
		int num14 = default(int);
		int num15 = default(int);
		XmlAttribute xmlAttribute3 = default(XmlAttribute);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					bool? isChecked;
					switch (try0000_dispatch)
					{
					default:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0007;
					case 4344:
						{
							num = num2;
							switch (num3)
							{
							case 1:
								break;
							default:
								goto end_IL_0000;
							}
							int num4 = unchecked(num + 1);
							num = 0;
							switch (num4)
							{
							case 1:
								break;
							case 2:
								goto IL_0007;
							case 3:
								goto IL_002a;
							case 4:
								goto IL_0050;
							case 5:
								goto IL_0072;
							case 6:
								goto IL_008e;
							case 7:
								goto IL_00ac;
							case 8:
								goto IL_00c8;
							case 9:
								goto IL_00e3;
							case 10:
								goto IL_00f6;
							case 11:
								goto IL_0107;
							case 12:
								goto IL_0137;
							case 13:
								goto IL_0167;
							case 14:
								goto IL_0176;
							case 15:
								goto IL_01c8;
							case 16:
								goto IL_01da;
							case 18:
								goto IL_01f3;
							case 19:
								goto IL_0254;
							case 21:
								goto IL_026d;
							case 22:
								goto IL_0278;
							case 23:
								goto IL_0283;
							case 24:
								goto IL_02af;
							case 25:
								goto IL_02be;
							case 26:
								goto IL_02cf;
							case 27:
								goto IL_02ea;
							case 28:
								goto IL_02f0;
							case 29:
								goto IL_0444;
							case 31:
								goto IL_0458;
							case 32:
								goto IL_0488;
							case 33:
								goto IL_048f;
							case 34:
								goto IL_04ae;
							case 36:
								goto IL_04de;
							case 37:
								goto IL_04ef;
							case 35:
							case 38:
								goto IL_0514;
							case 39:
								goto IL_0525;
							case 40:
								goto IL_053e;
							case 30:
							case 41:
								goto IL_0550;
							case 43:
								goto IL_0564;
							case 44:
								goto IL_058c;
							case 45:
								goto IL_059b;
							case 46:
								goto IL_05bb;
							case 47:
								goto IL_05c1;
							case 48:
								goto IL_071b;
							case 50:
								goto IL_072c;
							case 51:
								goto IL_0754;
							case 49:
							case 52:
								goto IL_0770;
							case 54:
								goto IL_0786;
							case 55:
								goto IL_0795;
							case 56:
								goto IL_07b3;
							case 57:
								goto IL_07b9;
							case 58:
								goto IL_0911;
							case 60:
								goto IL_0925;
							case 61:
								goto IL_0a85;
							case 62:
								goto IL_0a96;
							case 64:
								goto IL_0abe;
							case 65:
								goto IL_0acf;
							case 63:
							case 66:
								goto IL_0adc;
							case 67:
								goto IL_0aed;
							case 68:
								goto IL_0afe;
							case 42:
							case 53:
							case 59:
							case 69:
								goto IL_0b1a;
							case 70:
								goto IL_0b44;
							case 71:
								goto IL_0b63;
							case 72:
								goto IL_0b84;
							case 73:
								goto IL_0bad;
							case 75:
								goto IL_0bc1;
							case 76:
								goto IL_0bde;
							case 78:
								goto IL_0bf2;
							case 79:
								goto IL_0c1b;
							case 74:
							case 77:
							case 80:
								goto IL_0c2b;
							case 81:
								goto IL_0c2e;
							case 82:
								goto IL_0c71;
							case 83:
								goto IL_0cb2;
							case 84:
								goto IL_0cd1;
							case 85:
								goto IL_0cf8;
							case 87:
								goto IL_0d0c;
							case 88:
								goto IL_0d33;
							case 90:
								goto IL_0d47;
							case 91:
								goto IL_0d70;
							case 86:
							case 89:
							case 92:
								goto IL_0d82;
							case 93:
								goto IL_0d85;
							case 94:
								goto IL_0dbd;
							case 95:
								goto IL_0ddc;
							case 96:
								goto IL_0e05;
							case 98:
								goto IL_0e19;
							case 97:
							case 99:
								goto IL_0e2b;
							case 100:
								goto IL_0e2e;
							case 101:
								goto IL_0e7c;
							case 102:
								goto IL_0ebd;
							case 103:
								goto IL_0ec0;
							case 104:
								goto IL_0ecf;
							case 105:
								goto IL_0eda;
							case 106:
								goto IL_0ee5;
							case 107:
								goto IL_0eeb;
							case 108:
								goto IL_0eee;
							case 109:
								goto IL_0ef4;
							case 110:
								goto IL_0efa;
							case 111:
								goto IL_0f00;
							case 112:
								goto IL_0f06;
							case 113:
								goto end_IL_0000_2;
							default:
								goto end_IL_0000;
							case 17:
							case 20:
							case 114:
								goto end_IL_0000_3;
							}
							goto default;
						}
						IL_0c71:
						num2 = 82;
						xmlNode.Attributes[VH.A(206924)].Value = Conversions.ToString(numEvery.Value.Value);
						goto IL_0cb2;
						IL_0007:
						num2 = 2;
						num5 = (int)Math.Round(numInsert.Value.Value);
						goto IL_002a;
						IL_002a:
						num2 = 3;
						num6 = (int)Math.Round(numEvery.Value.Value);
						goto IL_0050;
						IL_0050:
						num2 = 4;
						num7 = (int)Math.Round(numRows.Value.Value);
						goto IL_0072;
						IL_0072:
						num2 = 5;
						value = optStyle.IsChecked.Value;
						goto IL_008e;
						IL_008e:
						num2 = 6;
						value2 = optClear.IsChecked.Value;
						goto IL_00ac;
						IL_00ac:
						num2 = 7;
						value3 = optSetTo.IsChecked.Value;
						goto IL_00c8;
						IL_00c8:
						num2 = 8;
						num8 = (float)numRowHeight.Value.Value;
						goto IL_00e3;
						IL_00e3:
						num2 = 9;
						application = MH.A.Application;
						goto IL_00f6;
						IL_00f6:
						num2 = 10;
						_ = this.m_A.Worksheet;
						goto IL_0107;
						IL_0107:
						num2 = 11;
						range = (Range)this.m_A.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
						goto IL_0137;
						IL_0137:
						num2 = 12;
						range2 = (Range)this.m_A.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
						goto IL_0167;
						IL_0167:
						num2 = 13;
						xmlDocument = KH.A.SettingsXml;
						goto IL_0176;
						IL_0176:
						num2 = 14;
						if (optFormat.IsChecked != true)
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
							if (optDelete.IsChecked != true)
							{
								goto IL_01f3;
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
						goto IL_01c8;
						IL_0cb2:
						num2 = 83;
						xmlAttribute = xmlNode.Attributes[VH.A(206939)];
						goto IL_0cd1;
						IL_0cd1:
						num2 = 84;
						if (optInherit.IsChecked == true)
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
							goto IL_0cf8;
						}
						goto IL_0d0c;
						IL_0cf8:
						num2 = 85;
						xmlAttribute.Value = Conversions.ToString(0);
						goto IL_0d82;
						IL_0ef4:
						num2 = 109;
						range2 = null;
						goto IL_0efa;
						IL_0d0c:
						num2 = 87;
						if (optStyle.IsChecked == true)
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
							goto IL_0d33;
						}
						goto IL_0d47;
						IL_01c8:
						num2 = 15;
						if (num5 > num6)
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
							goto IL_01da;
						}
						goto IL_01f3;
						IL_0efa:
						num2 = 110;
						range3 = null;
						goto IL_0f00;
						IL_01da:
						num2 = 16;
						Forms.WarningMessage(VH.A(207431));
						goto end_IL_0000_3;
						IL_01f3:
						num2 = 18;
						isChecked = optStyle.IsChecked;
						if (isChecked.HasValue)
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
							if (isChecked != true)
							{
								goto IL_026d;
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
						if (cbxStyles.SelectedIndex == -1)
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
							if (isChecked.HasValue)
							{
								goto IL_0254;
							}
						}
						goto IL_026d;
						IL_0d47:
						num2 = 90;
						if (optClear.IsChecked == true)
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
							goto IL_0d70;
						}
						goto IL_0d82;
						IL_0f00:
						num2 = 111;
						range = null;
						goto IL_0f06;
						IL_0d70:
						num2 = 91;
						xmlAttribute.Value = Conversions.ToString(2);
						goto IL_0d82;
						IL_0d33:
						num2 = 88;
						xmlAttribute.Value = Conversions.ToString(1);
						goto IL_0d82;
						IL_0d82:
						xmlAttribute = null;
						goto IL_0d85;
						IL_0d85:
						num2 = 93;
						xmlNode.Attributes[VH.A(206952)].Value = Conversions.ToString(cbxStyles.SelectedIndex);
						goto IL_0dbd;
						IL_0254:
						num2 = 19;
						Forms.WarningMessage(VH.A(207239));
						goto end_IL_0000_3;
						IL_026d:
						num2 = 21;
						application.ScreenUpdating = false;
						goto IL_0278;
						IL_0278:
						num2 = 22;
						application.EnableEvents = false;
						goto IL_0283;
						IL_0283:
						num2 = 23;
						if (optInsert.IsChecked == true)
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
							goto IL_02af;
						}
						goto IL_0564;
						IL_0dbd:
						num2 = 94;
						xmlAttribute2 = xmlNode.Attributes[VH.A(207418)];
						goto IL_0ddc;
						IL_02af:
						num2 = 24;
						num9 = num7;
						num10 = 1;
						goto IL_0547;
						IL_0547:
						if (num10 <= num9)
						{
							goto IL_02be;
						}
						goto IL_0550;
						IL_02be:
						num2 = 25;
						num11 = num6 * num10 + num5 * (num10 - 1);
						goto IL_02cf;
						IL_02cf:
						num2 = 26;
						range4 = range2.get_Offset((object)num11, (object)0);
						goto IL_02ea;
						IL_02ea:
						num2 = 27;
						range3 = null;
						goto IL_02f0;
						IL_02f0:
						num2 = 28;
						range3 = application.Intersect(this.m_A, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0444;
						IL_0444:
						num2 = 29;
						if (range3 != null)
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
							goto IL_0458;
						}
						goto IL_0550;
						IL_0ddc:
						num2 = 95;
						if (optNoChange.IsChecked == true)
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
							goto IL_0e05;
						}
						goto IL_0e19;
						IL_0458:
						num2 = 31;
						A(range3).EntireRow.Insert(XlInsertShiftDirection.xlShiftDown, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0488;
						IL_0488:
						num2 = 32;
						if (value)
						{
							goto IL_048f;
						}
						goto IL_04de;
						IL_048f:
						num2 = 33;
						range4 = range.get_Offset((object)num11, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_04ae;
						IL_04ae:
						num2 = 34;
						ExcelAddIn1.Format.Styles.ApplyStyle(A(range4), this.m_A[cbxStyles.SelectedIndex]);
						goto IL_0514;
						IL_04de:
						num2 = 36;
						if (value2)
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
							goto IL_04ef;
						}
						goto IL_0514;
						IL_0f06:
						num2 = 112;
						xmlDocument = null;
						break;
						IL_04ef:
						num2 = 37;
						range.get_Offset((object)num11, RuntimeHelpers.GetObjectValue(Missing.Value)).ClearFormats();
						goto IL_0514;
						IL_0514:
						num2 = 38;
						if (value3)
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
							goto IL_0525;
						}
						goto IL_053e;
						IL_0e05:
						num2 = 96;
						xmlAttribute2.Value = Conversions.ToString(0);
						goto IL_0e2b;
						IL_0525:
						num2 = 39;
						A(range4).RowHeight = num8;
						goto IL_053e;
						IL_053e:
						num2 = 40;
						num10++;
						goto IL_0547;
						IL_0550:
						num2 = 41;
						this.m_A.Select();
						goto IL_0b1a;
						IL_0564:
						num2 = 43;
						if (optDelete.IsChecked == true)
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
							goto IL_058c;
						}
						goto IL_0786;
						IL_0e19:
						num2 = 98;
						xmlAttribute2.Value = Conversions.ToString(1);
						goto IL_0e2b;
						IL_058c:
						num2 = 44;
						num12 = num7;
						num13 = 1;
						goto IL_075d;
						IL_075d:
						if (num13 <= num12)
						{
							goto IL_059b;
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
						goto IL_0770;
						IL_0e2b:
						xmlAttribute2 = null;
						goto IL_0e2e;
						IL_059b:
						num2 = 45;
						range4 = range2.get_Offset((object)((num6 - num5) * num13), (object)0);
						goto IL_05bb;
						IL_05bb:
						num2 = 46;
						range3 = null;
						goto IL_05c1;
						IL_05c1:
						num2 = 47;
						range3 = application.Intersect(this.m_A, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_071b;
						IL_071b:
						num2 = 48;
						if (range3 != null)
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
							goto IL_072c;
						}
						goto IL_0770;
						IL_0e2e:
						num2 = 100;
						xmlNode.Attributes[VH.A(206974)].Value = numRowHeight.Value.Value.ToString(CultureInfo.InvariantCulture);
						goto IL_0e7c;
						IL_072c:
						num2 = 50;
						A(range4).EntireRow.Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0754;
						IL_0754:
						num2 = 51;
						num13++;
						goto IL_075d;
						IL_0770:
						num2 = 52;
						this.m_A.Select();
						goto IL_0b1a;
						IL_0786:
						num2 = 54;
						num14 = num7;
						num15 = 1;
						goto IL_0b07;
						IL_0b07:
						if (num15 <= num14)
						{
							goto IL_0795;
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
						goto IL_0b1a;
						IL_0e7c:
						num2 = 101;
						xmlNode.Attributes[VH.A(206983)].Value = Conversions.ToString(chkSave.IsChecked.Value);
						goto IL_0ebd;
						IL_0795:
						num2 = 55;
						range4 = range.get_Offset((object)(num6 * num15 - num5), (object)0);
						goto IL_07b3;
						IL_07b3:
						num2 = 56;
						range3 = null;
						goto IL_07b9;
						IL_07b9:
						num2 = 57;
						range3 = application.Intersect(this.m_A, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0911;
						IL_0911:
						num2 = 58;
						if (range3 != null)
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
							goto IL_0925;
						}
						goto IL_0b1a;
						IL_0ebd:
						xmlNode = null;
						goto IL_0ec0;
						IL_0925:
						num2 = 60;
						range4 = application.Intersect(A(range3), this.m_A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0a85;
						IL_0a85:
						num2 = 61;
						if (value)
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
							goto IL_0a96;
						}
						goto IL_0abe;
						IL_0ec0:
						num2 = 103;
						KH.A.SaveSettings(xmlDocument);
						goto IL_0ecf;
						IL_0a96:
						num2 = 62;
						ExcelAddIn1.Format.Styles.ApplyStyle(range4, this.m_A[cbxStyles.SelectedIndex]);
						goto IL_0adc;
						IL_0abe:
						num2 = 64;
						if (value2)
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
							goto IL_0acf;
						}
						goto IL_0adc;
						IL_0ecf:
						num2 = 104;
						application.ScreenUpdating = true;
						goto IL_0eda;
						IL_0acf:
						num2 = 65;
						range4.ClearFormats();
						goto IL_0adc;
						IL_0adc:
						num2 = 66;
						if (value3)
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
							goto IL_0aed;
						}
						goto IL_0afe;
						IL_0eda:
						num2 = 105;
						application.EnableEvents = true;
						goto IL_0ee5;
						IL_0aed:
						num2 = 67;
						range4.RowHeight = num8;
						goto IL_0afe;
						IL_0afe:
						num2 = 68;
						num15++;
						goto IL_0b07;
						IL_0b1a:
						num2 = 69;
						if (chkSave.IsChecked == true)
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
							goto IL_0b44;
						}
						goto IL_0ecf;
						IL_0ee5:
						num2 = 106;
						application = null;
						goto IL_0eeb;
						IL_0b44:
						num2 = 70;
						xmlNode = xmlDocument.DocumentElement.SelectSingleNode(VH.A(207397));
						goto IL_0b63;
						IL_0b63:
						num2 = 71;
						xmlAttribute3 = xmlNode.Attributes[VH.A(206896)];
						goto IL_0b84;
						IL_0b84:
						num2 = 72;
						if (optFormat.IsChecked == true)
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
							goto IL_0bad;
						}
						goto IL_0bc1;
						IL_0eeb:
						num2 = 107;
						goto IL_0eee;
						IL_0bad:
						num2 = 73;
						xmlAttribute3.Value = Conversions.ToString(0);
						goto IL_0c2b;
						IL_0bc1:
						num2 = 75;
						if (optInsert.IsChecked == true)
						{
							goto IL_0bde;
						}
						goto IL_0bf2;
						IL_0bde:
						num2 = 76;
						xmlAttribute3.Value = Conversions.ToString(1);
						goto IL_0c2b;
						IL_0bf2:
						num2 = 78;
						if (optDelete.IsChecked == true)
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
							goto IL_0c1b;
						}
						goto IL_0c2b;
						IL_0eee:
						num2 = 108;
						range4 = null;
						goto IL_0ef4;
						IL_0c1b:
						num2 = 79;
						xmlAttribute3.Value = Conversions.ToString(2);
						goto IL_0c2b;
						IL_0c2b:
						xmlAttribute3 = null;
						goto IL_0c2e;
						IL_0c2e:
						num2 = 81;
						xmlNode.Attributes[VH.A(206909)].Value = Conversions.ToString(numInsert.Value.Value);
						goto IL_0c71;
						end_IL_0000_2:
						break;
					}
					num2 = 113;
					base.DialogResult = true;
					break;
				}
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 4344;
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

	private Range A(Range A)
	{
		return ((_Worksheet)A.Worksheet).get_Range((object)A, (object)A.get_Offset((object)(numInsert.Value - 1.0), RuntimeHelpers.GetObjectValue(Missing.Value)));
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = false;
		Close();
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (B)
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
			B = true;
			Uri resourceLocator = new Uri(VH.A(207597), UriKind.Relative);
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
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		//IL_004f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Expected O, but got Unknown
		//IL_00b1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bb: Expected O, but got Unknown
		//IL_0116: Unknown result type (might be due to invalid IL or missing references)
		//IL_0120: Expected O, but got Unknown
		//IL_022e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0238: Expected O, but got Unknown
		if (connectionId == 1)
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
					optFormat = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					lbl1 = (TextBlock)target;
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
					numInsert = (MacNumericUpDown)target;
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
					lbl2 = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					optInsert = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					lbl3 = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			numEvery = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					lbl4 = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					optDelete = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					lbl5 = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			numRows = (MacNumericUpDown)target;
			return;
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
					lbl6 = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					grpFormat = (GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					optInherit = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					optStyle = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					cbxStyles = (ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					optClear = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					grpHeight = (GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					optNoChange = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					optSetTo = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					numRowHeight = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkSave = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 23:
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				btnCancel = (Button)target;
				return;
			}
		case 24:
			btnOk = (Button)target;
			break;
		default:
			B = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
