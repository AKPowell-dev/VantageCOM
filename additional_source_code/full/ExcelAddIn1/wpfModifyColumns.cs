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
public sealed class wpfModifyColumns : System.Windows.Window, IComponentConnector
{
	private Microsoft.Office.Interop.Excel.Application m_A;

	private Range m_A;

	private bool m_A;

	private List<XmlNode> m_A;

	[AccessedThroughProperty("optFormat")]
	[CompilerGenerated]
	private RadioButton m_A;

	[AccessedThroughProperty("lbl1")]
	[CompilerGenerated]
	private TextBlock m_A;

	[AccessedThroughProperty("numInsert")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("lbl2")]
	[CompilerGenerated]
	private TextBlock B;

	[CompilerGenerated]
	[AccessedThroughProperty("optInsert")]
	private RadioButton B;

	[AccessedThroughProperty("lbl3")]
	[CompilerGenerated]
	private TextBlock C;

	[AccessedThroughProperty("numEvery")]
	[CompilerGenerated]
	private MacNumericUpDown B;

	[AccessedThroughProperty("lbl4")]
	[CompilerGenerated]
	private TextBlock D;

	[AccessedThroughProperty("optDelete")]
	[CompilerGenerated]
	private RadioButton C;

	[CompilerGenerated]
	[AccessedThroughProperty("lbl5")]
	private TextBlock E;

	[CompilerGenerated]
	[AccessedThroughProperty("numCols")]
	private MacNumericUpDown C;

	[AccessedThroughProperty("lbl6")]
	[CompilerGenerated]
	private TextBlock F;

	[AccessedThroughProperty("grpFormat")]
	[CompilerGenerated]
	private GroupBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("optInherit")]
	private RadioButton D;

	[AccessedThroughProperty("optStyle")]
	[CompilerGenerated]
	private RadioButton E;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxStyles")]
	private ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("optClear")]
	private RadioButton F;

	[CompilerGenerated]
	[AccessedThroughProperty("grpWidth")]
	private GroupBox B;

	[AccessedThroughProperty("optNoChange")]
	[CompilerGenerated]
	private RadioButton G;

	[CompilerGenerated]
	[AccessedThroughProperty("optSetTo")]
	private RadioButton H;

	[AccessedThroughProperty("numColWidth")]
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
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
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

	internal virtual MacNumericUpDown numCols
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
			MacRangeBaseValueChangedHandler val = new MacRangeBaseValueChangedHandler(numCols_ValueChanged);
			MacNumericUpDown val2 = C;
			if (val2 != null)
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
				val2.ValueChanged -= val;
			}
			C = value;
			val2 = C;
			if (val2 != null)
			{
				val2.ValueChanged += val;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				radioButton.Checked += value2;
				radioButton.Unchecked += value3;
				return;
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

	internal virtual GroupBox grpWidth
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
				radioButton.Checked -= value2;
				radioButton.Unchecked -= value3;
			}
			H = value;
			radioButton = H;
			if (radioButton != null)
			{
				radioButton.Checked += value2;
				radioButton.Unchecked += value3;
			}
		}
	}

	internal virtual MacNumericUpDown numColWidth
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
			this.m_A = value;
			button = this.m_A;
			if (button == null)
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
				button.Click += value2;
				return;
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
				button.Click -= value2;
			}
			this.B = value;
			button = this.B;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	public wpfModifyColumns()
	{
		base.Loaded += wpfModifyColumns_Loaded;
		base.Closing += wpfModifyColumns_Closing;
		this.m_A = true;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = MH.A.Application;
		this.m_A = (Range)this.m_A.Selection;
	}

	private void wpfModifyColumns_Loaded(object sender, RoutedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		XmlNode xmlNode = default(XmlNode);
		List<string> list = default(List<string>);
		Dictionary<string, List<XmlNode>>.Enumerator enumerator = default(Dictionary<string, List<XmlNode>>.Enumerator);
		KeyValuePair<string, List<XmlNode>> current = default(KeyValuePair<string, List<XmlNode>>);
		string value = default(string);
		string key = default(string);
		List<XmlNode>.Enumerator enumerator2 = default(List<XmlNode>.Enumerator);
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
				case 1542:
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
							goto IL_006b;
						case 7:
							goto IL_0078;
						case 8:
							goto IL_0095;
						case 9:
							goto IL_00ca;
						case 10:
							goto IL_00da;
						case 11:
							goto IL_00e8;
						case 12:
							goto IL_00f8;
						case 13:
							goto IL_011c;
						case 14:
							goto IL_012c;
						case 15:
							goto IL_013b;
						case 16:
							goto IL_0140;
						case 17:
							goto IL_015c;
						case 18:
							goto IL_0176;
						case 19:
							goto IL_0192;
						case 20:
							goto IL_01bb;
						case 22:
							goto IL_01df;
						case 23:
							goto IL_01fc;
						case 25:
							goto IL_0214;
						case 26:
							goto IL_0235;
						case 28:
							goto IL_024d;
						case 29:
							goto IL_026c;
						case 21:
						case 24:
						case 27:
						case 30:
							goto IL_0280;
						case 31:
							goto IL_02bb;
						case 32:
							goto IL_02f8;
						case 34:
							goto IL_031e;
						case 35:
							goto IL_033b;
						case 37:
							goto IL_0353;
						case 38:
							goto IL_0372;
						case 40:
							goto IL_0388;
						case 41:
							goto IL_03a7;
						case 33:
						case 36:
						case 39:
						case 42:
							goto IL_03bd;
						case 43:
							goto IL_03f1;
						case 45:
							goto IL_0419;
						case 46:
							goto IL_043a;
						case 48:
							goto IL_0452;
						case 49:
							goto IL_0471;
						case 44:
						case 47:
						case 50:
							goto IL_0485;
						case 51:
							goto IL_04c8;
						case 52:
							goto IL_0505;
						case 53:
							goto IL_0508;
						case 54:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 55:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0485:
					num2 = 50;
					numColWidth.Value = float.Parse(xmlNode.Attributes[VH.A(206974)].Value, CultureInfo.InvariantCulture);
					goto IL_04c8;
					IL_0007:
					num2 = 2;
					list = new List<string>();
					goto IL_000f;
					IL_000f:
					num2 = 3;
					numCols.Value = (double?)this.m_A.Columns.CountLarge;
					goto IL_0037;
					IL_0037:
					num2 = 4;
					this.m_A = new List<XmlNode>();
					goto IL_0044;
					IL_0044:
					num2 = 5;
					enumerator = KH.A.CustomCycles.GetEnumerator();
					goto IL_00fb;
					IL_00fb:
					if (enumerator.MoveNext())
					{
						current = enumerator.Current;
						goto IL_006b;
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
					goto IL_011c;
					IL_04c8:
					num2 = 51;
					chkSave.IsChecked = Conversions.ToBoolean(xmlNode.Attributes[VH.A(206983)].Value);
					goto IL_0505;
					IL_0508:
					num2 = 53;
					break;
					IL_0505:
					xmlNode = null;
					goto IL_0508;
					IL_011c:
					num2 = 13;
					((IDisposable)enumerator/*cast due to .constrained prefix*/).Dispose();
					goto IL_012c;
					IL_012c:
					num2 = 14;
					cbxStyles.ItemsSource = list;
					goto IL_013b;
					IL_013b:
					num2 = 15;
					list = null;
					goto IL_0140;
					IL_0140:
					num2 = 16;
					optFormat.Checked += UpdateLabels;
					goto IL_015c;
					IL_015c:
					num2 = 17;
					optInsert.Checked += UpdateLabels;
					goto IL_0176;
					IL_0176:
					num2 = 18;
					optDelete.Checked += UpdateLabels;
					goto IL_0192;
					IL_0192:
					num2 = 19;
					xmlNode = KH.A.SettingsXml.DocumentElement.SelectSingleNode(VH.A(206869));
					goto IL_01bb;
					IL_01bb:
					num2 = 20;
					value = xmlNode.Attributes[VH.A(206896)].Value;
					goto IL_01df;
					IL_01df:
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
						goto IL_01fc;
					}
					goto IL_0214;
					IL_006b:
					num2 = 6;
					key = current.Key;
					goto IL_0078;
					IL_01fc:
					num2 = 23;
					optFormat.IsChecked = true;
					goto IL_0280;
					IL_0214:
					num2 = 25;
					if (Operators.CompareString(value, Conversions.ToString(1), TextCompare: false) == 0)
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
						goto IL_0235;
					}
					goto IL_024d;
					IL_0078:
					num2 = 7;
					enumerator2 = current.Value.GetEnumerator();
					goto IL_00dd;
					IL_0235:
					num2 = 26;
					optInsert.IsChecked = true;
					goto IL_0280;
					IL_024d:
					num2 = 28;
					if (Operators.CompareString(value, Conversions.ToString(2), TextCompare: false) == 0)
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
						goto IL_026c;
					}
					goto IL_0280;
					IL_00dd:
					if (enumerator2.MoveNext())
					{
						current2 = enumerator2.Current;
						goto IL_0095;
					}
					goto IL_00e8;
					IL_026c:
					num2 = 29;
					optDelete.IsChecked = true;
					goto IL_0280;
					IL_0280:
					num2 = 30;
					numInsert.Value = Conversions.ToDouble(xmlNode.Attributes[VH.A(206909)].Value);
					goto IL_02bb;
					IL_02bb:
					num2 = 31;
					numEvery.Value = Conversions.ToDouble(xmlNode.Attributes[VH.A(206924)].Value);
					goto IL_02f8;
					IL_02f8:
					num2 = 32;
					value2 = xmlNode.Attributes[VH.A(206939)].Value;
					goto IL_031e;
					IL_031e:
					num2 = 34;
					if (Operators.CompareString(value2, Conversions.ToString(0), TextCompare: false) == 0)
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
						goto IL_033b;
					}
					goto IL_0353;
					IL_00e8:
					num2 = 11;
					((IDisposable)enumerator2/*cast due to .constrained prefix*/).Dispose();
					goto IL_00f8;
					IL_033b:
					num2 = 35;
					optInherit.IsChecked = true;
					goto IL_03bd;
					IL_0353:
					num2 = 37;
					if (Operators.CompareString(value2, Conversions.ToString(1), TextCompare: false) == 0)
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
						goto IL_0372;
					}
					goto IL_0388;
					IL_00f8:
					num2 = 12;
					goto IL_00fb;
					IL_0372:
					num2 = 38;
					optStyle.IsChecked = true;
					goto IL_03bd;
					IL_0388:
					num2 = 40;
					if (Operators.CompareString(value2, Conversions.ToString(2), TextCompare: false) == 0)
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
						goto IL_03a7;
					}
					goto IL_03bd;
					IL_00da:
					num2 = 10;
					goto IL_00dd;
					IL_03a7:
					num2 = 41;
					optClear.IsChecked = true;
					goto IL_03bd;
					IL_03bd:
					num2 = 42;
					cbxStyles.SelectedIndex = Conversions.ToInteger(xmlNode.Attributes[VH.A(206952)].Value);
					goto IL_03f1;
					IL_03f1:
					num2 = 43;
					value3 = xmlNode.Attributes[VH.A(206963)].Value;
					goto IL_0419;
					IL_0419:
					num2 = 45;
					if (Operators.CompareString(value3, Conversions.ToString(0), TextCompare: false) == 0)
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
						goto IL_043a;
					}
					goto IL_0452;
					IL_0095:
					num2 = 8;
					list.Add(key + VH.A(206862) + current2.SelectSingleNode(VH.A(19019)).InnerText);
					goto IL_00ca;
					IL_043a:
					num2 = 46;
					optNoChange.IsChecked = true;
					goto IL_0485;
					IL_0452:
					num2 = 48;
					if (Operators.CompareString(value3, Conversions.ToString(1), TextCompare: false) == 0)
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
						goto IL_0471;
					}
					goto IL_0485;
					IL_00ca:
					num2 = 9;
					this.m_A.Add(current2);
					goto IL_00da;
					IL_0471:
					num2 = 49;
					optSetTo.IsChecked = true;
					goto IL_0485;
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
				try0000_dispatch = 1542;
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
					switch (3)
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
					switch (4)
					{
					case 0:
						continue;
					}
					break;
				}
				lbl1.Text = VH.A(57289);
				lbl2.Text = VH.A(2862);
			}
			else if (Operators.CompareString(name, optDelete.Name, TextCompare: false) == 0)
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
				lbl1.Text = VH.A(207036);
				lbl2.Text = VH.A(207031);
			}
		}
		radioButton = null;
	}

	private void numCols_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
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
				Worksheet worksheet;
				object objectValue;
				object instance;
				string memberName;
				object[] obj;
				double? value;
				double? obj2;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 368:
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
							goto IL_0104;
						case 7:
							goto IL_0106;
						case 8:
							goto IL_0116;
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
					worksheet = range.Worksheet;
					objectValue = RuntimeHelpers.GetObjectValue(range.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)]);
					instance = range.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
					memberName = VH.A(60565);
					obj = new object[2]
					{
						Missing.Value,
						null
					};
					value = numCols.Value;
					if (!value.HasValue)
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
						obj2 = null;
					}
					else
					{
						obj2 = value.GetValueOrDefault() - 1.0;
					}
					obj[1] = obj2;
					this.m_A = ((_Worksheet)worksheet).get_Range(objectValue, RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(instance, null, memberName, obj, null, null, null)));
					goto IL_0104;
					IL_0007:
					num2 = 2;
					if (this.m_A)
					{
						goto end_IL_0000_3;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_0027;
					IL_0104:
					range = null;
					goto IL_0106;
					IL_0116:
					num2 = 8;
					this.m_A.ScreenUpdating = false;
					break;
					IL_0106:
					num2 = 7;
					this.m_A.Select();
					goto IL_0116;
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
			catch (object obj3) when (obj3 is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj3);
				try0000_dispatch = 368;
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
		((UIElement)(object)numColWidth).IsEnabled = optSetTo.IsChecked.Value;
	}

	private void optDelete_CheckedChanged(object sender, RoutedEventArgs e)
	{
		bool? isChecked = optDelete.IsChecked;
		bool? flag;
		if (!isChecked.HasValue)
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
			flag = isChecked;
		}
		else
		{
			flag = isChecked != true;
		}
		isChecked = flag;
		bool value = isChecked.Value;
		grpFormat.IsEnabled = value;
		grpWidth.IsEnabled = value;
	}

	private void wpfModifyColumns_Closing(object sender, CancelEventArgs e)
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
		Range range = default(Range);
		int num5 = default(int);
		int num6 = default(int);
		int num7 = default(int);
		bool value = default(bool);
		bool value2 = default(bool);
		bool value3 = default(bool);
		float num8 = default(float);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		Range range2 = default(Range);
		XmlDocument xmlDocument = default(XmlDocument);
		XmlAttribute xmlAttribute = default(XmlAttribute);
		Range range3 = default(Range);
		XmlNode xmlNode = default(XmlNode);
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
					case 4284:
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
								goto IL_0028;
							case 4:
								goto IL_004c;
							case 5:
								goto IL_0070;
							case 6:
								goto IL_008c;
							case 7:
								goto IL_00aa;
							case 8:
								goto IL_00c8;
							case 9:
								goto IL_00e7;
							case 10:
								goto IL_00f8;
							case 11:
								goto IL_0107;
							case 12:
								goto IL_0137;
							case 13:
								goto IL_0167;
							case 14:
								goto IL_0178;
							case 15:
								goto IL_01c4;
							case 16:
								goto IL_01cc;
							case 18:
								goto IL_01e5;
							case 19:
								goto IL_024e;
							case 21:
								goto IL_0267;
							case 22:
								goto IL_0272;
							case 23:
								goto IL_027d;
							case 24:
								goto IL_02a9;
							case 25:
								goto IL_02b8;
							case 26:
								goto IL_02c9;
							case 27:
								goto IL_02e4;
							case 28:
								goto IL_02ea;
							case 29:
								goto IL_043e;
							case 31:
								goto IL_0452;
							case 32:
								goto IL_0482;
							case 33:
								goto IL_0489;
							case 34:
								goto IL_04aa;
							case 36:
								goto IL_04da;
							case 37:
								goto IL_04eb;
							case 35:
							case 38:
								goto IL_0512;
							case 39:
								goto IL_0523;
							case 40:
								goto IL_053a;
							case 30:
							case 41:
								goto IL_0556;
							case 43:
								goto IL_056c;
							case 44:
								goto IL_0598;
							case 45:
								goto IL_05a7;
							case 46:
								goto IL_05c7;
							case 47:
								goto IL_05cd;
							case 48:
								goto IL_0729;
							case 50:
								goto IL_073a;
							case 51:
								goto IL_0760;
							case 49:
							case 52:
								goto IL_077c;
							case 54:
								goto IL_0790;
							case 55:
								goto IL_079f;
							case 56:
								goto IL_07bf;
							case 57:
								goto IL_07c5;
							case 58:
								goto IL_0919;
							case 60:
								goto IL_0923;
							case 61:
								goto IL_0a7b;
							case 62:
								goto IL_0a8c;
							case 64:
								goto IL_0ab0;
							case 65:
								goto IL_0ac1;
							case 63:
							case 66:
								goto IL_0ace;
							case 67:
								goto IL_0ad5;
							case 68:
								goto IL_0ae6;
							case 42:
							case 53:
							case 59:
							case 69:
								goto IL_0b02;
							case 70:
								goto IL_0b2a;
							case 71:
								goto IL_0b49;
							case 72:
								goto IL_0b66;
							case 73:
								goto IL_0b85;
							case 75:
								goto IL_0b99;
							case 76:
								goto IL_0bb6;
							case 78:
								goto IL_0bca;
							case 79:
								goto IL_0bf3;
							case 74:
							case 77:
							case 80:
								goto IL_0c05;
							case 81:
								goto IL_0c08;
							case 82:
								goto IL_0c4b;
							case 83:
								goto IL_0c8a;
							case 84:
								goto IL_0cab;
							case 85:
								goto IL_0cd2;
							case 87:
								goto IL_0ce6;
							case 88:
								goto IL_0d0b;
							case 90:
								goto IL_0d1f;
							case 91:
								goto IL_0d42;
							case 86:
							case 89:
							case 92:
								goto IL_0d52;
							case 93:
								goto IL_0d55;
							case 94:
								goto IL_0d8d;
							case 95:
								goto IL_0dac;
							case 96:
								goto IL_0dd1;
							case 98:
								goto IL_0de3;
							case 97:
							case 99:
								goto IL_0df3;
							case 100:
								goto IL_0df6;
							case 101:
								goto IL_0e42;
							case 102:
								goto IL_0e81;
							case 103:
								goto IL_0e84;
							case 104:
								goto IL_0e93;
							case 105:
								goto IL_0e9e;
							case 106:
								goto IL_0ea9;
							case 107:
								goto IL_0eaf;
							case 108:
								goto IL_0eb2;
							case 109:
								goto IL_0eb8;
							case 110:
								goto IL_0ebe;
							case 111:
								goto IL_0ec4;
							case 112:
								goto IL_0eca;
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
						IL_0eb8:
						num2 = 109;
						range = null;
						goto IL_0ebe;
						IL_0007:
						num2 = 2;
						num5 = (int)Math.Round(numInsert.Value.Value);
						goto IL_0028;
						IL_0028:
						num2 = 3;
						num6 = (int)Math.Round(numEvery.Value.Value);
						goto IL_004c;
						IL_004c:
						num2 = 4;
						num7 = (int)Math.Round(numCols.Value.Value);
						goto IL_0070;
						IL_0070:
						num2 = 5;
						value = optStyle.IsChecked.Value;
						goto IL_008c;
						IL_008c:
						num2 = 6;
						value2 = optClear.IsChecked.Value;
						goto IL_00aa;
						IL_00aa:
						num2 = 7;
						value3 = optSetTo.IsChecked.Value;
						goto IL_00c8;
						IL_00c8:
						num2 = 8;
						num8 = (float)numColWidth.Value.Value;
						goto IL_00e7;
						IL_00e7:
						num2 = 9;
						application = MH.A.Application;
						goto IL_00f8;
						IL_00f8:
						num2 = 10;
						_ = this.m_A.Worksheet;
						goto IL_0107;
						IL_0107:
						num2 = 11;
						range2 = (Range)this.m_A.Columns[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
						goto IL_0137;
						IL_0137:
						num2 = 12;
						range = (Range)this.m_A.Cells[1, RuntimeHelpers.GetObjectValue(Missing.Value)];
						goto IL_0167;
						IL_0167:
						num2 = 13;
						xmlDocument = KH.A.SettingsXml;
						goto IL_0178;
						IL_0178:
						num2 = 14;
						if (optFormat.IsChecked != true)
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
							if (optDelete.IsChecked != true)
							{
								goto IL_01e5;
							}
						}
						goto IL_01c4;
						IL_0cd2:
						num2 = 85;
						xmlAttribute.Value = Conversions.ToString(0);
						goto IL_0d52;
						IL_0ce6:
						num2 = 87;
						if (optStyle.IsChecked == true)
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
							goto IL_0d0b;
						}
						goto IL_0d1f;
						IL_0d0b:
						num2 = 88;
						xmlAttribute.Value = Conversions.ToString(1);
						goto IL_0d52;
						IL_0ebe:
						num2 = 110;
						range3 = null;
						goto IL_0ec4;
						IL_01c4:
						num2 = 15;
						if (num5 > num6)
						{
							goto IL_01cc;
						}
						goto IL_01e5;
						IL_01cc:
						num2 = 16;
						Forms.WarningMessage(VH.A(207067));
						goto end_IL_0000_3;
						IL_01e5:
						num2 = 18;
						isChecked = optStyle.IsChecked;
						if (isChecked.HasValue)
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
							if (isChecked != true)
							{
								goto IL_0267;
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
						if (cbxStyles.SelectedIndex == -1)
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
								goto IL_024e;
							}
						}
						goto IL_0267;
						IL_0ec4:
						num2 = 111;
						range2 = null;
						goto IL_0eca;
						IL_0d42:
						num2 = 91;
						xmlAttribute.Value = Conversions.ToString(2);
						goto IL_0d52;
						IL_0d52:
						xmlAttribute = null;
						goto IL_0d55;
						IL_0d1f:
						num2 = 90;
						if (optClear.IsChecked == true)
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
							goto IL_0d42;
						}
						goto IL_0d52;
						IL_0d55:
						num2 = 93;
						xmlNode.Attributes[VH.A(206952)].Value = Conversions.ToString(cbxStyles.SelectedIndex);
						goto IL_0d8d;
						IL_0d8d:
						num2 = 94;
						xmlAttribute2 = xmlNode.Attributes[VH.A(206963)];
						goto IL_0dac;
						IL_0dac:
						num2 = 95;
						if (optNoChange.IsChecked == true)
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
							goto IL_0dd1;
						}
						goto IL_0de3;
						IL_024e:
						num2 = 19;
						Forms.WarningMessage(VH.A(207239));
						goto end_IL_0000_3;
						IL_0267:
						num2 = 21;
						application.ScreenUpdating = false;
						goto IL_0272;
						IL_0272:
						num2 = 22;
						application.EnableEvents = false;
						goto IL_027d;
						IL_027d:
						num2 = 23;
						if (optInsert.IsChecked == true)
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
							goto IL_02a9;
						}
						goto IL_056c;
						IL_0eca:
						num2 = 112;
						xmlDocument = null;
						break;
						IL_02a9:
						num2 = 24;
						num9 = num7;
						num10 = 1;
						goto IL_0543;
						IL_0543:
						if (num10 <= num9)
						{
							goto IL_02b8;
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
						goto IL_0556;
						IL_0dd1:
						num2 = 96;
						xmlAttribute2.Value = Conversions.ToString(0);
						goto IL_0df3;
						IL_02b8:
						num2 = 25;
						num11 = num6 * num10 + num5 * (num10 - 1);
						goto IL_02c9;
						IL_02c9:
						num2 = 26;
						range4 = range.get_Offset((object)0, (object)num11);
						goto IL_02e4;
						IL_02e4:
						num2 = 27;
						range3 = null;
						goto IL_02ea;
						IL_02ea:
						num2 = 28;
						range3 = application.Intersect(this.m_A, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_043e;
						IL_043e:
						num2 = 29;
						if (range3 != null)
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
							goto IL_0452;
						}
						goto IL_0556;
						IL_0de3:
						num2 = 98;
						xmlAttribute2.Value = Conversions.ToString(1);
						goto IL_0df3;
						IL_0452:
						num2 = 31;
						A(range3).EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0482;
						IL_0482:
						num2 = 32;
						if (value)
						{
							goto IL_0489;
						}
						goto IL_04da;
						IL_0489:
						num2 = 33;
						range4 = range2.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)num11);
						goto IL_04aa;
						IL_04aa:
						num2 = 34;
						ExcelAddIn1.Format.Styles.ApplyStyle(A(range4), this.m_A[cbxStyles.SelectedIndex]);
						goto IL_0512;
						IL_04da:
						num2 = 36;
						if (value2)
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
							goto IL_04eb;
						}
						goto IL_0512;
						IL_0df3:
						xmlAttribute2 = null;
						goto IL_0df6;
						IL_04eb:
						num2 = 37;
						range2.get_Offset(RuntimeHelpers.GetObjectValue(Missing.Value), (object)num11).ClearFormats();
						goto IL_0512;
						IL_0512:
						num2 = 38;
						if (value3)
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
							goto IL_0523;
						}
						goto IL_053a;
						IL_0df6:
						num2 = 100;
						xmlNode.Attributes[VH.A(206974)].Value = numColWidth.Value.Value.ToString(CultureInfo.InvariantCulture);
						goto IL_0e42;
						IL_0523:
						num2 = 39;
						A(range4).ColumnWidth = num8;
						goto IL_053a;
						IL_053a:
						num2 = 40;
						num10++;
						goto IL_0543;
						IL_0556:
						num2 = 41;
						this.m_A.Select();
						goto IL_0b02;
						IL_056c:
						num2 = 43;
						if (optDelete.IsChecked == true)
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
							goto IL_0598;
						}
						goto IL_0790;
						IL_0e42:
						num2 = 101;
						xmlNode.Attributes[VH.A(206983)].Value = Conversions.ToString(chkSave.IsChecked.Value);
						goto IL_0e81;
						IL_0598:
						num2 = 44;
						num12 = num7;
						num13 = 1;
						goto IL_0769;
						IL_0769:
						if (num13 <= num12)
						{
							goto IL_05a7;
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
						goto IL_077c;
						IL_0e81:
						xmlNode = null;
						goto IL_0e84;
						IL_05a7:
						num2 = 45;
						range4 = range.get_Offset((object)0, (object)((num6 - num5) * num13));
						goto IL_05c7;
						IL_05c7:
						num2 = 46;
						range3 = null;
						goto IL_05cd;
						IL_05cd:
						num2 = 47;
						range3 = application.Intersect(this.m_A, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0729;
						IL_0729:
						num2 = 48;
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
							goto IL_073a;
						}
						goto IL_077c;
						IL_0e84:
						num2 = 103;
						KH.A.SaveSettings(xmlDocument);
						goto IL_0e93;
						IL_073a:
						num2 = 50;
						A(range4).EntireColumn.Delete(RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0760;
						IL_0760:
						num2 = 51;
						num13++;
						goto IL_0769;
						IL_077c:
						num2 = 52;
						this.m_A.Select();
						goto IL_0b02;
						IL_0790:
						num2 = 54;
						num14 = num7;
						num15 = 1;
						goto IL_0aef;
						IL_0aef:
						if (num15 <= num14)
						{
							goto IL_079f;
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
						goto IL_0b02;
						IL_0e93:
						num2 = 104;
						application.ScreenUpdating = true;
						goto IL_0e9e;
						IL_079f:
						num2 = 55;
						range4 = range2.get_Offset((object)0, (object)(num6 * num15 - num5));
						goto IL_07bf;
						IL_07bf:
						num2 = 56;
						range3 = null;
						goto IL_07c5;
						IL_07c5:
						num2 = 57;
						range3 = application.Intersect(this.m_A, range4, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0919;
						IL_0919:
						num2 = 58;
						if (range3 != null)
						{
							goto IL_0923;
						}
						goto IL_0b02;
						IL_0923:
						num2 = 60;
						range4 = application.Intersect(A(range3), this.m_A, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						goto IL_0a7b;
						IL_0a7b:
						num2 = 61;
						if (value)
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
							goto IL_0a8c;
						}
						goto IL_0ab0;
						IL_0e9e:
						num2 = 105;
						application.EnableEvents = true;
						goto IL_0ea9;
						IL_0a8c:
						num2 = 62;
						ExcelAddIn1.Format.Styles.ApplyStyle(range4, this.m_A[cbxStyles.SelectedIndex]);
						goto IL_0ace;
						IL_0ab0:
						num2 = 64;
						if (value2)
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
							goto IL_0ac1;
						}
						goto IL_0ace;
						IL_0ea9:
						num2 = 106;
						application = null;
						goto IL_0eaf;
						IL_0ac1:
						num2 = 65;
						range4.ClearFormats();
						goto IL_0ace;
						IL_0ace:
						num2 = 66;
						if (value3)
						{
							goto IL_0ad5;
						}
						goto IL_0ae6;
						IL_0ad5:
						num2 = 67;
						range4.ColumnWidth = num8;
						goto IL_0ae6;
						IL_0ae6:
						num2 = 68;
						num15++;
						goto IL_0aef;
						IL_0b02:
						num2 = 69;
						if (chkSave.IsChecked == true)
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
							goto IL_0b2a;
						}
						goto IL_0e93;
						IL_0eaf:
						num2 = 107;
						goto IL_0eb2;
						IL_0b2a:
						num2 = 70;
						xmlNode = xmlDocument.DocumentElement.SelectSingleNode(VH.A(206869));
						goto IL_0b49;
						IL_0b49:
						num2 = 71;
						xmlAttribute3 = xmlNode.Attributes[VH.A(206896)];
						goto IL_0b66;
						IL_0b66:
						num2 = 72;
						if (optFormat.IsChecked == true)
						{
							goto IL_0b85;
						}
						goto IL_0b99;
						IL_0b85:
						num2 = 73;
						xmlAttribute3.Value = Conversions.ToString(0);
						goto IL_0c05;
						IL_0b99:
						num2 = 75;
						if (optInsert.IsChecked == true)
						{
							goto IL_0bb6;
						}
						goto IL_0bca;
						IL_0bb6:
						num2 = 76;
						xmlAttribute3.Value = Conversions.ToString(1);
						goto IL_0c05;
						IL_0bca:
						num2 = 78;
						if (optDelete.IsChecked == true)
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
							goto IL_0bf3;
						}
						goto IL_0c05;
						IL_0eb2:
						num2 = 108;
						range4 = null;
						goto IL_0eb8;
						IL_0bf3:
						num2 = 79;
						xmlAttribute3.Value = Conversions.ToString(2);
						goto IL_0c05;
						IL_0c05:
						xmlAttribute3 = null;
						goto IL_0c08;
						IL_0c08:
						num2 = 81;
						xmlNode.Attributes[VH.A(206909)].Value = Conversions.ToString(numInsert.Value.Value);
						goto IL_0c4b;
						IL_0c4b:
						num2 = 82;
						xmlNode.Attributes[VH.A(206924)].Value = Conversions.ToString(numEvery.Value.Value);
						goto IL_0c8a;
						IL_0c8a:
						num2 = 83;
						xmlAttribute = xmlNode.Attributes[VH.A(206939)];
						goto IL_0cab;
						IL_0cab:
						num2 = 84;
						if (optInherit.IsChecked == true)
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
							goto IL_0cd2;
						}
						goto IL_0ce6;
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
				try0000_dispatch = 4284;
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
			switch (6)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private Range A(Range A)
	{
		Worksheet worksheet = A.Worksheet;
		object objectValue = RuntimeHelpers.GetObjectValue(Missing.Value);
		double? value = numInsert.Value;
		double? obj;
		if (!value.HasValue)
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
			obj = null;
		}
		else
		{
			obj = value.GetValueOrDefault() - 1.0;
		}
		return ((_Worksheet)worksheet).get_Range((object)A, (object)A.get_Offset(objectValue, (object)obj));
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
		if (!B)
		{
			B = true;
			Uri resourceLocator = new Uri(VH.A(207276), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b1: Expected O, but got Unknown
		//IL_0116: Unknown result type (might be due to invalid IL or missing references)
		//IL_0120: Expected O, but got Unknown
		//IL_0210: Unknown result type (might be due to invalid IL or missing references)
		//IL_021a: Expected O, but got Unknown
		if (connectionId == 1)
		{
			while (true)
			{
				switch (7)
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
				switch (2)
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
				switch (4)
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
			optInsert = (RadioButton)target;
			return;
		}
		if (connectionId == 6)
		{
			lbl3 = (TextBlock)target;
			return;
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					numEvery = (MacNumericUpDown)target;
					return;
				}
			}
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
				switch (1)
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
				switch (5)
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
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					numCols = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (6)
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
				switch (2)
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
				switch (5)
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
			cbxStyles = (ComboBox)target;
			return;
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (2)
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
			grpWidth = (GroupBox)target;
			return;
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (1)
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
			optSetTo = (RadioButton)target;
			return;
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					numColWidth = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			chkSave = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnCancel = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnOk = (Button)target;
					return;
				}
			}
		}
		B = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
