using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Xml;
using A;
using Foo.Controls;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Charts;

[DesignerGenerated]
public sealed class wpfChartSize : System.Windows.Window, IComponentConnector
{
	private Microsoft.Office.Interop.Excel.Application m_A;

	private Chart m_A;

	private bool m_A;

	private bool m_B;

	private bool m_C;

	private float m_A;

	private float m_B;

	private float m_C;

	private float m_D;

	private float E;

	private float F;

	private List<StandardSize> m_A;

	private const int m_A = 4;

	[AccessedThroughProperty("cbxChartSize")]
	[CompilerGenerated]
	private ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("numChartHeight")]
	private MacNumericUpDown m_A;

	[AccessedThroughProperty("numChartWidth")]
	[CompilerGenerated]
	private MacNumericUpDown m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("numPlotHeight")]
	private MacNumericUpDown m_C;

	[AccessedThroughProperty("numPlotWidth")]
	[CompilerGenerated]
	private MacNumericUpDown m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("numTop")]
	private MacNumericUpDown E;

	[CompilerGenerated]
	[AccessedThroughProperty("numLeft")]
	private MacNumericUpDown F;

	[AccessedThroughProperty("btnVert")]
	[CompilerGenerated]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnHoriz")]
	private Button m_B;

	[AccessedThroughProperty("optInches")]
	[CompilerGenerated]
	private RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("optCentimeters")]
	private RadioButton m_B;

	[AccessedThroughProperty("btnMemorize")]
	[CompilerGenerated]
	private Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private Button m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnReset")]
	private Button E;

	private bool m_D;

	internal virtual ComboBox cbxChartSize
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
			SelectionChangedEventHandler value2 = cbxChartSize_SelectionChanged;
			ComboBox comboBox = this.m_A;
			if (comboBox != null)
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
				comboBox.SelectionChanged -= value2;
			}
			this.m_A = value;
			comboBox = this.m_A;
			if (comboBox == null)
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
				comboBox.SelectionChanged += value2;
				return;
			}
		}
	}

	internal virtual MacNumericUpDown numChartHeight
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
			//IL_0007: Unknown result type (might be due to invalid IL or missing references)
			//IL_000d: Expected O, but got Unknown
			MacRangeBaseValueChangedHandler val = new MacRangeBaseValueChangedHandler(numChartHeight_ValueChanged);
			MacNumericUpDown val2 = this.m_A;
			if (val2 != null)
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
				val2.ValueChanged -= val;
			}
			this.m_A = value;
			val2 = this.m_A;
			if (val2 == null)
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
				val2.ValueChanged += val;
				return;
			}
		}
	}

	internal virtual MacNumericUpDown numChartWidth
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
			//IL_0007: Unknown result type (might be due to invalid IL or missing references)
			//IL_000d: Expected O, but got Unknown
			MacRangeBaseValueChangedHandler val = new MacRangeBaseValueChangedHandler(numChartWidth_ValueChanged);
			MacNumericUpDown val2 = this.m_B;
			if (val2 != null)
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
				val2.ValueChanged -= val;
			}
			this.m_B = value;
			val2 = this.m_B;
			if (val2 == null)
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
				val2.ValueChanged += val;
				return;
			}
		}
	}

	internal virtual MacNumericUpDown numPlotHeight
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
			//IL_0007: Unknown result type (might be due to invalid IL or missing references)
			//IL_000d: Expected O, but got Unknown
			MacRangeBaseValueChangedHandler val = new MacRangeBaseValueChangedHandler(numPlotHeight_ValueChanged);
			MacNumericUpDown val2 = this.m_C;
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
			this.m_C = value;
			val2 = this.m_C;
			if (val2 == null)
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
				val2.ValueChanged += val;
				return;
			}
		}
	}

	internal virtual MacNumericUpDown numPlotWidth
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
			//IL_0007: Unknown result type (might be due to invalid IL or missing references)
			//IL_000d: Expected O, but got Unknown
			MacRangeBaseValueChangedHandler val = new MacRangeBaseValueChangedHandler(numPlotWidth_ValueChanged);
			MacNumericUpDown val2 = this.m_D;
			if (val2 != null)
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
				val2.ValueChanged -= val;
			}
			this.m_D = value;
			val2 = this.m_D;
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

	internal virtual MacNumericUpDown numTop
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
			//IL_0007: Unknown result type (might be due to invalid IL or missing references)
			//IL_000d: Expected O, but got Unknown
			MacRangeBaseValueChangedHandler val = new MacRangeBaseValueChangedHandler(numTop_ValueChanged);
			MacNumericUpDown val2 = this.E;
			if (val2 != null)
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
				val2.ValueChanged -= val;
			}
			this.E = value;
			val2 = this.E;
			if (val2 == null)
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
				val2.ValueChanged += val;
				return;
			}
		}
	}

	internal virtual MacNumericUpDown numLeft
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
			//IL_0007: Unknown result type (might be due to invalid IL or missing references)
			//IL_000d: Expected O, but got Unknown
			MacRangeBaseValueChangedHandler val = new MacRangeBaseValueChangedHandler(numLeft_ValueChanged);
			MacNumericUpDown val2 = F;
			if (val2 != null)
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
				val2.ValueChanged -= val;
			}
			F = value;
			val2 = F;
			if (val2 != null)
			{
				val2.ValueChanged += val;
			}
		}
	}

	internal virtual Button btnVert
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
			RoutedEventHandler value2 = btnVert_Click;
			Button button = this.m_A;
			if (button != null)
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
				switch (5)
				{
				case 0:
					continue;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnHoriz
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
			RoutedEventHandler value2 = btnHoriz_Click;
			Button button = this.m_B;
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
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual RadioButton optInches
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
			RoutedEventHandler value2 = optInches_CheckedChanged;
			RoutedEventHandler value3 = optInches_CheckedChanged;
			RadioButton radioButton = this.m_A;
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
			this.m_A = value;
			radioButton = this.m_A;
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
				radioButton.Checked += value2;
				radioButton.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual RadioButton optCentimeters
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
			RoutedEventHandler value2 = optCentimeters_CheckedChanged;
			RoutedEventHandler value3 = optCentimeters_CheckedChanged;
			RadioButton radioButton = this.m_B;
			if (radioButton != null)
			{
				radioButton.Checked -= value2;
				radioButton.Unchecked -= value3;
			}
			this.m_B = value;
			radioButton = this.m_B;
			if (radioButton != null)
			{
				radioButton.Checked += value2;
				radioButton.Unchecked += value3;
			}
		}
	}

	internal virtual Button btnMemorize
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
			RoutedEventHandler value2 = Memorize;
			Button button = this.m_C;
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
			this.m_C = value;
			button = this.m_C;
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

	internal virtual Button btnClose
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
			RoutedEventHandler value2 = btnClose_Click;
			Button button = this.m_D;
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
			this.m_D = value;
			button = this.m_D;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnReset
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
			RoutedEventHandler value2 = btnReset_Click;
			Button button = E;
			if (button != null)
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
				button.Click -= value2;
			}
			E = value;
			button = E;
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

	public wpfChartSize(Chart cht)
	{
		base.Loaded += wpfChartSize_Loaded;
		base.Closing += wpfChartSize_Closing;
		this.m_A = true;
		int try0031_dispatch = -1;
		int num2 = default(int);
		int num = default(int);
		int num3 = default(int);
		List<StandardSizeItem> list = default(List<StandardSizeItem>);
		ChartObject chartObject = default(ChartObject);
		PlotArea plotArea = default(PlotArea);
		List<StandardSizeItem> list2 = default(List<StandardSizeItem>);
		IEnumerator enumerator = default(IEnumerator);
		XmlNode nd = default(XmlNode);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0031_dispatch)
				{
				default:
					num2 = 1;
					InitializeComponent();
					goto IL_0039;
				case 1053:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0031;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0039;
						case 3:
							goto IL_0048;
						case 4:
							goto IL_004f;
						case 5:
							goto IL_0065;
						case 6:
							goto IL_0077;
						case 8:
							goto IL_008e;
						case 7:
						case 9:
							goto IL_00a1;
						case 10:
							goto IL_00b3;
						case 11:
							goto IL_00e3;
						case 12:
							goto IL_0113;
						case 13:
							goto IL_0116;
						case 14:
							goto IL_0121;
						case 15:
							goto IL_014d;
						case 16:
							goto IL_0179;
						case 17:
							goto IL_01a5;
						case 18:
							goto IL_01d5;
						case 19:
							goto IL_01d8;
						case 20:
							goto IL_01fa;
						case 21:
							goto IL_021a;
						case 22:
							goto IL_0238;
						case 23:
							goto IL_025a;
						case 24:
							goto IL_027c;
						case 25:
							goto IL_02a0;
						case 26:
							goto IL_02b0;
						case 27:
							goto IL_02b9;
						case 28:
							goto IL_02bf;
						case 29:
							goto IL_02ce;
						case 30:
							goto IL_02ef;
						case 31:
							goto IL_0300;
						case 32:
							goto IL_0321;
						case 33:
							goto IL_0343;
						case 34:
							goto IL_0346;
						case 35:
							goto IL_0357;
						case 36:
							goto IL_035c;
						case 37:
							goto end_IL_0031_2;
						default:
							goto end_IL_0031;
						case 38:
							goto end_IL_0031_3;
						}
						goto default;
					}
					IL_0343:
					list = null;
					goto IL_0346;
					IL_0039:
					num2 = 2;
					base.Icon = Forms.GetIcon();
					goto IL_0048;
					IL_0048:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_004f;
					IL_004f:
					num2 = 4;
					this.m_A = MH.A.Application;
					goto IL_0065;
					IL_0065:
					num2 = 5;
					if (!RegionInfo.CurrentRegion.IsMetric)
					{
						goto IL_0077;
					}
					goto IL_008e;
					IL_0077:
					num2 = 6;
					optInches.IsChecked = true;
					goto IL_00a1;
					IL_008e:
					num2 = 8;
					optCentimeters.IsChecked = true;
					goto IL_00a1;
					IL_00a1:
					num2 = 9;
					chartObject = (ChartObject)cht.Parent;
					goto IL_00b3;
					IL_00b3:
					num2 = 10;
					numChartHeight.Value = Math.Round(A((float)chartObject.Height), 4);
					goto IL_00e3;
					IL_00e3:
					num2 = 11;
					numChartWidth.Value = Math.Round(A((float)chartObject.Width), 4);
					goto IL_0113;
					IL_0113:
					chartObject = null;
					goto IL_0116;
					IL_0116:
					num2 = 13;
					plotArea = cht.PlotArea;
					goto IL_0121;
					IL_0121:
					num2 = 14;
					numPlotHeight.Value = Math.Round(A((float)plotArea.InsideHeight), 4);
					goto IL_014d;
					IL_014d:
					num2 = 15;
					numPlotWidth.Value = Math.Round(A((float)plotArea.InsideWidth), 4);
					goto IL_0179;
					IL_0179:
					num2 = 16;
					numTop.Value = Math.Round(A((float)plotArea.InsideTop), 4);
					goto IL_01a5;
					IL_01a5:
					num2 = 17;
					numLeft.Value = Math.Round(A((float)plotArea.InsideLeft), 4);
					goto IL_01d5;
					IL_01d5:
					plotArea = null;
					goto IL_01d8;
					IL_01d8:
					num2 = 19;
					this.m_A = (float)numChartHeight.Value.Value;
					goto IL_01fa;
					IL_01fa:
					num2 = 20;
					this.m_B = (float)numChartWidth.Value.Value;
					goto IL_021a;
					IL_021a:
					num2 = 21;
					this.m_C = (float)numPlotHeight.Value.Value;
					goto IL_0238;
					IL_0238:
					num2 = 22;
					this.m_D = (float)numPlotWidth.Value.Value;
					goto IL_025a;
					IL_025a:
					num2 = 23;
					this.E = (float)numTop.Value.Value;
					goto IL_027c;
					IL_027c:
					num2 = 24;
					this.F = (float)numLeft.Value.Value;
					goto IL_02a0;
					IL_02a0:
					num2 = 25;
					this.m_A = clsPublish.GetStandardSizes();
					goto IL_02b0;
					IL_02b0:
					num2 = 26;
					list2 = new List<StandardSizeItem>();
					goto IL_02b9;
					IL_02b9:
					num2 = 27;
					list = list2;
					goto IL_02bf;
					IL_02bf:
					num2 = 28;
					list.Add(new StandardSizeItem());
					goto IL_02ce;
					IL_02ce:
					num2 = 29;
					enumerator = clsPublish.StandardSizeNodes().GetEnumerator();
					goto IL_0303;
					IL_0303:
					if (enumerator.MoveNext())
					{
						nd = (XmlNode)enumerator.Current;
						goto IL_02ef;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_0321;
					IL_0346:
					num2 = 34;
					cbxChartSize.ItemsSource = list2;
					goto IL_0357;
					IL_035c:
					num2 = 36;
					this.m_A = cht;
					break;
					IL_0357:
					num2 = 35;
					list2 = null;
					goto IL_035c;
					IL_0321:
					num2 = 32;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_0343;
					IL_02ef:
					num2 = 30;
					list.Add(new StandardSizeItem(nd));
					goto IL_0300;
					IL_0300:
					num2 = 31;
					goto IL_0303;
					end_IL_0031_2:
					break;
				}
				num2 = 37;
				this.m_A = false;
				break;
				end_IL_0031:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0031_dispatch = 1053;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0031_3:
			break;
		}
		if (num == 0)
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void wpfChartSize_Loaded(object sender, RoutedEventArgs e)
	{
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		float num5 = default(float);
		ComboBox comboBox = default(ComboBox);
		int num6 = default(int);
		int num7 = default(int);
		float num9 = default(float);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					bool? obj;
					bool? flag;
					double? value;
					double num8;
					bool? obj2;
					bool? flag2;
					switch (try0000_dispatch)
					{
					default:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0007;
					case 564:
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
								goto IL_0010;
							case 4:
								goto IL_0019;
							case 5:
								goto IL_0022;
							case 6:
								goto IL_003d;
							case 7:
								goto IL_0057;
							case 8:
								goto IL_006f;
							case 9:
								goto IL_009c;
							case 10:
								goto IL_00b3;
							case 11:
								goto IL_00ca;
							case 12:
								goto IL_01a6;
							case 14:
								goto IL_01b3;
							case 13:
							case 15:
								goto IL_01cf;
							case 16:
								goto end_IL_0000_2;
							default:
								goto end_IL_0000;
							case 17:
								goto end_IL_0000_3;
							}
							goto default;
						}
						IL_00b3:
						num2 = 10;
						num5 = (float)Math.Round(num5 * clsPublish.CENTIMETERS_PER_INCH, 4);
						goto IL_00ca;
						IL_0007:
						num2 = 2;
						this.m_A = true;
						goto IL_0010;
						IL_0010:
						num2 = 3;
						comboBox = cbxChartSize;
						goto IL_0019;
						IL_0019:
						num2 = 4;
						comboBox.SelectedIndex = 0;
						goto IL_0022;
						IL_0022:
						num2 = 5;
						num6 = comboBox.Items.Count - 1;
						num7 = 1;
						goto IL_01bc;
						IL_01bc:
						if (num7 <= num6)
						{
							goto IL_003d;
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
						goto IL_01cf;
						IL_00ca:
						num2 = 11;
						value = numChartHeight.Value;
						num8 = num9;
						if (!value.HasValue)
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
							obj = null;
						}
						else
						{
							obj = value.GetValueOrDefault() == num8;
						}
						flag = obj;
						if (flag.HasValue)
						{
							if (flag != true)
							{
								goto IL_01b3;
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
						value = numChartWidth.Value;
						num8 = num5;
						if (!value.HasValue)
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
							obj2 = null;
						}
						else
						{
							obj2 = value.GetValueOrDefault() == num8;
						}
						flag2 = obj2;
						if (flag2 == true)
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
							if (flag.HasValue)
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
								goto IL_01a6;
							}
						}
						goto IL_01b3;
						IL_003d:
						num2 = 6;
						num9 = this.m_A[num7 - 1].Height;
						goto IL_0057;
						IL_0057:
						num2 = 7;
						num5 = this.m_A[num7 - 1].Width;
						goto IL_006f;
						IL_006f:
						num2 = 8;
						if (optCentimeters.IsChecked == true)
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
							goto IL_009c;
						}
						goto IL_00ca;
						IL_01b3:
						num2 = 14;
						num7++;
						goto IL_01bc;
						IL_01a6:
						num2 = 12;
						comboBox.SelectedIndex = num7;
						goto IL_01cf;
						IL_01cf:
						comboBox = null;
						break;
						IL_009c:
						num2 = 9;
						num9 = (float)Math.Round(num9 * clsPublish.CENTIMETERS_PER_INCH, 4);
						goto IL_00b3;
						end_IL_0000_2:
						break;
					}
					num2 = 16;
					this.m_A = false;
					break;
				}
				end_IL_0000:;
			}
			catch (object obj3) when (obj3 is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj3);
				try0000_dispatch = 564;
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

	private void wpfChartSize_Closing(object sender, CancelEventArgs e)
	{
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
	}

	private void btnReset_Click(object sender, RoutedEventArgs e)
	{
		this.m_C = true;
		numChartHeight.Value = this.m_A;
		numChartWidth.Value = this.m_B;
		numPlotHeight.Value = this.m_C;
		numPlotWidth.Value = this.m_D;
		numTop.Value = this.E;
		numLeft.Value = this.F;
		this.m_C = false;
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = false;
		Close();
	}

	private void cbxChartSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_0071: Unknown result type (might be due to invalid IL or missing references)
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_0078: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e4: Unknown result type (might be due to invalid IL or missing references)
		//IL_010f: Unknown result type (might be due to invalid IL or missing references)
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		StandardSize val = default(StandardSize);
		int selectedIndex = default(int);
		int index = default(int);
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
				case 423:
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
							goto IL_0035;
						case 6:
							goto IL_003c;
						case 7:
							goto IL_004f;
						case 8:
							goto IL_0058;
						case 9:
							goto IL_0066;
						case 10:
							goto IL_007a;
						case 11:
							goto IL_00a1;
						case 12:
							goto IL_00be;
						case 14:
							goto IL_00db;
						case 15:
							goto IL_0104;
						case 13:
						case 16:
						case 17:
							goto IL_012d;
						case 18:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
						case 19:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_012d:
					num2 = 17;
					this.m_A.ScreenUpdating = true;
					break;
					IL_0007:
					num2 = 2;
					if (this.m_A)
					{
						goto end_IL_0000_3;
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
					goto IL_0027;
					IL_00a1:
					num2 = 11;
					numChartHeight.Value = val.Height;
					goto IL_00be;
					IL_00be:
					num2 = 12;
					numChartWidth.Value = val.Width;
					goto IL_012d;
					IL_00db:
					num2 = 14;
					numChartHeight.Value = Math.Round(val.Height * clsPublish.CENTIMETERS_PER_INCH, 4);
					goto IL_0104;
					IL_0027:
					num2 = 4;
					selectedIndex = cbxChartSize.SelectedIndex;
					goto IL_0035;
					IL_0035:
					num2 = 5;
					index = checked(selectedIndex - 1);
					goto IL_003c;
					IL_003c:
					num2 = 6;
					if (selectedIndex <= 0)
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
					goto IL_004f;
					IL_0104:
					num2 = 15;
					numChartWidth.Value = Math.Round(val.Width * clsPublish.CENTIMETERS_PER_INCH, 4);
					goto IL_012d;
					IL_004f:
					num2 = 7;
					this.m_B = true;
					goto IL_0058;
					IL_0058:
					num2 = 8;
					this.m_A.ScreenUpdating = false;
					goto IL_0066;
					IL_0066:
					num2 = 9;
					val = this.m_A[index];
					goto IL_007a;
					IL_007a:
					num2 = 10;
					if (optInches.IsChecked == true)
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
						goto IL_00a1;
					}
					goto IL_00db;
					end_IL_0000_2:
					break;
				}
				num2 = 18;
				this.m_B = false;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 423;
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
			switch (7)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void numChartHeight_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		if (this.m_A)
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
			((ChartObject)this.m_A.Parent).Height = A(numChartHeight);
			D();
			this.m_A = true;
			numPlotHeight.Value = A((float)this.m_A.PlotArea.InsideHeight);
			numTop.Value = A((float)this.m_A.PlotArea.InsideTop);
			A();
			return;
		}
	}

	private void numChartWidth_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		if (this.m_A)
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
			((ChartObject)this.m_A.Parent).Width = A(numChartWidth);
			D();
			this.m_A = true;
			numPlotWidth.Value = A((float)this.m_A.PlotArea.InsideWidth);
			numLeft.Value = A((float)this.m_A.PlotArea.InsideLeft);
			A();
			return;
		}
	}

	private void A()
	{
		//IL_0077: Unknown result type (might be due to invalid IL or missing references)
		//IL_007c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0098: Unknown result type (might be due to invalid IL or missing references)
		float num = (float)numChartWidth.Value.Value;
		float num2 = (float)numChartHeight.Value.Value;
		if (!this.m_B)
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
			cbxChartSize.SelectedIndex = 0;
		}
		checked
		{
			int num3 = this.m_A.Count - 1;
			int num4 = 0;
			while (true)
			{
				if (num4 <= num3)
				{
					if (num2 == this.m_A[num4].Height)
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
						if (num == this.m_A[num4].Width)
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
							cbxChartSize.SelectedIndex = num4 + 1;
							break;
						}
					}
					num4++;
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
			this.m_A = false;
		}
	}

	private void numPlotHeight_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		MacNumericUpDown a = default(MacNumericUpDown);
		PlotArea plotArea = default(PlotArea);
		float num5 = default(float);
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
				case 378:
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
							goto IL_0032;
						case 6:
							goto IL_003d;
						case 7:
							goto IL_004e;
						case 8:
							goto IL_005a;
						case 9:
							goto IL_0062;
						case 10:
							goto IL_006c;
						case 11:
							goto IL_0094;
						case 12:
							goto IL_009e;
						case 13:
							goto IL_00b3;
						case 14:
							goto IL_00f9;
						case 15:
							goto IL_0119;
						case 16:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
						case 17:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_009e:
					num2 = 12;
					if (!this.m_C)
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
						goto IL_00b3;
					}
					goto IL_0119;
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
					IL_00f9:
					num2 = 14;
					A(a, (float)plotArea.InsideHeight, VH.A(73326));
					goto IL_0119;
					IL_00b3:
					num2 = 13;
					if (Math.Round(A((float)plotArea.InsideHeight), 4) != Math.Round(A((float)Math.Floor(num5)), 4))
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
						goto IL_00f9;
					}
					goto IL_0119;
					IL_0119:
					plotArea = null;
					break;
					IL_0027:
					num2 = 4;
					a = numPlotHeight;
					goto IL_0032;
					IL_0032:
					num2 = 5;
					num5 = A(a);
					goto IL_003d;
					IL_003d:
					num2 = 6;
					plotArea = this.m_A.PlotArea;
					goto IL_004e;
					IL_004e:
					num2 = 7;
					plotArea.InsideHeight = num5;
					goto IL_005a;
					IL_005a:
					num2 = 8;
					D();
					goto IL_0062;
					IL_0062:
					num2 = 9;
					this.m_A = true;
					goto IL_006c;
					IL_006c:
					num2 = 10;
					numPlotWidth.Value = A((float)plotArea.InsideWidth);
					goto IL_0094;
					IL_0094:
					num2 = 11;
					this.m_A = false;
					goto IL_009e;
					end_IL_0000_2:
					break;
				}
				num2 = 16;
				a = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 378;
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

	private void numPlotWidth_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		MacNumericUpDown a = default(MacNumericUpDown);
		PlotArea plotArea = default(PlotArea);
		float num5 = default(float);
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
				case 372:
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
							goto IL_0032;
						case 6:
							goto IL_003f;
						case 7:
							goto IL_0050;
						case 8:
							goto IL_005c;
						case 9:
							goto IL_0064;
						case 10:
							goto IL_006e;
						case 11:
							goto IL_0096;
						case 12:
							goto IL_00a0;
						case 13:
							goto IL_00b5;
						case 14:
							goto IL_00f5;
						case 15:
							goto IL_0113;
						case 16:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
						case 17:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00a0:
					num2 = 12;
					if (!this.m_C)
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
						goto IL_00b5;
					}
					goto IL_0113;
					IL_0007:
					num2 = 2;
					if (this.m_A)
					{
						goto end_IL_0000_3;
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
					goto IL_0027;
					IL_00f5:
					num2 = 14;
					A(a, (float)plotArea.InsideWidth, VH.A(73405));
					goto IL_0113;
					IL_00b5:
					num2 = 13;
					if (Math.Round(A((float)plotArea.InsideWidth), 4) != Math.Round(A((float)Math.Floor(num5)), 4))
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
						goto IL_00f5;
					}
					goto IL_0113;
					IL_0113:
					plotArea = null;
					break;
					IL_0027:
					num2 = 4;
					a = numPlotWidth;
					goto IL_0032;
					IL_0032:
					num2 = 5;
					num5 = A(a);
					goto IL_003f;
					IL_003f:
					num2 = 6;
					plotArea = this.m_A.PlotArea;
					goto IL_0050;
					IL_0050:
					num2 = 7;
					plotArea.InsideWidth = num5;
					goto IL_005c;
					IL_005c:
					num2 = 8;
					D();
					goto IL_0064;
					IL_0064:
					num2 = 9;
					this.m_A = true;
					goto IL_006e;
					IL_006e:
					num2 = 10;
					numPlotHeight.Value = A((float)plotArea.InsideHeight);
					goto IL_0096;
					IL_0096:
					num2 = 11;
					this.m_A = false;
					goto IL_00a0;
					end_IL_0000_2:
					break;
				}
				num2 = 16;
				a = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 372;
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

	private void numTop_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		MacNumericUpDown a = default(MacNumericUpDown);
		PlotArea plotArea = default(PlotArea);
		float num5 = default(float);
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
				case 339:
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
							goto IL_0032;
						case 6:
							goto IL_003f;
						case 7:
							goto IL_0050;
						case 8:
							goto IL_005c;
						case 9:
							goto IL_0064;
						case 10:
							goto IL_0072;
						case 11:
							goto IL_0093;
						case 13:
							goto IL_00b5;
						case 14:
							goto IL_00d2;
						case 12:
						case 15:
							goto IL_00f2;
						case 16:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
						case 17:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0093:
					num2 = 11;
					A(a, (float)plotArea.InsideTop, VH.A(73482));
					goto IL_00f2;
					IL_0007:
					num2 = 2;
					if (this.m_A)
					{
						goto end_IL_0000_3;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_0027;
					IL_00b5:
					num2 = 13;
					if (Math.Round(plotArea.InsideTop, 4) > Math.Round(num5, 4))
					{
						goto IL_00d2;
					}
					goto IL_00f2;
					IL_00d2:
					num2 = 14;
					A(a, (float)plotArea.InsideTop, VH.A(73573));
					goto IL_00f2;
					IL_00f2:
					plotArea = null;
					break;
					IL_0027:
					num2 = 4;
					a = numTop;
					goto IL_0032;
					IL_0032:
					num2 = 5;
					num5 = A(a);
					goto IL_003f;
					IL_003f:
					num2 = 6;
					plotArea = this.m_A.PlotArea;
					goto IL_0050;
					IL_0050:
					num2 = 7;
					plotArea.InsideTop = num5;
					goto IL_005c;
					IL_005c:
					num2 = 8;
					D();
					goto IL_0064;
					IL_0064:
					num2 = 9;
					if (!this.m_C)
					{
						goto IL_0072;
					}
					goto IL_00f2;
					IL_0072:
					num2 = 10;
					if (Math.Round(plotArea.InsideTop, 4) < Math.Round(num5, 4))
					{
						goto IL_0093;
					}
					goto IL_00b5;
					end_IL_0000_2:
					break;
				}
				num2 = 16;
				a = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 339;
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

	private void numLeft_ValueChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		PlotArea plotArea = default(PlotArea);
		float num5 = default(float);
		MacNumericUpDown a = default(MacNumericUpDown);
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
				case 355:
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
							goto IL_0032;
						case 6:
							goto IL_003d;
						case 7:
							goto IL_004e;
						case 8:
							goto IL_005a;
						case 9:
							goto IL_0062;
						case 10:
							goto IL_007a;
						case 11:
							goto IL_0099;
						case 13:
							goto IL_00b9;
						case 14:
							goto IL_00e2;
						case 12:
						case 15:
							goto IL_0102;
						case 16:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 3:
						case 17:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_007a:
					num2 = 10;
					if (Math.Round(plotArea.InsideLeft, 4) < Math.Round(num5, 4))
					{
						goto IL_0099;
					}
					goto IL_00b9;
					IL_0007:
					num2 = 2;
					if (this.m_A)
					{
						goto end_IL_0000_3;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					goto IL_0027;
					IL_0099:
					num2 = 11;
					A(a, (float)plotArea.InsideLeft, VH.A(73664));
					goto IL_0102;
					IL_00b9:
					num2 = 13;
					if (Math.Round(plotArea.InsideLeft, 4) > Math.Round(num5, 4))
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
						goto IL_00e2;
					}
					goto IL_0102;
					IL_0102:
					plotArea = null;
					break;
					IL_0027:
					num2 = 4;
					a = numLeft;
					goto IL_0032;
					IL_0032:
					num2 = 5;
					num5 = A(a);
					goto IL_003d;
					IL_003d:
					num2 = 6;
					plotArea = this.m_A.PlotArea;
					goto IL_004e;
					IL_004e:
					num2 = 7;
					plotArea.InsideLeft = num5;
					goto IL_005a;
					IL_005a:
					num2 = 8;
					D();
					goto IL_0062;
					IL_0062:
					num2 = 9;
					if (!this.m_C)
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
						goto IL_007a;
					}
					goto IL_0102;
					IL_00e2:
					num2 = 14;
					A(a, (float)plotArea.InsideLeft, VH.A(73757));
					goto IL_0102;
					end_IL_0000_2:
					break;
				}
				num2 = 16;
				a = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 355;
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

	private void A(MacNumericUpDown A, float B, string C)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), C);
		this.m_A = true;
		A.Value = this.A(B);
		this.m_A = false;
	}

	private void btnVert_Click(object sender, RoutedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		int num5 = default(int);
		PlotArea plotArea = default(PlotArea);
		double num6 = default(double);
		PlotArea plotArea2 = default(PlotArea);
		double height = default(double);
		PlotArea plotArea3 = default(PlotArea);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				double top;
				double insideTop;
				double insideHeight;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 523:
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
							goto IL_002c;
						case 4:
							goto IL_0034;
						case 5:
							goto IL_0045;
						case 6:
							goto IL_0061;
						case 7:
							goto IL_006d;
						case 8:
							goto IL_0070;
						case 9:
							goto IL_0075;
						case 10:
							goto IL_0087;
						case 11:
							goto IL_0092;
						case 12:
							goto IL_00a0;
						case 13:
							goto IL_00ae;
						case 14:
							goto IL_00ba;
						case 15:
							goto IL_00bd;
						case 16:
							goto IL_00c0;
						case 17:
							goto IL_00c3;
						case 18:
							goto IL_00e9;
						case 19:
							goto IL_00ff;
						case 20:
							goto IL_0111;
						case 21:
							goto IL_012c;
						case 22:
							goto IL_0158;
						case 23:
							goto IL_0164;
						case 24:
							goto IL_0167;
						case 25:
							goto IL_0178;
						case 26:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 27:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0167:
					num2 = 24;
					num5 = checked(num5 + 1);
					if (num5 <= 2)
					{
						goto IL_0075;
					}
					goto IL_0178;
					IL_0007:
					num2 = 2;
					this.m_A.PlotArea.InsideHeight = A(numPlotHeight);
					goto IL_002c;
					IL_002c:
					num2 = 3;
					D();
					goto IL_0034;
					IL_0034:
					num2 = 4;
					plotArea = this.m_A.PlotArea;
					goto IL_0045;
					IL_0045:
					num2 = 5;
					plotArea.Top = 0.0 - this.m_A.ChartArea.Top;
					goto IL_0061;
					IL_0061:
					num2 = 6;
					num6 = 0.0 - plotArea.Top;
					goto IL_006d;
					IL_006d:
					plotArea = null;
					goto IL_0070;
					IL_0070:
					num2 = 8;
					num5 = 1;
					goto IL_0075;
					IL_0075:
					num2 = 9;
					plotArea2 = this.m_A.PlotArea;
					goto IL_0087;
					IL_0087:
					num2 = 10;
					top = plotArea2.Top;
					goto IL_0092;
					IL_0092:
					num2 = 11;
					height = plotArea2.Height;
					goto IL_00a0;
					IL_00a0:
					num2 = 12;
					insideTop = plotArea2.InsideTop;
					goto IL_00ae;
					IL_00ae:
					num2 = 13;
					insideHeight = plotArea2.InsideHeight;
					goto IL_00ba;
					IL_00ba:
					plotArea2 = null;
					goto IL_00bd;
					IL_00bd:
					num2 = 15;
					goto IL_00c0;
					IL_00c0:
					num2 = 16;
					goto IL_00c3;
					IL_00c3:
					num2 = 17;
					if (this.m_A.HasLegend && this.m_A.Legend.IncludeInLayout)
					{
						goto IL_00e9;
					}
					goto IL_00ff;
					IL_0178:
					num2 = 25;
					D();
					break;
					IL_00e9:
					num2 = 18;
					this.m_A.Legend.IncludeInLayout = false;
					goto IL_00ff;
					IL_00ff:
					num2 = 19;
					plotArea3 = this.m_A.PlotArea;
					goto IL_0111;
					IL_0111:
					num2 = 20;
					plotArea3.Height /= 2.0;
					goto IL_012c;
					IL_012c:
					num2 = 21;
					plotArea3.Top = (this.m_A.ChartArea.Height - height - num6) / 2.0;
					goto IL_0158;
					IL_0158:
					num2 = 22;
					plotArea3.Height = height;
					goto IL_0164;
					IL_0164:
					plotArea3 = null;
					goto IL_0167;
					end_IL_0000_2:
					break;
				}
				num2 = 26;
				B();
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 523;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void btnHoriz_Click(object sender, RoutedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		PlotArea plotArea = default(PlotArea);
		PlotArea plotArea2 = default(PlotArea);
		double num5 = default(double);
		int num6 = default(int);
		PlotArea plotArea3 = default(PlotArea);
		double width = default(double);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				double left;
				double insideLeft;
				double insideWidth;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 564:
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
							goto IL_002a;
						case 4:
							goto IL_0032;
						case 5:
							goto IL_0043;
						case 6:
							goto IL_005d;
						case 7:
							goto IL_006b;
						case 8:
							goto IL_006e;
						case 9:
							goto IL_0073;
						case 10:
							goto IL_0083;
						case 11:
							goto IL_0090;
						case 12:
							goto IL_009e;
						case 13:
							goto IL_00ac;
						case 14:
							goto IL_00b8;
						case 15:
							goto IL_00bb;
						case 16:
							goto IL_00be;
						case 17:
							goto IL_00c1;
						case 18:
							goto IL_0106;
						case 19:
							goto IL_011a;
						case 20:
							goto IL_012c;
						case 21:
							goto IL_0149;
						case 22:
							goto IL_0177;
						case 23:
							goto IL_0183;
						case 24:
							goto IL_0186;
						case 25:
							goto IL_01a1;
						case 26:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 27:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_012c:
					num2 = 20;
					plotArea.Width /= 2.0;
					goto IL_0149;
					IL_0007:
					num2 = 2;
					this.m_A.PlotArea.InsideWidth = A(numPlotWidth);
					goto IL_002a;
					IL_002a:
					num2 = 3;
					D();
					goto IL_0032;
					IL_0032:
					num2 = 4;
					plotArea2 = this.m_A.PlotArea;
					goto IL_0043;
					IL_0043:
					num2 = 5;
					plotArea2.Left = 0.0 - this.m_A.ChartArea.Left;
					goto IL_005d;
					IL_005d:
					num2 = 6;
					num5 = 0.0 - plotArea2.Left;
					goto IL_006b;
					IL_006b:
					plotArea2 = null;
					goto IL_006e;
					IL_006e:
					num2 = 8;
					num6 = 1;
					goto IL_0073;
					IL_0073:
					num2 = 9;
					plotArea3 = this.m_A.PlotArea;
					goto IL_0083;
					IL_0083:
					num2 = 10;
					left = plotArea3.Left;
					goto IL_0090;
					IL_0090:
					num2 = 11;
					width = plotArea3.Width;
					goto IL_009e;
					IL_009e:
					num2 = 12;
					insideLeft = plotArea3.InsideLeft;
					goto IL_00ac;
					IL_00ac:
					num2 = 13;
					insideWidth = plotArea3.InsideWidth;
					goto IL_00b8;
					IL_00b8:
					plotArea3 = null;
					goto IL_00bb;
					IL_00bb:
					num2 = 15;
					goto IL_00be;
					IL_00be:
					num2 = 16;
					goto IL_00c1;
					IL_00c1:
					num2 = 17;
					if (this.m_A.HasLegend)
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
						if (this.m_A.Legend.IncludeInLayout)
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
							goto IL_0106;
						}
					}
					goto IL_011a;
					IL_0149:
					num2 = 21;
					plotArea.Left = (this.m_A.ChartArea.Width - width - num5) / 2.0;
					goto IL_0177;
					IL_0177:
					num2 = 22;
					plotArea.Width = width;
					goto IL_0183;
					IL_0186:
					num2 = 24;
					num6 = checked(num6 + 1);
					if (num6 <= 2)
					{
						goto IL_0073;
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
					goto IL_01a1;
					IL_0183:
					plotArea = null;
					goto IL_0186;
					IL_01a1:
					num2 = 25;
					D();
					break;
					IL_0106:
					num2 = 18;
					this.m_A.Legend.IncludeInLayout = false;
					goto IL_011a;
					IL_011a:
					num2 = 19;
					plotArea = this.m_A.PlotArea;
					goto IL_012c;
					end_IL_0000_2:
					break;
				}
				num2 = 26;
				B();
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 564;
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
			switch (4)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void B()
	{
		this.m_A = true;
		PlotArea plotArea = this.m_A.PlotArea;
		numLeft.Value = A((float)plotArea.InsideLeft);
		numTop.Value = A((float)plotArea.InsideTop);
		numPlotWidth.Value = A((float)plotArea.InsideWidth);
		numPlotHeight.Value = A((float)plotArea.InsideHeight);
		plotArea = null;
		this.m_A = false;
	}

	private void Memorize(object sender, RoutedEventArgs e)
	{
		MemorizeApply.Memorize();
	}

	private void optInches_CheckedChanged(object sender, RoutedEventArgs e)
	{
		if (this.m_A)
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
			if (optInches.IsChecked != true)
			{
				return;
			}
			double? value;
			double num;
			if (cbxChartSize.SelectedIndex == 0)
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
				MacNumericUpDown[] array = (MacNumericUpDown[])(object)new MacNumericUpDown[2] { numChartHeight, numChartWidth };
				foreach (MacNumericUpDown obj in array)
				{
					value = obj.Value;
					num = clsPublish.CENTIMETERS_PER_INCH;
					double? value2;
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
						value2 = null;
					}
					else
					{
						value2 = value.GetValueOrDefault() / num;
					}
					obj.Value = value2;
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
			else
			{
				C();
			}
			this.m_A = true;
			MacNumericUpDown obj2 = numTop;
			value = numTop.Value;
			num = clsPublish.CENTIMETERS_PER_INCH;
			double? value3;
			if (!value.HasValue)
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
				value3 = null;
			}
			else
			{
				value3 = value.GetValueOrDefault() / num;
			}
			obj2.Value = value3;
			MacNumericUpDown obj3 = numLeft;
			value = numLeft.Value;
			num = clsPublish.CENTIMETERS_PER_INCH;
			double? value4;
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
				value4 = null;
			}
			else
			{
				value4 = value.GetValueOrDefault() / num;
			}
			obj3.Value = value4;
			this.m_A = false;
			return;
		}
	}

	private void optCentimeters_CheckedChanged(object sender, RoutedEventArgs e)
	{
		if (this.m_A)
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
			if (optCentimeters.IsChecked != true)
			{
				return;
			}
			if (cbxChartSize.SelectedIndex == 0)
			{
				MacNumericUpDown[] array = (MacNumericUpDown[])(object)new MacNumericUpDown[2] { numChartHeight, numChartWidth };
				foreach (MacNumericUpDown obj in array)
				{
					obj.Value *= (double)clsPublish.CENTIMETERS_PER_INCH;
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
			else
			{
				C();
			}
			this.m_A = true;
			MacNumericUpDown obj2 = numTop;
			double? value = numTop.Value;
			double num = clsPublish.CENTIMETERS_PER_INCH;
			double? value2;
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
				value2 = null;
			}
			else
			{
				value2 = value.GetValueOrDefault() * num;
			}
			obj2.Value = value2;
			numLeft.Value *= (double)clsPublish.CENTIMETERS_PER_INCH;
			this.m_A = false;
			return;
		}
	}

	private void C()
	{
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0023: Unknown result type (might be due to invalid IL or missing references)
		//IL_0025: Unknown result type (might be due to invalid IL or missing references)
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0050: Unknown result type (might be due to invalid IL or missing references)
		this.m_B = true;
		int index = checked(cbxChartSize.SelectedIndex - 1);
		StandardSize val = this.m_A[index];
		numChartHeight.Value = val.Height * A();
		numChartWidth.Value = val.Width * A();
		this.m_B = false;
	}

	private float A()
	{
		if (optInches.IsChecked == true)
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
					return 1f;
				}
			}
		}
		return clsPublish.CENTIMETERS_PER_INCH;
	}

	private float A(MacNumericUpDown A)
	{
		if (optInches.IsChecked == true)
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
					return (float)this.m_A.InchesToPoints(A.Value.Value);
				}
			}
		}
		return (float)this.m_A.CentimetersToPoints(A.Value.Value);
	}

	private float A(float A)
	{
		if (optInches.IsChecked == true)
		{
			return A / clsPublish.POINTS_PER_INCH;
		}
		return A / clsPublish.POINTS_PER_CENTIMETER;
	}

	private void D()
	{
		Helpers.A(this.m_A, this.m_A);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_D)
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
			this.m_D = true;
			Uri resourceLocator = new Uri(VH.A(73850), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		//IL_003b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0045: Expected O, but got Unknown
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Expected O, but got Unknown
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
		//IL_0067: Unknown result type (might be due to invalid IL or missing references)
		//IL_0071: Expected O, but got Unknown
		//IL_0093: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Expected O, but got Unknown
		//IL_0082: Unknown result type (might be due to invalid IL or missing references)
		//IL_008c: Expected O, but got Unknown
		if (connectionId == 1)
		{
			cbxChartSize = (ComboBox)target;
			return;
		}
		if (connectionId == 2)
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
					numChartHeight = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			numChartWidth = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 4)
		{
			numPlotHeight = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					numPlotWidth = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					numTop = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			numLeft = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnVert = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnHoriz = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			optInches = (RadioButton)target;
			return;
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					optCentimeters = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnMemorize = (Button)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 13:
			while (true)
			{
				switch (1)
				{
				case 0:
					continue;
				}
				btnClose = (Button)target;
				return;
			}
		case 14:
			btnReset = (Button)target;
			break;
		default:
			this.m_D = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
