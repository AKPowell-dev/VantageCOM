using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Workbook;

[DesignerGenerated]
public sealed class wpfCleanNames : System.Windows.Window, IComponentConnector
{
	private readonly Action m_A;

	private bool m_A;

	private bool m_B;

	private double m_A;

	private string m_A;

	private int m_A;

	private readonly Dictionary<string, bool> m_A;

	[AccessedThroughProperty("btnSave")]
	[CompilerGenerated]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("optBasic")]
	private RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("optDeep")]
	private RadioButton m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkHidden")]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("grdCleaning")]
	[CompilerGenerated]
	private Grid m_A;

	[AccessedThroughProperty("pbCleaning")]
	[CompilerGenerated]
	private ProgressBar m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lblRemovedCount")]
	private TextBlock m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("pnlButtons")]
	private StackPanel m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnOk")]
	private Button m_B;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private Button m_C;

	private bool m_C;

	internal virtual Button btnSave
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
			RoutedEventHandler value2 = btnSave_Click;
			Button button = this.m_A;
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
			this.m_A = value;
			button = this.m_A;
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

	internal virtual RadioButton optBasic
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

	internal virtual RadioButton optDeep
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

	internal virtual System.Windows.Controls.CheckBox chkHidden
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

	internal virtual Grid grdCleaning
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

	internal virtual ProgressBar pbCleaning
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

	internal virtual TextBlock lblRemovedCount
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

	internal virtual StackPanel pnlButtons
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

	internal virtual Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
			Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
			if (button == null)
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnCancel
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
			RoutedEventHandler value2 = btnCancel_Click;
			Button button = this.m_C;
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
			this.m_C = value;
			button = this.m_C;
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

	public wpfCleanNames(Action externalActionsStopper)
	{
		base.Closing += Window_Closing;
		this.m_B = false;
		this.m_A = new Dictionary<string, bool>();
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		grdCleaning.Visibility = Visibility.Collapsed;
		this.m_A = externalActionsStopper;
		this.m_A = false;
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Range range = default(Range);
		Name A = default(Name);
		bool value = default(bool);
		bool value2 = default(bool);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		XlCalculation calculation = default(XlCalculation);
		Names names = default(Names);
		int count = default(int);
		Stopwatch stopwatch = default(Stopwatch);
		int num5 = default(int);
		bool flag = default(bool);
		string text = default(string);
		string text2 = default(string);
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
				case 1638:
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
							goto IL_001c;
						case 4:
							goto IL_0025;
						case 5:
							goto IL_002e;
						case 6:
							goto IL_0048;
						case 7:
							goto IL_0064;
						case 8:
							goto IL_006c;
						case 9:
							goto IL_007e;
						case 10:
							goto IL_0089;
						case 11:
							goto IL_0094;
						case 12:
							goto IL_00a2;
						case 13:
							goto IL_00b1;
						case 14:
							goto IL_00c5;
						case 15:
							goto IL_00d2;
						case 16:
							goto IL_00dc;
						case 17:
							goto IL_00ef;
						case 18:
							goto IL_0108;
						case 19:
							goto IL_0111;
						case 20:
							goto IL_011a;
						case 21:
							goto IL_012b;
						case 22:
							goto IL_013c;
						case 23:
							goto IL_014a;
						case 24:
							goto IL_0154;
						case 25:
							goto IL_015e;
						case 26:
							goto IL_0170;
						case 27:
							goto IL_0197;
						case 28:
							goto IL_01a0;
						case 29:
							goto IL_01a8;
						case 31:
							goto IL_01c0;
						case 32:
							goto IL_01ca;
						case 33:
							goto IL_01e4;
						case 34:
							goto IL_01ed;
						case 35:
							goto IL_0217;
						case 36:
							goto IL_023d;
						case 38:
							goto IL_024d;
						case 39:
							goto IL_0262;
						case 40:
							goto IL_0287;
						case 42:
							goto IL_0297;
						case 43:
							goto IL_02be;
						case 45:
							goto IL_02c9;
						case 46:
							goto IL_02d3;
						case 47:
							goto IL_02d9;
						case 48:
							goto IL_02e7;
						case 49:
							goto IL_02fb;
						case 50:
							goto IL_0301;
						case 51:
							goto IL_0350;
						case 52:
							goto IL_03a1;
						case 53:
							goto IL_03c1;
						case 54:
							goto IL_03d0;
						case 55:
							goto IL_03e1;
						case 57:
							goto IL_03ee;
						case 58:
							goto IL_03f4;
						case 37:
						case 41:
						case 44:
						case 56:
						case 59:
							goto IL_03fa;
						case 30:
						case 60:
							goto IL_0415;
						case 61:
							goto IL_0433;
						case 62:
							goto IL_043c;
						case 63:
							goto IL_0446;
						case 64:
							goto IL_044c;
						case 65:
							goto IL_045b;
						case 66:
							goto IL_046c;
						case 67:
							goto IL_0478;
						case 68:
							goto IL_0483;
						case 69:
							goto IL_048e;
						case 70:
							goto IL_0491;
						case 71:
							goto IL_0496;
						case 72:
							goto IL_0510;
						case 73:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 74:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_02d9:
					num2 = 47;
					range = A.RefersToRange;
					goto IL_02e7;
					IL_0007:
					num2 = 2;
					clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(186089));
					goto IL_001c;
					IL_001c:
					num2 = 3;
					this.m_A = 0;
					goto IL_0025;
					IL_0025:
					num2 = 4;
					this.m_A = true;
					goto IL_002e;
					IL_002e:
					num2 = 5;
					value = chkHidden.IsChecked.Value;
					goto IL_0048;
					IL_0048:
					num2 = 6;
					value2 = optDeep.IsChecked.Value;
					goto IL_0064;
					IL_0064:
					num2 = 7;
					C();
					goto IL_006c;
					IL_006c:
					num2 = 8;
					application = MH.A.Application;
					goto IL_007e;
					IL_007e:
					num2 = 9;
					application.ScreenUpdating = false;
					goto IL_0089;
					IL_0089:
					num2 = 10;
					application.DisplayAlerts = false;
					goto IL_0094;
					IL_0094:
					num2 = 11;
					calculation = application.Calculation;
					goto IL_00a2;
					IL_00a2:
					num2 = 12;
					application.Calculation = XlCalculation.xlCalculationManual;
					goto IL_00b1;
					IL_00b1:
					num2 = 13;
					names = application.ActiveWorkbook.Names;
					goto IL_00c5;
					IL_00c5:
					num2 = 14;
					count = names.Count;
					goto IL_00d2;
					IL_00d2:
					num2 = 15;
					this.m_B = false;
					goto IL_00dc;
					IL_00dc:
					num2 = 16;
					pbCleaning.Maximum = count;
					goto IL_00ef;
					IL_00ef:
					num2 = 17;
					pbCleaning.Value = 0.0;
					goto IL_0108;
					IL_0108:
					num2 = 18;
					this.A();
					goto IL_0111;
					IL_0111:
					num2 = 19;
					E();
					goto IL_011a;
					IL_011a:
					num2 = 20;
					btnOk.Visibility = Visibility.Collapsed;
					goto IL_012b;
					IL_012b:
					num2 = 21;
					grdCleaning.Visibility = Visibility.Visible;
					goto IL_013c;
					IL_013c:
					num2 = 22;
					this.m_A.Clear();
					goto IL_014a;
					IL_014a:
					num2 = 23;
					stopwatch = new Stopwatch();
					goto IL_0154;
					IL_0154:
					num2 = 24;
					stopwatch.Start();
					goto IL_015e;
					IL_015e:
					num2 = 25;
					num5 = names.Count;
					goto IL_0403;
					IL_0403:
					if (num5 >= 1)
					{
						goto IL_0170;
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
					goto IL_0415;
					IL_02e7:
					num2 = 48;
					if (range == null)
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
						goto IL_02fb;
					}
					goto IL_03ee;
					IL_0170:
					num2 = 26;
					if (stopwatch.ElapsedMilliseconds >= 200)
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
						goto IL_0197;
					}
					goto IL_01ca;
					IL_03fa:
					num2 = 59;
					num5 = checked(num5 + -1);
					goto IL_0403;
					IL_02fb:
					num2 = 49;
					flag = true;
					goto IL_0301;
					IL_0301:
					num2 = 50;
					if (Strings.InStr(text, VH.A(7120)) > 0)
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
						if (Strings.InStr(text, VH.A(43340)) > 0)
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
							goto IL_0350;
						}
					}
					goto IL_03d0;
					IL_0197:
					num2 = 27;
					B();
					goto IL_01a0;
					IL_01a0:
					num2 = 28;
					JH.A();
					goto IL_01a8;
					IL_01a8:
					num2 = 29;
					if (!this.m_B)
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
						goto IL_01c0;
					}
					goto IL_0415;
					IL_03e1:
					num2 = 55;
					this.A(ref A);
					goto IL_03fa;
					IL_01c0:
					num2 = 31;
					stopwatch.Restart();
					goto IL_01ca;
					IL_0415:
					num2 = 60;
					if (this.m_A > 0.0)
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
						goto IL_0433;
					}
					goto IL_043c;
					IL_03ee:
					num2 = 57;
					range = null;
					goto IL_03f4;
					IL_0433:
					num2 = 61;
					B();
					goto IL_043c;
					IL_043c:
					num2 = 62;
					stopwatch.Stop();
					goto IL_0446;
					IL_0446:
					num2 = 63;
					stopwatch = null;
					goto IL_044c;
					IL_044c:
					num2 = 64;
					grdCleaning.Visibility = Visibility.Collapsed;
					goto IL_045b;
					IL_045b:
					num2 = 65;
					btnOk.Visibility = Visibility.Visible;
					goto IL_046c;
					IL_046c:
					num2 = 66;
					application.Calculation = calculation;
					goto IL_0478;
					IL_0478:
					num2 = 67;
					application.ScreenUpdating = true;
					goto IL_0483;
					IL_0483:
					num2 = 68;
					application.DisplayAlerts = true;
					goto IL_048e;
					IL_048e:
					application = null;
					goto IL_0491;
					IL_0491:
					num2 = 70;
					names = null;
					goto IL_0496;
					IL_0496:
					num2 = 71;
					Forms.InfoMessage(VH.A(186125) + Strings.Format(this.m_A, VH.A(186142)) + VH.A(180981) + Strings.Format(count, VH.A(186142)) + VH.A(177765));
					goto IL_0510;
					IL_0510:
					num2 = 72;
					this.m_A = false;
					break;
					IL_01ca:
					num2 = 32;
					this.m_A += 1.0;
					goto IL_01e4;
					IL_01e4:
					num2 = 33;
					E();
					goto IL_01ed;
					IL_01ed:
					num2 = 34;
					A = names.Item(num5, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					goto IL_0217;
					IL_0217:
					num2 = 35;
					if (value)
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
						if (!A.Visible)
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
							goto IL_023d;
						}
					}
					goto IL_024d;
					IL_03f4:
					num2 = 58;
					A = null;
					goto IL_03fa;
					IL_0350:
					num2 = 51;
					text2 = checked(Strings.Mid(text, Strings.InStr(text, VH.A(39851)) + 1, Strings.InStr(text, VH.A(43340)) - Strings.InStr(text, VH.A(39851)) - 1));
					goto IL_03a1;
					IL_03a1:
					num2 = 52;
					text2 = Strings.Replace(text2, VH.A(7120), "");
					goto IL_03c1;
					IL_023d:
					num2 = 36;
					this.A(ref A);
					goto IL_03fa;
					IL_024d:
					num2 = 38;
					text = Conversions.ToString(A.RefersTo);
					goto IL_0262;
					IL_0262:
					num2 = 39;
					if (LikeOperator.LikeString(text, VH.A(153926), CompareMethod.Binary))
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
						goto IL_0287;
					}
					goto IL_0297;
					IL_03c1:
					num2 = 53;
					flag = this.A(text2);
					goto IL_03d0;
					IL_0287:
					num2 = 40;
					this.A(ref A);
					goto IL_03fa;
					IL_0297:
					num2 = 42;
					if (A.Name.StartsWith(VH.A(186112)))
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
						goto IL_02be;
					}
					goto IL_02c9;
					IL_03d0:
					num2 = 54;
					if (flag)
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
						goto IL_03e1;
					}
					goto IL_03ee;
					IL_02be:
					num2 = 43;
					A = null;
					goto IL_03fa;
					IL_02c9:
					num2 = 45;
					if (value2)
					{
						goto IL_02d3;
					}
					goto IL_03f4;
					IL_02d3:
					num2 = 46;
					flag = false;
					goto IL_02d9;
					end_IL_0000_2:
					break;
				}
				num2 = 73;
				base.DialogResult = true;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1638;
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

	private void A(ref Name A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
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
				case 103:
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
							goto IL_0011;
						case 4:
							goto IL_001a;
						case 5:
							goto IL_002c;
						case 6:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 7:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_002c:
					num2 = 5;
					D();
					break;
					IL_0007:
					num2 = 2;
					A.Visible = true;
					goto IL_0011;
					IL_0011:
					num2 = 3;
					A.Delete();
					goto IL_001a;
					IL_001a:
					num2 = 4;
					if (Information.Err().Number != 0)
					{
						break;
					}
					goto IL_002c;
					end_IL_0000_2:
					break;
				}
				num2 = 6;
				A = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 103;
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
			switch (1)
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

	private void A()
	{
		this.m_A = pbCleaning.Value;
		this.m_A = lblRemovedCount.Text;
	}

	private void B()
	{
		pbCleaning.Value = this.m_A;
		lblRemovedCount.Text = this.m_A;
	}

	private bool A(string A)
	{
		string text = (A ?? string.Empty).ToLower();
		if (!this.m_A.TryGetValue(text, out var value))
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
			try
			{
				int num;
				if (!clsFile.IsPathUrl(text))
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
					num = ((!File.Exists(A)) ? 1 : 0);
				}
				else
				{
					num = 0;
				}
				value = (byte)num != 0;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				value = false;
				ProjectData.ClearProjectError();
			}
			this.m_A[text] = value;
		}
		return value;
	}

	private void C()
	{
		try
		{
			this.m_A();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
	}

	private void D()
	{
		checked
		{
			this.m_A++;
			E();
		}
	}

	private void E()
	{
		this.m_A = VH.A(186125) + Strings.Format(this.m_A, VH.A(186142)) + VH.A(180981) + Strings.Format(this.m_A, VH.A(186142)) + VH.A(186153);
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		if (this.m_A)
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
			base.DialogResult = false;
			Close();
			return;
		}
	}

	private void Window_Closing(object sender, CancelEventArgs e)
	{
		e.Cancel = this.m_A;
		if (!this.m_A)
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
			this.m_B = true;
			return;
		}
	}

	private void btnSave_Click(object sender, RoutedEventArgs e)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
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
				case 145:
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
							goto IL_0017;
						case 4:
							goto IL_0030;
						case 5:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 6:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0030:
					num2 = 4;
					QuickSave.Save(MH.A.Application.ActiveWorkbook);
					break;
					IL_0007:
					num2 = 2;
					btnSave.IsEnabled = false;
					goto IL_0017;
					IL_0017:
					num2 = 3;
					btnSave.Content = VH.A(186182);
					goto IL_0030;
					end_IL_0000_2:
					break;
				}
				num2 = 5;
				btnSave.Content = VH.A(186201);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 145;
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
			switch (5)
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

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_C)
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
			this.m_C = true;
			Uri resourceLocator = new Uri(VH.A(186212), UriKind.Relative);
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
		if (connectionId == 1)
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
					btnSave = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 2)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					optBasic = (RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			optDeep = (RadioButton)target;
			return;
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkHidden = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					grdCleaning = (Grid)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					pbCleaning = (ProgressBar)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 7:
			lblRemovedCount = (TextBlock)target;
			break;
		case 8:
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				pnlButtons = (StackPanel)target;
				return;
			}
		case 9:
			btnOk = (Button)target;
			break;
		case 10:
			btnCancel = (Button)target;
			break;
		default:
			this.m_C = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
