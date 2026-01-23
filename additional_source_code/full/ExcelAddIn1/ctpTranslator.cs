using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Timers;
using System.Windows.Forms;
using A;
using ExcelAddIn1.Formulas;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[DesignerGenerated]
public sealed class ctpTranslator : UserControl
{
	private IContainer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnBack")]
	private Button m_A;

	[AccessedThroughProperty("btnForward")]
	[CompilerGenerated]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("PictureBox1")]
	private PictureBox m_A;

	[AccessedThroughProperty("CTP")]
	[CompilerGenerated]
	private CustomTaskPane m_A;

	private Collection m_A;

	private int m_A;

	private System.Timers.Timer m_A;

	internal virtual Button btnBack
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
			EventHandler value2 = B;
			Button button = this.m_A;
			if (button != null)
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
				switch (1)
				{
				case 0:
					continue;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnForward
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
			EventHandler value2 = C;
			Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual PictureBox PictureBox1
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

	public virtual CustomTaskPane CTP
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

	public ctpTranslator()
	{
		base.VisibleChanged += D;
		System.Windows.Forms.Application.EnableVisualStyles();
		A();
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).AddEventHandler(MH.A.Application, new AppEvents_SheetSelectionChangeEventHandler(A));
		this.m_A = new Collection();
		this.m_A = 0;
	}

	[DebuggerNonUserCode]
	protected override void Dispose(bool disposing)
	{
		try
		{
			if (!disposing || this.m_A == null)
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
				this.m_A.Dispose();
				return;
			}
		}
		finally
		{
			base.Dispose(disposing);
		}
	}

	[DebuggerStepThrough]
	private void A()
	{
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(ctpTranslator));
		btnBack = new Button();
		btnForward = new Button();
		PictureBox1 = new PictureBox();
		((ISupportInitialize)PictureBox1).BeginInit();
		SuspendLayout();
		btnBack.BackColor = Color.Transparent;
		btnBack.Enabled = false;
		btnBack.Image = (Image)componentResourceManager.GetObject(VH.A(204686));
		btnBack.Location = new System.Drawing.Point(24, 1);
		btnBack.Name = VH.A(204713);
		btnBack.Padding = new Padding(0, 1, 0, 0);
		btnBack.Size = new Size(20, 20);
		btnBack.TabIndex = 2;
		btnBack.UseVisualStyleBackColor = false;
		btnForward.BackColor = Color.Transparent;
		btnForward.Enabled = false;
		btnForward.Image = (Image)componentResourceManager.GetObject(VH.A(204728));
		btnForward.Location = new System.Drawing.Point(45, 1);
		btnForward.Name = VH.A(204761);
		btnForward.Padding = new Padding(1, 1, 0, 0);
		btnForward.Size = new Size(20, 20);
		btnForward.TabIndex = 3;
		btnForward.UseVisualStyleBackColor = false;
		PictureBox1.Image = (Image)componentResourceManager.GetObject(VH.A(204782));
		PictureBox1.Location = new System.Drawing.Point(3, 3);
		PictureBox1.Name = VH.A(204817);
		PictureBox1.Size = new Size(16, 16);
		PictureBox1.TabIndex = 5;
		PictureBox1.TabStop = false;
		base.AutoScaleMode = AutoScaleMode.None;
		BackColor = Color.Gray;
		base.Controls.Add(PictureBox1);
		base.Controls.Add(btnForward);
		base.Controls.Add(btnBack);
		Font = new System.Drawing.Font(VH.A(50021), 9f, FontStyle.Regular, GraphicsUnit.Point, 0);
		base.Name = VH.A(204840);
		base.Size = new Size(436, 22);
		((ISupportInitialize)PictureBox1).EndInit();
		ResumeLayout(performLayout: false);
	}

	private void A(object A, Range B)
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
				case 163:
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
							goto IL_0014;
						case 4:
							goto IL_002e;
						case 5:
							goto IL_0047;
						case 6:
							goto IL_0055;
						case 7:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 8:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0055:
					num2 = 6;
					this.m_A.AutoReset = false;
					break;
					IL_0007:
					num2 = 2;
					this.m_A.Dispose();
					goto IL_0014;
					IL_0014:
					num2 = 3;
					this.m_A = new System.Timers.Timer(KH.A.TranslatorDelay);
					goto IL_002e;
					IL_002e:
					num2 = 4;
					this.m_A.Elapsed += this.A;
					goto IL_0047;
					IL_0047:
					num2 = 5;
					this.m_A.SynchronizingObject = this;
					goto IL_0055;
					end_IL_0000_2:
					break;
				}
				num2 = 7;
				this.m_A.Enabled = true;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 163;
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

	private void A(object A, ElapsedEventArgs B)
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
				case 75:
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
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 5:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_000f:
					num2 = 3;
					this.m_A = null;
					break;
					IL_0007:
					num2 = 2;
					Translate();
					goto IL_000f;
					end_IL_0000_2:
					break;
				}
				num2 = 4;
				this.m_A.Dispose();
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 75;
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

	public void Translate()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		int num = default(int);
		int num3 = default(int);
		Range activeCell = default(Range);
		int num5 = default(int);
		int B = default(int);
		string pattern = default(string);
		int num6 = default(int);
		Microsoft.Office.Interop.Excel.Application application2 = default(Microsoft.Office.Interop.Excel.Application);
		Range range = default(Range);
		Microsoft.Office.Interop.Excel.Application application3 = default(Microsoft.Office.Interop.Excel.Application);
		string text = default(string);
		string[] array = default(string[]);
		MatchCollection matchCollection = default(MatchCollection);
		string[] array2 = default(string[]);
		int num7 = default(int);
		string text2 = default(string);
		Range range2 = default(Range);
		string[] array3 = default(string[]);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					int num4;
					object instance;
					string memberName;
					object[] array4;
					ref string reference;
					object[] array5;
					bool[] obj;
					bool[] array6;
					object obj2;
					switch (try0000_dispatch)
					{
					default:
						num2 = 1;
						application = MH.A.Application;
						goto IL_000f;
					case 1450:
						{
							num = num2;
							switch (num3)
							{
							case 2:
								break;
							case 1:
								goto IL_04b6;
							default:
								goto end_IL_0000;
							}
							goto IL_0476;
						}
						IL_04b6:
						num4 = unchecked(num + 1);
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_000f;
						case 3:
							goto IL_001f;
						case 4:
							goto IL_002b;
						case 5:
							goto IL_0030;
						case 6:
							goto IL_0036;
						case 7:
							goto IL_0044;
						case 8:
							goto IL_004b;
						case 9:
							goto IL_0065;
						case 10:
							goto IL_00af;
						case 11:
							goto IL_0119;
						case 12:
							goto IL_0132;
						case 13:
							goto IL_014d;
						case 14:
							goto IL_0154;
						case 15:
							goto IL_015b;
						case 17:
							goto IL_018e;
						case 19:
							goto IL_01a4;
						case 20:
							goto IL_01aa;
						case 21:
							goto IL_01b5;
						case 22:
							goto IL_01c0;
						case 23:
							goto IL_01cb;
						case 24:
							goto IL_01ce;
						case 25:
							goto IL_01ef;
						case 26:
							goto IL_0205;
						case 27:
							goto IL_0230;
						case 28:
							goto IL_024d;
						case 29:
							goto IL_0263;
						case 31:
							goto IL_027d;
						case 32:
							goto IL_028d;
						case 30:
						case 33:
							goto IL_029d;
						case 34:
							goto IL_02a4;
						case 35:
							goto IL_02ba;
						case 36:
							goto IL_02cb;
						case 37:
							goto IL_02eb;
						case 38:
							goto IL_0302;
						case 39:
							goto IL_034b;
						case 40:
							goto IL_0367;
						case 41:
							goto IL_03f6;
						case 42:
							goto IL_0407;
						case 44:
							goto IL_0417;
						case 45:
							goto IL_0437;
						case 43:
						case 46:
							goto IL_043d;
						case 47:
							goto IL_0459;
						case 48:
							goto IL_0462;
						case 16:
						case 18:
						case 49:
							goto IL_0476;
						case 50:
							goto IL_047c;
						case 51:
							goto IL_0487;
						case 52:
							goto IL_0492;
						case 53:
							goto IL_049d;
						case 54:
							goto IL_04a0;
						case 55:
							goto IL_04a5;
						case 56:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 57:
							goto end_IL_0000_3;
						}
						goto default;
						IL_000f:
						num2 = 2;
						RuntimeHelpers.GetObjectValue(application.ActiveSheet);
						goto IL_001f;
						IL_001f:
						num2 = 3;
						activeCell = application.ActiveCell;
						goto IL_002b;
						IL_002b:
						num2 = 4;
						num5 = 0;
						goto IL_0030;
						IL_0030:
						num2 = 5;
						B = 70;
						goto IL_0036;
						IL_0036:
						num2 = 6;
						pattern = VH.A(155686);
						goto IL_0044;
						IL_0044:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_004b;
						IL_004b:
						num2 = 8;
						num6 = base.Controls.Count - 1;
						goto IL_013b;
						IL_013b:
						if (num6 >= 0)
						{
							goto IL_0065;
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
						goto IL_014d;
						IL_0492:
						num2 = 52;
						application2.DisplayAlerts = true;
						goto IL_049d;
						IL_014d:
						ProjectData.ClearProjectError();
						num3 = 2;
						goto IL_0154;
						IL_0154:
						num2 = 14;
						range = activeCell;
						goto IL_015b;
						IL_015b:
						num2 = 15;
						if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(range.get_Value(RuntimeHelpers.GetObjectValue(Missing.Value)))))
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
							goto IL_018e;
						}
						goto IL_0476;
						IL_049d:
						application2 = null;
						goto IL_04a0;
						IL_018e:
						num2 = 17;
						if (!Conversions.ToBoolean(range.HasArray))
						{
							goto IL_01a4;
						}
						goto IL_0476;
						IL_01a4:
						num2 = 19;
						application3 = application;
						goto IL_01aa;
						IL_01aa:
						num2 = 20;
						application3.ScreenUpdating = false;
						goto IL_01b5;
						IL_01b5:
						num2 = 21;
						application3.EnableEvents = false;
						goto IL_01c0;
						IL_01c0:
						num2 = 22;
						application3.DisplayAlerts = false;
						goto IL_01cb;
						IL_01cb:
						application3 = null;
						goto IL_01ce;
						IL_01ce:
						num2 = 24;
						A(Conversions.ToString(Helpers.GetLabelCell(range).Text), ref B);
						goto IL_01ef;
						IL_01ef:
						num2 = 25;
						A(VH.A(48936), ref B);
						goto IL_0205;
						IL_0205:
						num2 = 26;
						text = Regex.Replace(Conversions.ToString(range.Formula), VH.A(157948), "");
						goto IL_0230;
						IL_0230:
						num2 = 27;
						text = Regex.Replace(text, VH.A(157953), "");
						goto IL_024d;
						IL_024d:
						num2 = 28;
						if (Versioned.IsNumeric(text))
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
							goto IL_0263;
						}
						goto IL_027d;
						IL_04a0:
						num2 = 54;
						application = null;
						goto IL_04a5;
						IL_0263:
						num2 = 29;
						array = Regex.Split(text, VH.A(150544));
						goto IL_029d;
						IL_027d:
						num2 = 31;
						matchCollection = Regex.Matches(text, pattern);
						goto IL_028d;
						IL_028d:
						num2 = 32;
						array = Regex.Split(text, pattern);
						goto IL_029d;
						IL_029d:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_02a4;
						IL_02a4:
						num2 = 34;
						array2 = array;
						num7 = 0;
						goto IL_046b;
						IL_046b:
						if (num7 < array2.Length)
						{
							text2 = array2[num7];
							goto IL_02ba;
						}
						goto IL_0476;
						IL_04a5:
						num2 = 55;
						range2 = null;
						break;
						IL_02ba:
						num2 = 35;
						if (!Versioned.IsNumeric(text2))
						{
							goto IL_02cb;
						}
						goto IL_03f6;
						IL_02cb:
						num2 = 36;
						text2 = Strings.Replace(text2, VH.A(39851), "");
						goto IL_02eb;
						IL_02eb:
						num2 = 37;
						if (Strings.InStr(text2, VH.A(7827)) == 0)
						{
							goto IL_0302;
						}
						goto IL_034b;
						IL_0302:
						num2 = 38;
						text2 = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateGet(application.ActiveSheet, null, VH.A(19019), new object[0], null, null, null), VH.A(7827)), text2));
						goto IL_034b;
						IL_034b:
						num2 = 39;
						array3 = Strings.Split(text2, VH.A(7827));
						goto IL_0367;
						IL_0367:
						num2 = 40;
						instance = application.ActiveWorkbook.Sheets[array3[0]];
						memberName = VH.A(41315);
						array4 = new object[1];
						reference = ref array3[1];
						array4[0] = reference;
						array5 = array4;
						obj = new bool[1] { true };
						array6 = obj;
						obj2 = NewLateBinding.LateGet(instance, null, memberName, array4, null, null, obj);
						if (array6[0])
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
							reference = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array5[0]), typeof(string));
						}
						range2 = (Range)obj2;
						goto IL_03f6;
						IL_00af:
						num2 = 10;
						if ((Operators.CompareString(base.Controls[num6].Name, btnBack.Name, TextCompare: false) != 0) & (Operators.CompareString(base.Controls[num6].Name, btnForward.Name, TextCompare: false) != 0))
						{
							goto IL_0119;
						}
						goto IL_0132;
						IL_0119:
						num2 = 11;
						base.Controls[num6].Dispose();
						goto IL_0132;
						IL_0065:
						num2 = 9;
						if ((base.Controls[num6] is Button) | (base.Controls[num6] is Label))
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
							goto IL_00af;
						}
						goto IL_0132;
						IL_03f6:
						num2 = 41;
						if (range2 == null)
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
							goto IL_0407;
						}
						goto IL_0417;
						IL_0132:
						num2 = 12;
						num6 += -1;
						goto IL_013b;
						IL_0407:
						num2 = 42;
						A(text2, ref B);
						goto IL_043d;
						IL_0417:
						num2 = 44;
						A(Conversions.ToString(Helpers.GetLabelCell(range2).Text), ref B, range2);
						goto IL_0437;
						IL_0437:
						num2 = 45;
						range2 = null;
						goto IL_043d;
						IL_043d:
						num2 = 46;
						A(matchCollection[num5].ToString(), ref B);
						goto IL_0459;
						IL_0459:
						num2 = 47;
						num5++;
						goto IL_0462;
						IL_0462:
						num2 = 48;
						num7++;
						goto IL_046b;
						IL_0476:
						num2 = 49;
						application2 = application;
						goto IL_047c;
						IL_047c:
						num2 = 50;
						application2.ScreenUpdating = true;
						goto IL_0487;
						IL_0487:
						num2 = 51;
						application2.EnableEvents = true;
						goto IL_0492;
						end_IL_0000_2:
						break;
					}
					num2 = 56;
					range = null;
					break;
				}
				end_IL_0000:;
			}
			catch (object obj3) when (obj3 is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj3);
				try0000_dispatch = 1450;
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void A(string A, ref int B, Range C = null)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Label label = default(Label);
		Label label2 = default(Label);
		Button button = default(Button);
		LinkLabel linkLabel = default(LinkLabel);
		System.Drawing.Font font = default(System.Drawing.Font);
		LinkLabel linkLabel2 = default(LinkLabel);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				string text;
				uint num5;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 1158:
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
							goto IL_0018;
						case 5:
							goto IL_0021;
						case 6:
							goto IL_002c;
						case 7:
							goto IL_0047;
						case 8:
							goto IL_006e;
						case 10:
							goto IL_0284;
						case 11:
							goto IL_028b;
						case 12:
							goto IL_02b5;
						case 13:
							goto IL_02c0;
						case 14:
							goto IL_02d1;
						case 15:
							goto IL_02dc;
						case 16:
							goto IL_02e7;
						case 17:
							goto IL_02f9;
						case 18:
							goto IL_030b;
						case 19:
							goto IL_031e;
						case 22:
							goto IL_0326;
						case 23:
							goto IL_032d;
						case 24:
							goto IL_0338;
						case 25:
							goto IL_0344;
						case 26:
							goto IL_0350;
						case 27:
							goto IL_0361;
						case 28:
							goto IL_036c;
						case 29:
							goto IL_0377;
						case 30:
							goto IL_0382;
						case 31:
							goto IL_038d;
						case 32:
							goto IL_039f;
						case 33:
							goto IL_03b5;
						case 34:
							goto IL_03c7;
						case 35:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 9:
						case 20:
						case 21:
						case 36:
						case 37:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0284:
					num2 = 10;
					label = label2;
					goto IL_028b;
					IL_0007:
					num2 = 2;
					button = new Button();
					goto IL_000f;
					IL_000f:
					num2 = 3;
					label2 = new Label();
					goto IL_0018;
					IL_0018:
					num2 = 4;
					linkLabel = new LinkLabel();
					goto IL_0021;
					IL_0021:
					num2 = 5;
					button.CreateGraphics();
					goto IL_002c;
					IL_002c:
					num2 = 6;
					font = new System.Drawing.Font(VH.A(50021), 9f, FontStyle.Regular);
					goto IL_0047;
					IL_0047:
					num2 = 7;
					A = Strings.Replace(A, VH.A(186705), VH.A(204867));
					goto IL_006e;
					IL_006e:
					num2 = 8;
					text = A;
					num5 = TH.A(text);
					if (num5 <= 755801111)
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
						if (num5 <= 705468254)
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
							if (num5 != 671913016)
							{
								if (num5 != 705468254)
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
								}
								else
								{
									if (Operators.CompareString(text, VH.A(75498), TextCompare: false) == 0)
									{
										goto IL_0284;
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
							}
							else
							{
								if (Operators.CompareString(text, VH.A(13778), TextCompare: false) == 0)
								{
									goto IL_0284;
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
						else if (num5 != 739023492)
						{
							if (num5 != 755801111)
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
							}
							else if (Operators.CompareString(text, VH.A(39848), TextCompare: false) == 0)
							{
								goto IL_0284;
							}
						}
						else
						{
							if (Operators.CompareString(text, VH.A(39904), TextCompare: false) == 0)
							{
								goto IL_0284;
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
					}
					else if (num5 <= 789356349)
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
						if (num5 != 772578730)
						{
							if (num5 != 789356349)
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
							}
							else
							{
								if (Operators.CompareString(text, VH.A(75231), TextCompare: false) == 0)
								{
									goto IL_0284;
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
						else
						{
							if (Operators.CompareString(text, VH.A(54459), TextCompare: false) == 0)
							{
								goto IL_0284;
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
					}
					else if (num5 != 940354920)
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
						if (num5 != 2166136261u)
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
						}
						else if (Operators.CompareString(text, "", TextCompare: false) == 0)
						{
							goto end_IL_0000_3;
						}
					}
					else
					{
						if (Operators.CompareString(text, VH.A(48936), TextCompare: false) == 0)
						{
							goto IL_0284;
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
					goto IL_0326;
					IL_028b:
					num2 = 11;
					A = Strings.Replace(A, VH.A(13778), VH.A(204872));
					goto IL_02b5;
					IL_03c7:
					num2 = 34;
					B = checked(B + linkLabel2.Width + 0);
					break;
					IL_038d:
					num2 = 31;
					linkLabel2.Location = new System.Drawing.Point(B, 0);
					goto IL_039f;
					IL_02c0:
					num2 = 13;
					label.ForeColor = Color.White;
					goto IL_02d1;
					IL_02d1:
					num2 = 14;
					label.AutoSize = true;
					goto IL_02dc;
					IL_02f9:
					num2 = 17;
					base.Controls.Add(label2);
					goto IL_030b;
					IL_030b:
					num2 = 18;
					B = checked(B + label.Width + 0);
					goto IL_031e;
					IL_031e:
					label = null;
					goto end_IL_0000_3;
					IL_0326:
					num2 = 22;
					linkLabel2 = linkLabel;
					goto IL_032d;
					IL_032d:
					num2 = 23;
					linkLabel2.Text = A;
					goto IL_0338;
					IL_02dc:
					num2 = 15;
					label.TextAlign = ContentAlignment.TopCenter;
					goto IL_02e7;
					IL_0338:
					num2 = 24;
					linkLabel2.Font = font;
					goto IL_0344;
					IL_02e7:
					num2 = 16;
					label.Location = new System.Drawing.Point(B, 2);
					goto IL_02f9;
					IL_0344:
					num2 = 25;
					linkLabel2.TextAlign = ContentAlignment.MiddleCenter;
					goto IL_0350;
					IL_0361:
					num2 = 27;
					linkLabel2.LinkBehavior = LinkBehavior.HoverUnderline;
					goto IL_036c;
					IL_036c:
					num2 = 28;
					linkLabel2.LinkVisited = false;
					goto IL_0377;
					IL_0350:
					num2 = 26;
					linkLabel2.LinkColor = Color.White;
					goto IL_0361;
					IL_0377:
					num2 = 29;
					linkLabel2.AutoSize = true;
					goto IL_0382;
					IL_02b5:
					num2 = 12;
					label.Text = A;
					goto IL_02c0;
					IL_0382:
					num2 = 30;
					linkLabel2.Tag = C;
					goto IL_038d;
					IL_039f:
					num2 = 32;
					linkLabel2.Click += this.A;
					goto IL_03b5;
					IL_03b5:
					num2 = 33;
					base.Controls.Add(linkLabel);
					goto IL_03c7;
					end_IL_0000_2:
					break;
				}
				linkLabel2 = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 1158;
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void A(object A, EventArgs B)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		LinkLabel linkLabel = default(LinkLabel);
		Microsoft.Office.Interop.Excel.Application application = default(Microsoft.Office.Interop.Excel.Application);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				checked
				{
					switch (try0000_dispatch)
					{
					default:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0007;
					case 349:
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
								goto IL_0020;
							case 5:
								goto IL_0047;
							case 6:
								goto IL_0062;
							case 7:
								goto IL_0079;
							case 8:
								goto IL_0098;
							case 9:
								goto IL_00a8;
							case 10:
								goto IL_00b9;
							case 11:
								goto IL_00cf;
							case 13:
								goto IL_00da;
							case 14:
								goto IL_00e3;
							case 12:
							case 15:
								goto end_IL_0000_2;
							default:
								goto end_IL_0000;
							case 16:
								goto end_IL_0000_3;
							}
							goto default;
						}
						IL_00a8:
						num2 = 9;
						btnBack.Enabled = true;
						goto IL_00b9;
						IL_0007:
						num2 = 2;
						linkLabel = (LinkLabel)A;
						goto IL_0010;
						IL_0010:
						num2 = 3;
						application = MH.A.Application;
						goto IL_0020;
						IL_0020:
						num2 = 4;
						if (linkLabel.Tag is Range)
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
							goto IL_0047;
						}
						goto IL_00da;
						IL_00b9:
						num2 = 10;
						this.A((Range)linkLabel.Tag);
						goto IL_00cf;
						IL_00cf:
						num2 = 11;
						this.B();
						break;
						IL_00da:
						num2 = 13;
						this.B();
						goto IL_00e3;
						IL_0047:
						num2 = 5;
						if (this.m_A.Count == 0)
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
							goto IL_0062;
						}
						goto IL_0079;
						IL_00e3:
						num2 = 14;
						application.SendKeys(VH.A(204875), RuntimeHelpers.GetObjectValue(Missing.Value));
						break;
						IL_0062:
						num2 = 6;
						this.m_A.Add(application.ActiveCell);
						goto IL_0079;
						IL_0079:
						num2 = 7;
						this.m_A.Add(RuntimeHelpers.GetObjectValue(linkLabel.Tag));
						goto IL_0098;
						IL_0098:
						num2 = 8;
						this.m_A++;
						goto IL_00a8;
						end_IL_0000_2:
						break;
					}
					application = null;
					break;
				}
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 349;
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

	private void B()
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.SendKeys(VH.A(168648), RuntimeHelpers.GetObjectValue(Missing.Value));
		application.SendKeys(VH.A(204908), RuntimeHelpers.GetObjectValue(Missing.Value));
		_ = null;
	}

	private void B(object A, EventArgs B)
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
				case 153:
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
							goto IL_001f;
						case 4:
							goto IL_0029;
						case 5:
							goto IL_0039;
						case 6:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 7:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0039:
					num2 = 5;
					btnForward.Enabled = true;
					break;
					IL_0007:
					num2 = 2;
					this.m_A = Math.Max(0, checked(this.m_A - 1));
					goto IL_001f;
					IL_001f:
					num2 = 3;
					if (this.m_A == 0)
					{
						goto IL_0029;
					}
					goto IL_0039;
					IL_0029:
					num2 = 4;
					btnBack.Enabled = false;
					goto IL_0039;
					end_IL_0000_2:
					break;
				}
				num2 = 6;
				this.A((Range)this.m_A[checked(this.m_A + 1)]);
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 153;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void C(object A, EventArgs B)
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
				checked
				{
					switch (try0000_dispatch)
					{
					default:
						ProjectData.ClearProjectError();
						num3 = 1;
						goto IL_0007;
					case 201:
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
								goto IL_002d;
							case 4:
								goto IL_0059;
							case 5:
								goto IL_0069;
							case 6:
								goto end_IL_0000_2;
							default:
								goto end_IL_0000;
							case 7:
								goto end_IL_0000_3;
							}
							goto default;
						}
						IL_0059:
						num2 = 4;
						btnForward.Enabled = false;
						goto IL_0069;
						IL_0007:
						num2 = 2;
						this.m_A = Math.Min(this.m_A.Count - 1, this.m_A + 1);
						goto IL_002d;
						IL_002d:
						num2 = 3;
						if (this.m_A == this.m_A.Count - 1)
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
							goto IL_0059;
						}
						goto IL_0069;
						IL_0069:
						num2 = 5;
						btnBack.Enabled = true;
						break;
						end_IL_0000_2:
						break;
					}
					num2 = 6;
					this.A((Range)this.m_A[this.m_A + 1]);
					break;
				}
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 201;
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
			ProjectData.ClearProjectError();
			return;
		}
	}

	private void A(Range A)
	{
		Ranges.ScrollIntoView(A);
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		application.EnableEvents = false;
		A.Select();
		application.EnableEvents = true;
		_ = null;
		Translate();
	}

	private void D(object A, EventArgs B)
	{
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		if (!base.Visible)
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
			this.m_A = null;
			new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).RemoveEventHandler(application, (AppEvents_SheetSelectionChangeEventHandler)([SpecialName] (object obj, Range range) =>
			{
				Translate();
			}));
		}
		else
		{
			Translate();
		}
		application = null;
	}

	[SpecialName]
	[CompilerGenerated]
	private void B(object A, Range B)
	{
		Translate();
	}
}
