using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Markup;
using A;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Workbook;

[DesignerGenerated]
public sealed class wpfStyleScrubber : System.Windows.Window, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class XG
	{
		public bool A;

		public wpfStyleScrubber A;

		[SpecialName]
		internal void A()
		{
			this.A = this.A.radAll.IsChecked.Value;
		}
	}

	[CompilerGenerated]
	internal sealed class YG
	{
		public int A;

		public bool A;

		public string A;

		public wpfStyleScrubber A;

		public YG(YG A)
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
				this.A = A.A;
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A = this.A.lbxCorrupt.Items.Count;
			this.A = this.A.radAll.IsChecked.Value;
		}
	}

	[CompilerGenerated]
	internal sealed class ZG
	{
		public int A;

		public YG A;

		public ZG(ZG A)
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
		internal void A()
		{
			this.A.A = this.A.A.lbxCorrupt.Items[this.A].ToString();
		}
	}

	[CompilerGenerated]
	internal sealed class AH
	{
		public TextBlock A;

		public string A;

		[SpecialName]
		internal void A()
		{
			this.A.Text = this.A;
		}
	}

	[CompilerGenerated]
	internal sealed class BH
	{
		public string A;

		public wpfStyleScrubber A;

		[SpecialName]
		internal void A()
		{
			this.A.lbxCorrupt.Items.Add(this.A);
		}
	}

	[CompilerGenerated]
	internal sealed class CH
	{
		public int A;

		public wpfStyleScrubber A;

		[SpecialName]
		internal void A()
		{
			this.A.lbxCorrupt.Items.RemoveAt(this.A);
		}
	}

	private Microsoft.Office.Interop.Excel.Application m_A;

	private BackgroundWorker m_A;

	private BackgroundWorker m_B;

	private int m_A;

	private readonly string m_A;

	[AccessedThroughProperty("tbCount")]
	[CompilerGenerated]
	private TextBlock m_A;

	[AccessedThroughProperty("radAll")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("radUnused")]
	private System.Windows.Controls.RadioButton m_B;

	[AccessedThroughProperty("lbxCorrupt")]
	[CompilerGenerated]
	private System.Windows.Controls.ListBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("tbStatus")]
	private TextBlock m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("pbProgress")]
	private System.Windows.Controls.ProgressBar m_A;

	[AccessedThroughProperty("btnStart")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	private bool m_A;

	internal virtual TextBlock tbCount
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

	internal virtual System.Windows.Controls.RadioButton radAll
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

	internal virtual System.Windows.Controls.RadioButton radUnused
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

	internal virtual System.Windows.Controls.ListBox lbxCorrupt
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

	internal virtual TextBlock tbStatus
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

	internal virtual System.Windows.Controls.ProgressBar pbProgress
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

	internal virtual System.Windows.Controls.Button btnStart
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
			RoutedEventHandler value2 = btnStart_Click;
			System.Windows.Controls.Button button = this.m_A;
			if (button != null)
			{
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click += value2;
				return;
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
				switch (1)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click += value2;
				return;
			}
		}
	}

	public wpfStyleScrubber()
	{
		base.Loaded += wpfStyleScrubber_Loaded;
		base.Closing += wpfStyleScrubber_Closing;
		this.m_A = 0;
		this.m_A = VH.A(49303) + ((_Application)MH.A.Application).get_International((object)XlApplicationInternational.xlThousandsSeparator).ToString() + VH.A(52500);
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = new BackgroundWorker();
		BackgroundWorker a = this.m_A;
		a.WorkerSupportsCancellation = true;
		a.WorkerReportsProgress = true;
		a.DoWork += bgw_DoWork;
		a.ProgressChanged += bgw_ProgressChanged;
		a.RunWorkerCompleted += bgw_RunWorkerCompleted;
		_ = null;
		this.m_A = MH.A.Application;
		A(tbCount, VH.A(103971) + this.m_A.ActiveWorkbook.Styles.Count.ToString(this.m_A) + VH.A(181139));
	}

	private void wpfStyleScrubber_Loaded(object sender, RoutedEventArgs e)
	{
		base.MinHeight = base.ActualHeight;
		base.MinWidth = base.ActualWidth;
	}

	private void wpfStyleScrubber_Closing(object sender, CancelEventArgs e)
	{
		this.m_A = null;
		this.m_B = null;
		this.m_A = null;
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		if (this.m_A.IsBusy)
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
			this.m_A.CancelAsync();
		}
		if (this.m_B != null && this.m_B.IsBusy)
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
			this.m_B.CancelAsync();
		}
		string left = btnCancel.Content.ToString();
		if (Operators.CompareString(left, VH.A(180569), TextCompare: false) != 0)
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
			if (Operators.CompareString(left, VH.A(180582), TextCompare: false) != 0)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						Operators.CompareString(left, VH.A(180593), TextCompare: false);
						return;
					}
				}
			}
		}
		Close();
	}

	private void btnStart_Click(object sender, RoutedEventArgs e)
	{
		lbxCorrupt.Items.Clear();
		A();
		this.m_A.ScreenUpdating = false;
		this.m_A.RunWorkerAsync();
	}

	private void A()
	{
		btnStart.IsEnabled = false;
		btnCancel.Content = VH.A(180593);
		btnCancel.IsCancel = false;
		pbProgress.Value = 0.0;
		tbStatus.Text = VH.A(180602);
		pbProgress.Visibility = Visibility.Visible;
		tbStatus.Visibility = Visibility.Visible;
		radAll.IsEnabled = false;
		radUnused.IsEnabled = false;
	}

	private void B()
	{
		pbProgress.Visibility = Visibility.Collapsed;
		tbStatus.Visibility = Visibility.Collapsed;
		btnStart.IsEnabled = true;
		radAll.IsEnabled = true;
		radUnused.IsEnabled = true;
		btnCancel.Content = VH.A(180582);
		btnCancel.IsCancel = true;
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		try
		{
			bool A = default(bool);
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				A = radAll.IsChecked.Value;
			});
			if (A)
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
						this.A(ref e);
						return;
					}
				}
			}
			B(ref e);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		this.m_A.ScreenUpdating = true;
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)8, VH.A(180625));
		B();
		if (e.Cancelled)
		{
			if (this.m_A == 0)
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
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
		}
		if (e.Cancelled)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					Forms.InfoMessage(VH.A(52374) + this.m_A.ToString(this.m_A) + VH.A(180654));
					return;
				}
			}
		}
		int count = lbxCorrupt.Items.Count;
		if (count == 0)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					Forms.InfoMessage(VH.A(52374) + this.m_A.ToString(this.m_A) + VH.A(180654));
					Close();
					return;
				}
			}
		}
		if (Forms.OkCancelMessage(VH.A(52374) + this.m_A.ToString(this.m_A) + VH.A(180689) + count.ToString(this.m_A) + VH.A(180768)) == System.Windows.Forms.DialogResult.OK)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
				{
					A();
					this.m_A.ScreenUpdating = false;
					this.m_B = new BackgroundWorker();
					BackgroundWorker b = this.m_B;
					b.WorkerSupportsCancellation = true;
					b.WorkerReportsProgress = true;
					b.DoWork += bgwCorrupt_DoWork;
					b.ProgressChanged += bgwCorrupt_ProgressChanged;
					b.RunWorkerCompleted += bgwCorrupt_RunWorkerCompleted;
					b.RunWorkerAsync();
					_ = null;
					return;
				}
				}
			}
		}
		Close();
	}

	private void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbProgress.Value = e.ProgressPercentage;
	}

	private void A(ref DoWorkEventArgs A)
	{
		Styles styles = this.m_A.ActiveWorkbook.Styles;
		int num = 0;
		int num2 = styles.Count;
		int B = num2;
		this.m_A = 0;
		int num3 = styles.Count;
		checked
		{
			while (true)
			{
				if (num3 >= 1)
				{
					if (this.m_A.CancellationPending)
					{
						A.Cancel = true;
						break;
					}
					if (!styles[num3].BuiltIn)
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
						this.A(styles[num3], ref B);
					}
					else
					{
						num2--;
					}
					num++;
					this.m_A.ReportProgress((int)Math.Round((double)num / (double)num2 * 100.0));
					num3 += -1;
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
				break;
			}
			styles = null;
		}
	}

	private void B(ref DoWorkEventArgs A)
	{
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = this.m_A.ActiveWorkbook;
		Styles styles = activeWorkbook.Styles;
		int num = styles.Count;
		int B = num;
		List<string> list = new List<string>();
		int count = activeWorkbook.Worksheets.Count;
		this.m_A = 0;
		int num2 = 1;
		IEnumerator enumerator = activeWorkbook.Worksheets.GetEnumerator();
		checked
		{
			try
			{
				IEnumerator enumerator2 = default(IEnumerator);
				IEnumerator enumerator3 = default(IEnumerator);
				while (true)
				{
					if (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)enumerator.Current;
						if (this.m_A.CancellationPending)
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
								A.Cancel = true;
								break;
							}
							break;
						}
						this.A(tbStatus, VH.A(180950) + num2 + VH.A(180981) + count + VH.A(116000));
						Range usedRange = worksheet.UsedRange;
						Microsoft.Office.Interop.Excel.Style style = null;
						try
						{
							style = (Microsoft.Office.Interop.Excel.Style)usedRange.Style;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						if (style != null)
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
							if (!style.BuiltIn)
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
								list.Add(style.Name);
							}
						}
						else
						{
							try
							{
								enumerator2 = usedRange.Rows.GetEnumerator();
								while (true)
								{
									if (enumerator2.MoveNext())
									{
										Range range = (Range)enumerator2.Current;
										if (this.m_A.CancellationPending)
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												A.Cancel = true;
												break;
											}
											break;
										}
										this.A(tbStatus, VH.A(180990) + range.Row + VH.A(181017) + num2 + VH.A(116000));
										style = null;
										try
										{
											style = (Microsoft.Office.Interop.Excel.Style)range.Style;
										}
										catch (Exception ex3)
										{
											ProjectData.SetProjectError(ex3);
											Exception ex4 = ex3;
											ProjectData.ClearProjectError();
										}
										if (style != null)
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
											if (style.BuiltIn)
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
											list.Add(style.Name);
											continue;
										}
										{
											enumerator3 = range.Cells.GetEnumerator();
											try
											{
												while (enumerator3.MoveNext())
												{
													Range range2 = (Range)enumerator3.Current;
													this.A(tbStatus, VH.A(181038) + range2.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)) + VH.A(181017) + num2 + VH.A(116000));
													try
													{
														style = (Microsoft.Office.Interop.Excel.Style)range2.Style;
														if (!style.BuiltIn)
														{
															list.Add(style.Name);
														}
													}
													catch (Exception ex5)
													{
														ProjectData.SetProjectError(ex5);
														Exception ex6 = ex5;
														ProjectData.ClearProjectError();
													}
												}
												while (true)
												{
													switch (4)
													{
													case 0:
														break;
													default:
														goto end_IL_0358;
													}
													continue;
													end_IL_0358:
													break;
												}
											}
											finally
											{
												IDisposable disposable2 = enumerator3 as IDisposable;
												if (disposable2 != null)
												{
													disposable2.Dispose();
												}
											}
										}
										continue;
									}
									while (true)
									{
										switch (6)
										{
										case 0:
											break;
										default:
											goto end_IL_0386;
										}
										continue;
										end_IL_0386:
										break;
									}
									break;
								}
							}
							finally
							{
								if (enumerator2 is IDisposable)
								{
									while (true)
									{
										switch (3)
										{
										case 0:
											continue;
										}
										(enumerator2 as IDisposable).Dispose();
										break;
									}
								}
							}
						}
						this.m_A.ReportProgress((int)Math.Round((double)num2 / (double)count * 100.0));
						num2++;
						continue;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_03e7;
						}
						continue;
						end_IL_03e7:
						break;
					}
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
			if (!A.Cancel)
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
				list = list.Distinct().ToList();
				this.A(tbStatus, VH.A(181067));
				num2 = 0;
				for (int i = styles.Count; i >= 1; i += -1)
				{
					if (this.m_A.CancellationPending)
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
						A.Cancel = true;
						break;
					}
					Microsoft.Office.Interop.Excel.Style style2 = styles[i];
					if (!style2.BuiltIn)
					{
						if (!list.Contains(style2.Name))
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
							this.A(style2, ref B);
						}
						else
						{
							num--;
						}
					}
					else
					{
						num--;
					}
					num2++;
					this.m_A.ReportProgress((int)Math.Round((double)num2 / (double)num * 100.0));
				}
			}
			activeWorkbook = null;
			styles = null;
		}
	}

	private void A(Microsoft.Office.Interop.Excel.Style A, ref int B)
	{
		bool flag = false;
		checked
		{
			try
			{
				Microsoft.Office.Interop.Excel.Style style = A;
				try
				{
					style.Locked = false;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				style.Delete();
				try
				{
					if (style.Name.Length > 0)
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
							flag = true;
							break;
						}
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					this.m_A++;
					this.A(tbStatus, VH.A(52374) + this.m_A.ToString(this.m_A) + VH.A(181118));
					B--;
					this.A(tbCount, VH.A(103971) + B.ToString(this.m_A) + VH.A(181139));
					ProjectData.ClearProjectError();
				}
				style = null;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				flag = true;
				ProjectData.ClearProjectError();
			}
			if (!flag)
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
				this.A(A.Name);
				return;
			}
		}
	}

	private void bgwCorrupt_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		this.m_A.ScreenUpdating = true;
		if (!e.Cancelled)
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
			if (!Conversions.ToBoolean(e.Result))
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						Close();
						return;
					}
				}
			}
		}
		B();
	}

	private void bgwCorrupt_ProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbProgress.Value = e.ProgressPercentage;
	}

	private void bgwCorrupt_DoWork(object sender, DoWorkEventArgs e)
	{
		C(ref e);
	}

	private void C(ref DoWorkEventArgs A)
	{
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = this.m_A.ActiveWorkbook;
		bool flag = false;
		if (activeWorkbook.Path.Length == 0)
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
					Forms.WarningMessage(VH.A(181154));
					activeWorkbook = null;
					return;
				}
			}
		}
		Regex regex = new Regex(VH.A(181233));
		string text = this.A(activeWorkbook, VH.A(181268));
		activeWorkbook.SaveCopyAs(text);
		SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(text, isEditable: true);
		checked
		{
			try
			{
				YG a = default(YG);
				YG CS_0024_003C_003E8__locals7 = new YG(a);
				CS_0024_003C_003E8__locals7.A = this;
				WorkbookStylesPart workbookStylesPart = spreadsheetDocument.WorkbookPart.WorkbookStylesPart;
				Stylesheet stylesheet = workbookStylesPart.Stylesheet;
				CellStyles cellStyles = stylesheet.CellStyles;
				int num = (int)cellStyles.Count.Value - 1;
				base.Dispatcher.Invoke([SpecialName] () =>
				{
					CS_0024_003C_003E8__locals7.A = CS_0024_003C_003E8__locals7.A.lbxCorrupt.Items.Count;
					CS_0024_003C_003E8__locals7.A = CS_0024_003C_003E8__locals7.A.radAll.IsChecked.Value;
				});
				int num2 = 0;
				CellStyle cellStyle;
				try
				{
					ZG zG = default(ZG);
					for (int num3 = CS_0024_003C_003E8__locals7.A - 1; num3 >= 0; num3 += -1)
					{
						zG = new ZG(zG);
						zG.A = CS_0024_003C_003E8__locals7;
						zG.A = num3;
						if (this.m_B.CancellationPending)
						{
							A.Cancel = true;
							break;
						}
						base.Dispatcher.Invoke(zG.A);
						int num4 = num;
						while (true)
						{
							if (num4 >= 0)
							{
								cellStyle = (CellStyle)cellStyles.ElementAt(num4);
								bool flag2;
								try
								{
									flag2 = cellStyle.BuiltinId.HasValue;
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									flag2 = false;
									ProjectData.ClearProjectError();
								}
								if (Operators.CompareString(cellStyle.Name.Value, zG.A.A, TextCompare: false) != 0 && !regex.IsMatch(cellStyle.Name.Value))
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
									if (!zG.A.A || flag2)
									{
										num4 += -1;
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
								}
								cellStyle.Remove();
								num--;
								break;
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
							break;
						}
						this.A(num3);
						num2++;
						this.m_B.ReportProgress((int)Math.Round((double)num2 / (double)zG.A.A * 100.0));
						this.A(tbStatus, VH.A(52374) + num2 + VH.A(181285));
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					Forms.ErrorMessage(ex4.Message);
					clsReporting.LogException(ex4);
					flag = true;
					ProjectData.ClearProjectError();
				}
				cellStyle = null;
				if (!flag)
				{
					if (workbookStylesPart.Stylesheet.Descendants<StylesheetExtensionList>().Any())
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
						workbookStylesPart.Stylesheet.RemoveAllChildren<StylesheetExtensionList>();
					}
					stylesheet.Save();
					spreadsheetDocument.WorkbookPart.Workbook.Save();
				}
			}
			finally
			{
				if (spreadsheetDocument != null)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						((IDisposable)spreadsheetDocument).Dispose();
						break;
					}
				}
			}
			if (!flag)
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
				try
				{
					activeWorkbook = this.m_A.Workbooks.Open(text, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					activeWorkbook.Activate();
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
			}
			regex = null;
			activeWorkbook = null;
			A.Result = flag;
		}
	}

	private void A(TextBlock A, string B)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			A.Text = B;
		});
	}

	private void A(string A)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			lbxCorrupt.Items.Add(A);
		});
	}

	private void A(int A)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			lbxCorrupt.Items.RemoveAt(A);
		});
	}

	private string A(Microsoft.Office.Interop.Excel.Workbook A, string B)
	{
		return Path.Combine(A.Path, Path.GetFileNameWithoutExtension(A.Name) + B + Path.GetExtension(A.Name));
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_A)
		{
			this.m_A = true;
			Uri resourceLocator = new Uri(VH.A(181322), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					tbCount = (TextBlock)target;
					return;
				}
			}
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
					radAll = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			radUnused = (System.Windows.Controls.RadioButton)target;
			return;
		}
		if (connectionId == 4)
		{
			lbxCorrupt = (System.Windows.Controls.ListBox)target;
			return;
		}
		if (connectionId == 5)
		{
			tbStatus = (TextBlock)target;
			return;
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
					pbProgress = (System.Windows.Controls.ProgressBar)target;
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
					btnStart = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (2)
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
}
