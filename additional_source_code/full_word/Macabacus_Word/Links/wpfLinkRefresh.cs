using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Threading;
using A;
using MacabacusMacros;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Macabacus_Word.Shapes;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

[DesignerGenerated]
public sealed class wpfLinkRefresh : System.Windows.Window, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class BB
	{
		public RefreshInstance A;

		public bool A;

		public wpfLinkRefresh A;

		public BB(BB A)
		{
			if (A == null)
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
				this.A = A.A;
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			//IL_000e: Unknown result type (might be due to invalid IL or missing references)
			//IL_0018: Expected O, but got Unknown
			this.A = new RefreshInstance(System.Windows.Window.GetWindow(this.A));
		}
	}

	[CompilerGenerated]
	internal sealed class CB
	{
		public Microsoft.Office.Interop.Word.ContentControl A;

		public DB A;

		public CB(CB A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			Refresh.UpdateShapeLink(this.A, ref this.A.A.A, ref this.A.A.A, this.A.A);
		}
	}

	[CompilerGenerated]
	internal sealed class DB
	{
		public CopierAsPicture A;

		public BB A;

		public DB(DB A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class EB
	{
		public InlineShape A;

		public DB A;

		public EB(EB A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			Refresh.UpdateShapeLink(this.A, ref this.A.A.A, ref this.A.A.A, this.A.A);
		}
	}

	[CompilerGenerated]
	internal sealed class FB
	{
		public Shape A;

		public DB A;

		public FB(FB A)
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
			Refresh.UpdateShapeLink(this.A, ref this.A.A.A, ref this.A.A.A, this.A.A);
		}
	}

	[CompilerGenerated]
	internal sealed class GB
	{
		public Table A;

		public DB A;

		public GB(GB A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			Refresh.UpdateShapeLink(this.A, ref this.A.A.A, ref this.A.A.A, this.A.A);
		}
	}

	[CompilerGenerated]
	internal sealed class HB
	{
		public int A;

		public int B;

		public int C;

		public wpfLinkRefresh A;

		[SpecialName]
		internal void A()
		{
			this.A.tbStatus.Text = XC.A(44108) + this.A + XC.A(13138) + B + XC.A(20691) + string.Format(XC.A(44141), (double)C / 100.0) + XC.A(44154);
		}
	}

	private Microsoft.Office.Interop.Word.Application m_A;

	private BackgroundWorker m_A;

	private Dictionary<object, string> m_A;

	private int m_A;

	private List<InlineShape> m_A;

	private List<Shape> m_A;

	private List<Table> m_A;

	private List<Microsoft.Office.Interop.Word.ContentControl> m_A;

	private bool m_A;

	private bool m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("pbLink")]
	private ProgressBar m_A;

	[AccessedThroughProperty("tbStatus")]
	[CompilerGenerated]
	private TextBlock m_A;

	private bool m_C;

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
			MouseEventHandler value2 = btnCancel_MouseEnter;
			MouseEventHandler value3 = btnCancel_MouseLeave;
			RoutedEventHandler value4 = btnCancel_Click;
			Button button = this.m_A;
			if (button != null)
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
				button.MouseEnter -= value2;
				button.MouseLeave -= value3;
				button.Click -= value4;
			}
			this.m_A = value;
			button = this.m_A;
			if (button != null)
			{
				button.MouseEnter += value2;
				button.MouseLeave += value3;
				button.Click += value4;
			}
		}
	}

	internal virtual ProgressBar pbLink
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
			return this.m_A;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	public wpfLinkRefresh(List<InlineShape> listInlineShapes, List<Shape> listShapes, List<Table> listTables, List<Microsoft.Office.Interop.Word.ContentControl> listContentControls, bool blnAll)
	{
		base.Loaded += wpfLinkRefresh_Loaded;
		base.Closing += wpfLinkRefresh_Closing;
		base.MouseDown += Window_MouseDown;
		InitializeComponent();
		this.m_A = listInlineShapes;
		this.m_A = listShapes;
		this.m_A = listTables;
		this.m_A = listContentControls;
		this.m_A = blnAll;
		this.m_A = PC.A.Application;
		this.m_A = new Dictionary<object, string>();
		this.m_A = new BackgroundWorker();
		BackgroundWorker a = this.m_A;
		a.WorkerSupportsCancellation = true;
		a.WorkerReportsProgress = true;
		a.DoWork += RefreshDoWork;
		a.ProgressChanged += RefreshProgressChanged;
		a.RunWorkerCompleted += RefreshCompleted;
		_ = null;
		this.m_B = Navigate.A(this.m_A);
	}

	private void wpfLinkRefresh_Loaded(object sender, RoutedEventArgs e)
	{
		tbStatus.Text = XC.A(14849);
		pbLink.IsIndeterminate = true;
		this.m_A.ScreenUpdating = false;
		this.m_A.RunWorkerAsync();
	}

	private void wpfLinkRefresh_Closing(object sender, CancelEventArgs e)
	{
		if (this.m_A != null)
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
			if (this.m_A.IsBusy)
			{
				this.m_A.CancelAsync();
				e.Cancel = true;
			}
		}
		if (e.Cancel)
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
			this.m_A = null;
			this.m_A = null;
			ReleaseHelper.ClearDictKeyReferences<object, string>(ref this.m_A, false, (Action<object>)null);
			ReleaseHelper.ClearListReferences<InlineShape>(ref this.m_A, false, (Action<InlineShape>)null);
			ReleaseHelper.ClearListReferences<Shape>(ref this.m_A, false, (Action<Shape>)null);
			ReleaseHelper.ClearListReferences<Table>(ref this.m_A, false, (Action<Table>)null);
			ReleaseHelper.ClearListReferences<Microsoft.Office.Interop.Word.ContentControl>(ref this.m_A, false, (Action<Microsoft.Office.Interop.Word.ContentControl>)null);
			ReleaseHelper.DoGarbageCollection();
			return;
		}
	}

	private void btnCancel_MouseEnter(object sender, MouseEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			btnCancel.Opacity = 1.0;
		}, DispatcherPriority.Background);
	}

	private void btnCancel_MouseLeave(object sender, MouseEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			btnCancel.Opacity = 0.6;
		}, DispatcherPriority.Background);
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			if (!this.m_A.IsBusy)
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
				this.m_A.CancelAsync();
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void Window_MouseDown(object sender, MouseButtonEventArgs e)
	{
		if (e.ChangedButton == MouseButton.Left)
		{
			DragMove();
		}
	}

	private void RefreshDoWork(object sender, DoWorkEventArgs e)
	{
		//IL_003c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0042: Expected O, but got Unknown
		//IL_0043: Expected O, but got Unknown
		//IL_0557: Unknown result type (might be due to invalid IL or missing references)
		//IL_055d: Expected O, but got Unknown
		//IL_055f: Expected O, but got Unknown
		//IL_06a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_06ac: Expected O, but got Unknown
		//IL_06ae: Expected O, but got Unknown
		//IL_0829: Unknown result type (might be due to invalid IL or missing references)
		//IL_082f: Expected O, but got Unknown
		//IL_0831: Expected O, but got Unknown
		//IL_09a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_09ac: Expected O, but got Unknown
		//IL_09ae: Expected O, but got Unknown
		//IL_0416: Unknown result type (might be due to invalid IL or missing references)
		//IL_0420: Expected O, but got Unknown
		BB a = default(BB);
		BB CS_0024_003C_003E8__locals6 = new BB(a);
		CS_0024_003C_003E8__locals6.A = this;
		CS_0024_003C_003E8__locals6.A = null;
		UndoRecord undoRecord = this.m_A.UndoRecord;
		int num = 0;
		try
		{
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				//IL_000e: Unknown result type (might be due to invalid IL or missing references)
				//IL_0018: Expected O, but got Unknown
				CS_0024_003C_003E8__locals6.A = new RefreshInstance(System.Windows.Window.GetWindow(CS_0024_003C_003E8__locals6.A));
			});
		}
		catch (UpdateLinkException ex)
		{
			ProjectData.SetProjectError((Exception)ex);
			UpdateLinkException ex2 = ex;
			MessageBox.Show(((Exception)(object)ex2).Message, XC.A(2438), MessageBoxButton.OK, MessageBoxImage.Exclamation);
			e.Cancel = true;
			ProjectData.ClearProjectError();
		}
		checked
		{
			try
			{
				for (int num2 = this.m_A.Count - 1; num2 >= 0; num2 += -1)
				{
					if (this.m_A == null)
					{
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (this.m_A.CancellationPending)
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
						e.Cancel = true;
						break;
					}
					if (!Common.IsLinked(this.m_A[num2]))
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
						this.m_A.RemoveAt(num2);
					}
					else
					{
						if (!A(this.m_A[num2]))
						{
							continue;
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
						this.m_A.RemoveAt(num2);
					}
				}
				if (!e.Cancel)
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
					int num3 = this.m_A.Count - 1;
					while (true)
					{
						if (num3 >= 0)
						{
							if (this.m_A == null)
							{
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
							if (this.m_A.CancellationPending)
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
								e.Cancel = true;
								break;
							}
							if (!Common.IsLinked(this.m_A[num3]))
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
								this.m_A.RemoveAt(num3);
							}
							else if (A(this.m_A[num3]))
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
								this.m_A.RemoveAt(num3);
							}
							num3 += -1;
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
						break;
					}
				}
				if (!e.Cancel)
				{
					int num4 = this.m_A.Count - 1;
					while (true)
					{
						if (num4 >= 0)
						{
							if (this.m_A == null)
							{
								break;
							}
							if (this.m_A.CancellationPending)
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
								e.Cancel = true;
								break;
							}
							if (!Common.IsLinked(this.m_A[num4]))
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
								this.m_A.RemoveAt(num4);
							}
							else if (A(this.m_A[num4]))
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
								this.m_A.RemoveAt(num4);
							}
							num4 += -1;
							continue;
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
						break;
					}
				}
				if (!e.Cancel)
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
					int num5 = this.m_A.Count - 1;
					while (true)
					{
						if (num5 >= 0)
						{
							if (this.m_A == null)
							{
								break;
							}
							if (this.m_A.CancellationPending)
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
								e.Cancel = true;
								break;
							}
							if (!Common.IsLinked(this.m_A[num5]))
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
								this.m_A.RemoveAt(num5);
							}
							else if (A(this.m_A[num5]))
							{
								this.m_A.RemoveAt(num5);
							}
							num5 += -1;
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
				}
				this.m_A = this.m_A.Count + this.m_A.Count + this.m_A.Count + this.m_A.Count;
				if (this.m_A <= 0)
				{
					return;
				}
				DB dB = default(DB);
				CB cB = default(CB);
				EB eB = default(EB);
				FB fB = default(FB);
				GB gB = default(GB);
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					if (e.Cancel)
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
						dB = new DB(dB);
						dB.A = CS_0024_003C_003E8__locals6;
						dB.A = new CopierAsPicture();
						dB.A.A = this.m_A == 1;
						undoRecord.StartCustomRecord(XC.A(14892));
						this.m_A.ReportProgress(0);
						base.Dispatcher.Invoke([SpecialName] () =>
						{
							pbLink.IsIndeterminate = false;
						}, DispatcherPriority.Background);
						if (!e.Cancel && !dB.A.A.Canceled)
						{
							using List<Microsoft.Office.Interop.Word.ContentControl>.Enumerator enumerator = this.m_A.GetEnumerator();
							while (true)
							{
								if (enumerator.MoveNext())
								{
									cB = new CB(cB);
									cB.A = dB;
									cB.A = enumerator.Current;
									if (this.m_A == null)
									{
										break;
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
									if (this.m_A.CancellationPending)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											e.Cancel = true;
											break;
										}
										break;
									}
									if (cB.A.A.A.Canceled)
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												break;
											default:
												goto end_IL_051a;
											}
											continue;
											end_IL_051a:
											break;
										}
										break;
									}
									num++;
									A(num, this.m_A);
									try
									{
										base.Dispatcher.Invoke(cB.A, DispatcherPriority.Background);
									}
									catch (UpdateLinkException ex3)
									{
										ProjectData.SetProjectError((Exception)ex3);
										UpdateLinkException ex4 = ex3;
										this.m_A.Add(cB.A, ((Exception)(object)ex4).Message);
										ProjectData.ClearProjectError();
									}
									catch (Exception ex5)
									{
										ProjectData.SetProjectError(ex5);
										Exception ex6 = ex5;
										this.m_A.Add(cB.A, ex6.Message);
										if (!this.m_A)
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
											clsReporting.LogException(ex6);
										}
										ProjectData.ClearProjectError();
									}
									continue;
								}
								while (true)
								{
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_05d2;
									}
									continue;
									end_IL_05d2:
									break;
								}
								break;
							}
						}
						using (List<InlineShape>.Enumerator enumerator2 = this.m_A.GetEnumerator())
						{
							while (enumerator2.MoveNext())
							{
								eB = new EB(eB);
								eB.A = dB;
								eB.A = enumerator2.Current;
								if (this.m_A == null)
								{
									break;
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
								if (this.m_A.CancellationPending)
								{
									e.Cancel = true;
									break;
								}
								if (eB.A.A.A.Canceled)
								{
									while (true)
									{
										switch (4)
										{
										case 0:
											break;
										default:
											goto end_IL_066b;
										}
										continue;
										end_IL_066b:
										break;
									}
									break;
								}
								num++;
								A(num, this.m_A);
								try
								{
									base.Dispatcher.Invoke(eB.A, DispatcherPriority.Background);
								}
								catch (UpdateLinkException ex7)
								{
									ProjectData.SetProjectError((Exception)ex7);
									UpdateLinkException ex8 = ex7;
									this.m_A.Add(eB.A, ((Exception)(object)ex8).Message);
									ProjectData.ClearProjectError();
								}
								catch (Exception ex9)
								{
									ProjectData.SetProjectError(ex9);
									Exception ex10 = ex9;
									this.m_A.Add(eB.A, ex10.Message);
									if (!this.m_A)
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
										clsReporting.LogException(ex10);
									}
									ProjectData.ClearProjectError();
								}
							}
						}
						if (!e.Cancel)
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
							if (!dB.A.A.Canceled)
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
								using List<Shape>.Enumerator enumerator3 = this.m_A.GetEnumerator();
								while (enumerator3.MoveNext())
								{
									fB = new FB(fB);
									fB.A = dB;
									fB.A = enumerator3.Current;
									if (this.m_A == null)
									{
										break;
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
									if (this.m_A.CancellationPending)
									{
										while (true)
										{
											switch (4)
											{
											case 0:
												continue;
											}
											e.Cancel = true;
											break;
										}
										break;
									}
									if (fB.A.A.A.Canceled)
									{
										while (true)
										{
											switch (4)
											{
											case 0:
												break;
											default:
												goto end_IL_07ec;
											}
											continue;
											end_IL_07ec:
											break;
										}
										break;
									}
									num++;
									A(num, this.m_A);
									try
									{
										base.Dispatcher.Invoke(fB.A, DispatcherPriority.Background);
									}
									catch (UpdateLinkException ex11)
									{
										ProjectData.SetProjectError((Exception)ex11);
										UpdateLinkException ex12 = ex11;
										this.m_A.Add(fB.A, ((Exception)(object)ex12).Message);
										ProjectData.ClearProjectError();
									}
									catch (Exception ex13)
									{
										ProjectData.SetProjectError(ex13);
										Exception ex14 = ex13;
										this.m_A.Add(fB.A, ex14.Message);
										if (!this.m_A)
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
											clsReporting.LogException(ex14);
										}
										ProjectData.ClearProjectError();
									}
								}
							}
						}
						if (!e.Cancel)
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
							if (!dB.A.A.Canceled)
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
								using List<Table>.Enumerator enumerator4 = this.m_A.GetEnumerator();
								while (enumerator4.MoveNext())
								{
									gB = new GB(gB);
									gB.A = dB;
									gB.A = enumerator4.Current;
									if (this.m_A == null)
									{
										break;
									}
									if (this.m_A.CancellationPending)
									{
										while (true)
										{
											switch (7)
											{
											case 0:
												continue;
											}
											e.Cancel = true;
											break;
										}
										break;
									}
									if (gB.A.A.A.Canceled)
									{
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
												goto end_IL_0969;
											}
											continue;
											end_IL_0969:
											break;
										}
										break;
									}
									num++;
									A(num, this.m_A);
									try
									{
										base.Dispatcher.Invoke(gB.A, DispatcherPriority.Background);
									}
									catch (UpdateLinkException ex15)
									{
										ProjectData.SetProjectError((Exception)ex15);
										UpdateLinkException ex16 = ex15;
										this.m_A.Add(gB.A, ((Exception)(object)ex16).Message);
										ProjectData.ClearProjectError();
									}
									catch (Exception ex17)
									{
										ProjectData.SetProjectError(ex17);
										Exception ex18 = ex17;
										this.m_A.Add(gB.A, ex18.Message);
										if (!this.m_A)
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
											clsReporting.LogException(ex18);
										}
										ProjectData.ClearProjectError();
									}
								}
							}
						}
						undoRecord.EndCustomRecord();
						undoRecord = null;
						if (e.Cancel)
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
							if (!dB.A.A.Canceled)
							{
								B();
								Thread.Sleep(300);
							}
							return;
						}
					}
				}
			}
			finally
			{
				Base.ReleaseRefreshInstance(ref CS_0024_003C_003E8__locals6.A, true);
			}
		}
	}

	private bool A(object A)
	{
		Type typeFromHandle = typeof(Common);
		string memberName = XC.A(11777);
		object[] obj = new object[1] { A };
		object[] array = obj;
		bool[] obj2 = new bool[1] { true };
		bool[] array2 = obj2;
		object instance = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
		if (array2[0])
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
			A = RuntimeHelpers.GetObjectValue(array[0]);
		}
		return Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(instance, null, XC.A(13872), new object[0], null, null, null), string.Empty, TextCompare: false);
	}

	private void RefreshProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbLink.Value = e.ProgressPercentage;
	}

	private void RefreshCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
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
			try
			{
				if (Navigate.A(this.m_A))
				{
					Microsoft.Office.Interop.Word.View view = this.m_A.ActiveWindow.View;
					view.SeekView = WdSeekView.wdSeekCurrentPageFooter;
					view.SeekView = WdSeekView.wdSeekMainDocument;
					_ = null;
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		this.m_A.ScreenUpdating = true;
		ExcelToWord.ActivateWord(this.m_A);
		base.Topmost = false;
		Hide();
		if (e.Cancelled)
		{
			C(XC.A(1663));
		}
		else if (!this.m_A.Any())
		{
			if (this.m_A > 0)
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
				if (this.m_A)
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
						A(XC.A(14917));
					}
				}
			}
			else if (!this.m_A)
			{
				B(XC.A(13338));
			}
			else
			{
				B(XC.A(14984));
			}
		}
		else if (this.m_A == 1)
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
			D(this.m_A.Values.ElementAtOrDefault(0));
		}
		else if (MessageBox.Show(XC.A(15049), XC.A(2438), MessageBoxButton.YesNo, MessageBoxImage.Hand) == MessageBoxResult.Yes)
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
			List<LinkError> list = new List<LinkError>();
			int num = 1;
			using (Dictionary<object, string>.Enumerator enumerator = this.m_A.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					KeyValuePair<object, string> current = enumerator.Current;
					list.Add(new LinkError(RuntimeHelpers.GetObjectValue(current.Key), num.ToString(), current.Value));
					num = checked(num + 1);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						goto end_IL_01f7;
					}
					continue;
					end_IL_01f7:
					break;
				}
			}
			wpfLinkUpdateErrors obj = new wpfLinkUpdateErrors();
			obj.lvErrors.ItemsSource = list;
			obj.Show();
			_ = null;
			list = null;
		}
		Common.LogActivity(XC.A(15201));
		Close();
	}

	private void A(int A, int B)
	{
		int C = checked((int)Math.Round((double)(A - 1) / (double)this.m_A * 100.0));
		this.m_A.ReportProgress(C);
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			tbStatus.Text = XC.A(44108) + A + XC.A(13138) + B + XC.A(20691) + string.Format(XC.A(44141), (double)C / 100.0) + XC.A(44154);
		}, DispatcherPriority.Background);
	}

	private void B()
	{
		this.m_A.ReportProgress(100);
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			tbStatus.Text = XC.A(15333);
		}, DispatcherPriority.Background);
	}

	private void A(string A)
	{
		Forms.SuccessMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void B(string A)
	{
		Forms.InfoMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
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
			switch (1)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_C = true;
			Uri resourceLocator = new Uri(XC.A(15226), UriKind.Relative);
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
					btnCancel = (Button)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 2:
			pbLink = (ProgressBar)target;
			break;
		case 3:
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				tbStatus = (TextBlock)target;
				return;
			}
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

	[SpecialName]
	[CompilerGenerated]
	private void C()
	{
		btnCancel.Opacity = 1.0;
	}

	[SpecialName]
	[CompilerGenerated]
	private void D()
	{
		btnCancel.Opacity = 0.6;
	}

	[SpecialName]
	[CompilerGenerated]
	private void E()
	{
		pbLink.IsIndeterminate = false;
	}

	[SpecialName]
	[CompilerGenerated]
	private void F()
	{
		tbStatus.Text = XC.A(15333);
	}
}
