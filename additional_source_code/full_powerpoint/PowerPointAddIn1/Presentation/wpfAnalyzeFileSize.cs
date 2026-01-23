using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Presentation;

[DesignerGenerated]
public sealed class wpfAnalyzeFileSize : Window, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class OF
	{
		public ObservableCollection<SlideSize> A;

		public wpfAnalyzeFileSize A;

		public OF(OF A)
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

		[SpecialName]
		internal void A()
		{
			this.A.lbxSize.ItemsSource = this.A;
		}
	}

	private BackgroundWorker m_A;

	private long m_A;

	private bool m_A;

	[AccessedThroughProperty("lbxSize")]
	[CompilerGenerated]
	private ListBox m_A;

	[AccessedThroughProperty("pbLoading")]
	[CompilerGenerated]
	private ProgressBar m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private Button m_A;

	private bool B;

	internal virtual ListBox lbxSize
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
			SelectionChangedEventHandler value2 = lbxSize_SelectionChanged;
			SizeChangedEventHandler value3 = lbxSize_SizeChanged;
			ListBox listBox = this.m_A;
			if (listBox != null)
			{
				listBox.SelectionChanged -= value2;
				listBox.SizeChanged -= value3;
			}
			this.m_A = value;
			listBox = this.m_A;
			if (listBox == null)
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
				listBox.SelectionChanged += value2;
				listBox.SizeChanged += value3;
				return;
			}
		}
	}

	internal virtual ProgressBar pbLoading
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

	internal virtual Button btnClose
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
			RoutedEventHandler value2 = btnClose_Click;
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
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	public wpfAnalyzeFileSize()
	{
		base.Loaded += wpfAnalyzeFileSize_Loaded;
		base.Closing += frmAnalyzeFileSize_FormClosing;
		this.m_A = 0L;
		this.m_A = false;
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
	}

	private void wpfAnalyzeFileSize_Loaded(object sender, RoutedEventArgs e)
	{
		pbLoading.Value = 0.0;
		pbLoading.Visibility = Visibility.Visible;
		this.m_A.RunWorkerAsync();
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		OF a = default(OF);
		OF CS_0024_003C_003E8__locals6 = new OF(a);
		CS_0024_003C_003E8__locals6.A = this;
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
		CS_0024_003C_003E8__locals6.A = new ObservableCollection<SlideSize>();
		Dictionary<int, long> dictionary = new Dictionary<int, long>();
		checked
		{
			try
			{
				string text = Path.Combine(Path.GetTempPath(), AH.A(118496));
				try
				{
					string text2 = text + AH.A(118527);
					Directory.Move(text, text2);
					Directory.Delete(text2, recursive: true);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				Directory.CreateDirectory(text);
				string text3 = Path.Combine(text, AH.A(118544));
				activePresentation.SaveCopyAs(text3);
				Microsoft.Office.Interop.PowerPoint.Presentation presentation = application.Presentations.Open(text3, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
				for (int i = presentation.Slides.Count; i >= 1; i += -1)
				{
					presentation.Slides[i].Delete();
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
					presentation.Save();
					presentation.Close();
					presentation = null;
					long num = A(text3);
					File.Delete(text3);
					int count = activePresentation.Slides.Count;
					enumerator = activePresentation.Slides.GetEnumerator();
					try
					{
						while (true)
						{
							if (enumerator.MoveNext())
							{
								Slide slide = (Slide)enumerator.Current;
								if (this.m_A.CancellationPending)
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
									this.m_A.ReportProgress((int)Math.Round((double)slide.SlideIndex / (double)count * 100.0));
									activePresentation.SaveCopyAs(text3);
									Microsoft.Office.Interop.PowerPoint.Presentation presentation2 = application.Presentations.Open(text3, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
									Slide slide2 = presentation2.Slides[slide.SlideIndex];
									for (int j = presentation2.Slides.Count; j >= 1; j += -1)
									{
										if (presentation2.Slides[j] == slide2)
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
										presentation2.Slides[j].Delete();
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_0245;
										}
										continue;
										end_IL_0245:
										break;
									}
									presentation2.Save();
									presentation2.Close();
									presentation2 = null;
									long num2 = A(text3);
									dictionary.Add(slide.SlideIndex, num2);
									this.m_A = Math.Max(this.m_A, num2);
									File.Delete(text3);
									break;
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
									goto end_IL_02a2;
								}
								continue;
								end_IL_02a2:
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
					double b = A();
					using (Dictionary<int, long>.Enumerator enumerator2 = dictionary.GetEnumerator())
					{
						while (enumerator2.MoveNext())
						{
							KeyValuePair<int, long> current = enumerator2.Current;
							long num2 = current.Value;
							string text4;
							if (num2 - num < 1000000)
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
								text4 = ((double)(num2 - num) / 1000.0).ToString(AH.A(118563)) + AH.A(118578);
							}
							else
							{
								text4 = ((double)(num2 - num) / 1000000.0).ToString(AH.A(118563)) + AH.A(118583);
							}
							CS_0024_003C_003E8__locals6.A.Add(new SlideSize
							{
								Label = AH.A(36272) + current.Key + AH.A(17804) + text4 + AH.A(14255),
								BarWidth = A(num2, b),
								Size = num2
							});
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_0416;
							}
							continue;
							end_IL_0416:
							break;
						}
					}
					Directory.Delete(text, recursive: true);
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				CS_0024_003C_003E8__locals6.A.lbxSize.ItemsSource = CS_0024_003C_003E8__locals6.A;
			});
			activePresentation = null;
			application = null;
			dictionary = null;
			CS_0024_003C_003E8__locals6.A = null;
		}
	}

	private long A(string A)
	{
		return NB.A.FileSystem.GetFileInfo(A).Length;
	}

	private double A()
	{
		return lbxSize.ActualWidth - 130.0;
	}

	private double A(long A, double B)
	{
		return (double)A / (double)this.m_A * B;
	}

	private void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbLoading.Value = e.ProgressPercentage;
	}

	private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		pbLoading.Visibility = Visibility.Hidden;
		base.SizeToContent = SizeToContent.Manual;
		this.m_A = true;
	}

	private void lbxSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		NG.A.Application.ActiveWindow.View.GotoSlide(checked(lbxSize.SelectedIndex + 1));
	}

	private void lbxSize_SizeChanged(object sender, SizeChangedEventArgs e)
	{
		if (!this.m_A)
		{
			return;
		}
		IEnumerator<SlideSize> enumerator = default(IEnumerator<SlideSize>);
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
			if (!e.WidthChanged)
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
				double b = A();
				ObservableCollection<SlideSize> observableCollection = (ObservableCollection<SlideSize>)lbxSize.ItemsSource;
				try
				{
					enumerator = observableCollection.GetEnumerator();
					while (enumerator.MoveNext())
					{
						SlideSize current = enumerator.Current;
						current.BarWidth = A(current.Size, b);
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (1)
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
		}
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private void frmAnalyzeFileSize_FormClosing(object sender, CancelEventArgs e)
	{
		this.m_A = null;
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (!B)
		{
			B = true;
			Uri resourceLocator = new Uri(AH.A(118588), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
		if (connectionId == 1)
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
					lbxSize = (ListBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 2:
			pbLoading = (ProgressBar)target;
			break;
		case 3:
			while (true)
			{
				switch (5)
				{
				case 0:
					continue;
				}
				btnClose = (Button)target;
				return;
			}
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
