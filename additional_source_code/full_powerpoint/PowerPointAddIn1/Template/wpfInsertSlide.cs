using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Library2;
using PowerPointAddIn1.Presentation;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Template;

[DesignerGenerated]
public sealed class wpfInsertSlide : Window, INotifyPropertyChanged, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class PF
	{
		public CustomLayout A;

		public wpfInsertSlide A;

		public PF(PF A)
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
	}

	[CompilerGenerated]
	internal sealed class QF
	{
		public System.Drawing.Image A;

		public PF A;

		public QF(QF A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.A.LayoutItems.Add(new LayoutThumb(this.A.A.Name, Forms.GetImageSource((Bitmap)this.A)));
		}
	}

	[CompilerGenerated]
	internal sealed class RF
	{
		public Exception A;

		public PF A;

		public RF(RF A)
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
			this.A.A.B(AH.A(170874) + this.A.Message);
		}
	}

	[CompilerGenerated]
	internal sealed class SF
	{
		public Exception A;

		public wpfInsertSlide A;

		public SF(SF A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.B(this.A.Message);
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private string m_A;

	private BackgroundWorker m_A;

	private static ObservableCollection<LayoutThumb> m_A;

	[AccessedThroughProperty("dckError")]
	[CompilerGenerated]
	private DockPanel m_A;

	[AccessedThroughProperty("tbNoTemplate")]
	[CompilerGenerated]
	private TextBlock m_A;

	[AccessedThroughProperty("lbxThumbs")]
	[CompilerGenerated]
	private ListBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("pbLoading")]
	private ProgressBar m_A;

	[AccessedThroughProperty("btnInsert")]
	[CompilerGenerated]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private Button m_B;

	private bool m_A;

	public ObservableCollection<LayoutThumb> LayoutItems
	{
		get
		{
			return wpfInsertSlide.m_A;
		}
		set
		{
			wpfInsertSlide.m_A = value;
			A(AH.A(134171));
		}
	}

	internal virtual DockPanel dckError
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

	internal virtual TextBlock tbNoTemplate
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

	internal virtual ListBox lbxThumbs
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
			SelectionChangedEventHandler value2 = lbxThumbs_SelectionChanged;
			ListBox listBox = this.m_A;
			if (listBox != null)
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
				listBox.SelectionChanged -= value2;
			}
			this.m_A = value;
			listBox = this.m_A;
			if (listBox == null)
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
				listBox.SelectionChanged += value2;
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

	internal virtual Button btnInsert
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
			RoutedEventHandler value2 = btnInsert_Click;
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
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnClose_Click;
			Button button = this.m_B;
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

	public event PropertyChangedEventHandler PropertyChanged
	{
		[CompilerGenerated]
		add
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Combine(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
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
				return;
			}
		}
		[CompilerGenerated]
		remove
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Remove(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
		}
	}

	public wpfInsertSlide()
	{
		base.Loaded += wpfInsertSlide_Loaded;
		base.Closing += wpfInsertSlide_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = "";
		dckError.Visibility = Visibility.Hidden;
	}

	private void A(string A)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(A));
			return;
		}
	}

	private void wpfInsertSlide_Loaded(object sender, RoutedEventArgs e)
	{
		pbLoading.IsIndeterminate = true;
		pbLoading.Visibility = Visibility.Visible;
		this.m_A = new BackgroundWorker();
		BackgroundWorker a = this.m_A;
		a.WorkerSupportsCancellation = true;
		a.WorkerReportsProgress = true;
		a.DoWork += bgw_DoWork;
		a.ProgressChanged += bgw_ProgressChanged;
		a.RunWorkerCompleted += bgw_RunWorkerCompleted;
		_ = null;
		this.m_A.RunWorkerAsync();
	}

	private void wpfInsertSlide_Closing(object sender, CancelEventArgs e)
	{
		try
		{
			this.m_A.CancelAsync();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		this.m_A = null;
		LayoutItems = null;
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation A = null;
		bool B = false;
		this.A(ref A, ref B);
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			PF pF = default(PF);
			QF qF = default(QF);
			IEnumerator enumerator2 = default(IEnumerator);
			RF rF = default(RF);
			SF a3 = default(SF);
			if (A != null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						this.m_A = A.FullName;
						Dictionary<string, ObservableCollection<LayoutThumb>> layoutThumbnails = InsertSlide.LayoutThumbnails;
						string a = this.m_A;
						ObservableCollection<LayoutThumb> value = LayoutItems;
						bool num = layoutThumbnails.TryGetValue(a, out value);
						LayoutItems = value;
						if (!num)
						{
							LayoutItems = new ObservableCollection<LayoutThumb>();
							base.Dispatcher.Invoke([SpecialName] () =>
							{
								pbLoading.IsIndeterminate = false;
								pbLoading.Value = 0.0;
							});
							Pen pen = new Pen(Color.FromArgb(200, 200, 200), 0.5f);
							pen.DashPattern = new float[2] { 2f, 1f };
							Graphics graphics;
							System.Drawing.Image image;
							try
							{
								int count = A.SlideMaster.CustomLayouts.Count;
								int num2 = 1;
								int num3 = clsPowerPoint.ScaleHeight(200, A);
								PageSetup pageSetup = A.PageSetup;
								float num4 = (float)num3 / pageSetup.SlideHeight;
								float num5 = 200f / pageSetup.SlideWidth;
								pageSetup = null;
								try
								{
									enumerator = A.Designs[1].SlideMaster.CustomLayouts.GetEnumerator();
									while (true)
									{
										if (!enumerator.MoveNext())
										{
											while (true)
											{
												switch (7)
												{
												case 0:
													break;
												default:
													goto end_IL_0467;
												}
												continue;
												end_IL_0467:
												break;
											}
											break;
										}
										pF = new PF(pF);
										pF.A = this;
										pF.A = (CustomLayout)enumerator.Current;
										if (this.m_A.CancellationPending)
										{
											e.Cancel = true;
											break;
										}
										this.m_A.ReportProgress((int)Math.Round((double)num2 / (double)count * 100.0));
										if (!PowerPointAddIn1.Slides.Helpers.IsSpecialLayout(pF.A))
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
											if (!pF.A.Name.EndsWith(AH.A(134194)) && !pF.A.Name.EndsWith(AH.A(134205)))
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
													Slide slide = A.Slides.AddSlide(1, pF.A);
													string text = modFunctionsIO.PathGetTempFileName();
													slide.Export(text, AH.A(63328), 200, num3);
													FileStream fileStream = new FileStream(text, FileMode.Open, FileAccess.Read);
													try
													{
														qF = new QF(qF);
														qF.A = pF;
														image = System.Drawing.Image.FromStream(fileStream);
														int width = image.Width;
														int height = image.Height;
														qF.A = new Bitmap(width, height, PixelFormat.Format32bppArgb);
														graphics = Graphics.FromImage(qF.A);
														graphics.DrawImage(image, 0, 0, width, height);
														try
														{
															try
															{
																enumerator2 = qF.A.A.Shapes.Placeholders.GetEnumerator();
																while (enumerator2.MoveNext())
																{
																	Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
																	if (shape.Fill.Visible == MsoTriState.msoFalse && shape.Line.Visible == MsoTriState.msoFalse)
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
																		graphics.DrawRectangle(pen, shape.Left * num5, shape.Top * num4, shape.Width * num5, shape.Height * num4);
																	}
																	shape = null;
																}
																while (true)
																{
																	switch (7)
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
																if (enumerator2 is IDisposable)
																{
																	while (true)
																	{
																		switch (3)
																		{
																		case 0:
																			break;
																		default:
																			(enumerator2 as IDisposable).Dispose();
																			goto end_IL_036d;
																		}
																		continue;
																		end_IL_036d:
																		break;
																	}
																}
															}
														}
														catch (Exception ex)
														{
															ProjectData.SetProjectError(ex);
															Exception ex2 = ex;
															clsReporting.LogException(ex2);
															ProjectData.ClearProjectError();
														}
														base.Dispatcher.Invoke(qF.A);
														image.Dispose();
														qF.A.Dispose();
														graphics.Dispose();
													}
													finally
													{
														if (fileStream != null)
														{
															while (true)
															{
																switch (4)
																{
																case 0:
																	break;
																default:
																	((IDisposable)fileStream).Dispose();
																	goto end_IL_03d6;
																}
																continue;
																end_IL_03d6:
																break;
															}
														}
													}
													File.Delete(text);
												}
												catch (Exception ex3)
												{
													ProjectData.SetProjectError(ex3);
													rF = new RF(rF);
													rF.A = pF;
													Exception a2 = ex3;
													rF.A = a2;
													if (!this.m_A.CancellationPending)
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
														clsReporting.LogException(rF.A);
														base.Dispatcher.Invoke(rF.A);
													}
													ProjectData.ClearProjectError();
												}
											}
										}
										num2++;
									}
								}
								finally
								{
									if (enumerator is IDisposable)
									{
										while (true)
										{
											switch (4)
											{
											case 0:
												break;
											default:
												(enumerator as IDisposable).Dispose();
												goto end_IL_047c;
											}
											continue;
											end_IL_047c:
											break;
										}
									}
								}
							}
							catch (Exception ex4)
							{
								ProjectData.SetProjectError(ex4);
								SF CS_0024_003C_003E8__locals4 = new SF(a3);
								CS_0024_003C_003E8__locals4.A = this;
								Exception a4 = ex4;
								CS_0024_003C_003E8__locals4.A = a4;
								if (this.m_A != null)
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
									base.Dispatcher.Invoke([SpecialName] () =>
									{
										CS_0024_003C_003E8__locals4.A.B(CS_0024_003C_003E8__locals4.A.Message);
									});
								}
								ProjectData.ClearProjectError();
							}
							pen.Dispose();
							pen = null;
							graphics = null;
							image = null;
						}
						if (B)
						{
							PowerPointAddIn1.Presentation.Helpers.CloseQuietly(A);
						}
						A = null;
						return;
					}
					}
				}
			}
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				C(AH.A(134726));
			});
		}
	}

	private void A(ref Microsoft.Office.Interop.PowerPoint.Presentation A, ref bool B)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		string templateId = Templates.GetTemplateId(application.ActivePresentation);
		if (templateId.Length > 0)
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
			Templates.B(ref A, ref B, templateId, application);
		}
		if (A == null)
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
			string tempFileName = Path.GetTempFileName();
			application.ActivePresentation.SaveCopyAs(tempFileName);
			A = application.Presentations.Open(tempFileName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
			B = true;
			base.Dispatcher.Invoke([SpecialName] () =>
			{
				dckError.Visibility = Visibility.Visible;
			});
		}
		application = null;
	}

	private void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbLoading.Value = e.ProgressPercentage;
	}

	private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		if (e.Cancelled)
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
			try
			{
				pbLoading.Visibility = Visibility.Hidden;
				if (LayoutItems.Count <= 0)
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
					if (this.m_A.Length > 0)
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
						InsertSlide.LayoutThumbnails.Add(this.m_A, LayoutItems);
					}
					ListBox listBox = lbxThumbs;
					listBox.Focus();
					try
					{
						ListBoxItem obj = (ListBoxItem)listBox.ItemContainerGenerator.ContainerFromIndex(0);
						obj.Focus();
						obj.IsSelected = true;
						_ = null;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					listBox = null;
					return;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	private void lbxThumbs_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		btnInsert.IsEnabled = lbxThumbs.SelectedItems.Count > 0;
	}

	private void btnInsert_Click(object sender, RoutedEventArgs e)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
		Slide slide = null;
		List<Slide> list = new List<Slide>();
		checked
		{
			try
			{
				DocumentWindow activeWindow = application.ActiveWindow;
				int num;
				if (activeWindow.Selection.Type == PpSelectionType.ppSelectionNone)
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
					try
					{
						if (activeWindow.ViewType != PpViewType.ppViewSlideSorter)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								if (activeWindow.Panes.Count <= 2)
								{
									break;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										continue;
									}
									if (activeWindow.Panes[3].ViewType == PpViewType.ppViewNotesPage)
									{
										activeWindow.Panes[3].Activate();
									}
									break;
								}
								break;
							}
						}
						else
						{
							((Slide)activeWindow.View.Slide).Select();
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					num = activeWindow.Selection.SlideRange[1].SlideIndex + 1;
				}
				else if (activeWindow.Selection.SlideRange.Count == 1)
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
					num = activeWindow.Selection.SlideRange[1].SlideIndex;
				}
				else
				{
					num = 500;
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = activeWindow.Selection.SlideRange.GetEnumerator();
						while (enumerator.MoveNext())
						{
							slide = (Slide)enumerator.Current;
							if (slide.SlideIndex >= num)
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
							num = slide.SlideIndex;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				activeWindow = null;
				application.StartNewUndoEntry();
				IEnumerator enumerator2 = default(IEnumerator);
				IEnumerator enumerator3 = default(IEnumerator);
				for (int i = lbxThumbs.SelectedItems.Count - 1; i >= 0; i += -1)
				{
					string layoutName = ((LayoutThumb)lbxThumbs.SelectedItems[i]).LayoutName;
					CustomLayout customLayout = null;
					try
					{
						enumerator2 = activePresentation.Designs.GetEnumerator();
						do
						{
							if (enumerator2.MoveNext())
							{
								Design design = (Design)enumerator2.Current;
								try
								{
									enumerator3 = design.SlideMaster.CustomLayouts.GetEnumerator();
									while (enumerator3.MoveNext())
									{
										CustomLayout customLayout2 = (CustomLayout)enumerator3.Current;
										if (Operators.CompareString(customLayout2.Name, layoutName, TextCompare: false) != 0)
										{
											continue;
										}
										while (true)
										{
											switch (2)
											{
											case 0:
												continue;
											}
											customLayout = customLayout2;
											break;
										}
										break;
									}
								}
								finally
								{
									if (enumerator3 is IDisposable)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											(enumerator3 as IDisposable).Dispose();
											break;
										}
									}
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
									goto end_IL_02cd;
								}
								continue;
								end_IL_02cd:
								break;
							}
							break;
						}
						while (customLayout == null);
					}
					finally
					{
						if (enumerator2 is IDisposable)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									continue;
								}
								(enumerator2 as IDisposable).Dispose();
								break;
							}
						}
					}
					if (customLayout != null)
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
						slide = activePresentation.Slides.AddSlide(num, customLayout);
						list.Add(slide);
						customLayout = null;
						PowerPointAddIn1.Slides.Helpers.DesignateSlideAsType(slide, SlideType.Content);
					}
					else
					{
						C(AH.A(127131) + layoutName + AH.A(134216));
					}
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					if (list.Count <= 0)
					{
						break;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						try
						{
							PowerPointAddIn1.Slides.Helpers.SelectMultipleSlides(activePresentation, list.Select([SpecialName] (Slide A) => A.SlideIndex));
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
						clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, AH.A(134440));
						break;
					}
					break;
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				if (A((Microsoft.Office.Interop.PowerPoint.Presentation)application))
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
					C(AH.A(134467));
				}
				else
				{
					B(ex6.Message);
					clsReporting.LogException(ex6);
				}
				ProjectData.ClearProjectError();
			}
			Focus();
			activePresentation = null;
			slide = null;
			application = null;
		}
	}

	private bool A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		bool result;
		try
		{
			if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(A, null, AH.A(134552), new object[0], null, null, null), null, AH.A(134577), new object[1] { 2 }, null, null, null), null, AH.A(134588), new object[0], null, null, null), PpViewType.ppViewSlideMaster, TextCompare: false))
			{
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
					result = true;
					break;
				}
			}
			else
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		if (this.m_A.IsBusy)
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
			try
			{
				this.m_A.CancelAsync();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		Close();
	}

	private void B(string A)
	{
		Forms.ErrorMessage(Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.WarningMessage(Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.InfoMessage(Window.GetWindow(this), A);
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
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(AH.A(134605), UriKind.Relative);
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
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					dckError = (DockPanel)target;
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
					tbNoTemplate = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					lbxThumbs = (ListBox)target;
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
					pbLoading = (ProgressBar)target;
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
					btnInsert = (Button)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnClose = (Button)target;
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

	[SpecialName]
	[CompilerGenerated]
	private void A()
	{
		pbLoading.IsIndeterminate = false;
		pbLoading.Value = 0.0;
	}

	[SpecialName]
	[CompilerGenerated]
	private void B()
	{
		C(AH.A(134726));
	}

	[SpecialName]
	[CompilerGenerated]
	private void C()
	{
		dckError.Visibility = Visibility.Visible;
	}
}
