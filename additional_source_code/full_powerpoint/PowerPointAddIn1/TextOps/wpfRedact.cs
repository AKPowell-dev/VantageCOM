using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.TextOps;

[DesignerGenerated]
public sealed class wpfRedact : Window, IComponentConnector
{
	[CompilerGenerated]
	internal sealed class CG
	{
		public string A;

		public wpfRedact A;

		public CG(CG A)
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
			this.A.lblStatus.Text = AH.A(157757) + this.A + AH.A(170951);
		}
	}

	private BackgroundWorker m_A;

	private string m_A;

	private int m_A;

	private bool m_A;

	[CompilerGenerated]
	private SlideRange m_A;

	[AccessedThroughProperty("txtFind")]
	[CompilerGenerated]
	private System.Windows.Controls.TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lblStatus")]
	private TextBlock m_A;

	[AccessedThroughProperty("btnRedact")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnCancel")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	private bool m_B;

	private SlideRange SelectedSlideRange
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal virtual System.Windows.Controls.TextBox txtFind
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

	internal virtual TextBlock lblStatus
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

	internal virtual System.Windows.Controls.Button btnRedact
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
			RoutedEventHandler value2 = btnRedact_Click;
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
				switch (2)
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
			this.m_B = value;
			button = this.m_B;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	public wpfRedact()
	{
		base.Closing += wpfRedact_Closing;
		base.Deactivated += wpfRedact_Deactivated;
		this.m_A = null;
		this.m_A = "";
		this.m_A = 0;
		this.m_A = false;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		lblStatus.Text = "";
		try
		{
			SelectedSlideRange = NG.A.Application.ActiveWindow.Selection.SlideRange;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void wpfRedact_Closing(object sender, CancelEventArgs e)
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
	}

	private void wpfRedact_Deactivated(object sender, EventArgs e)
	{
		A();
	}

	private void btnRedact_Click(object sender, RoutedEventArgs e)
	{
		bool flag = false;
		if (txtFind.Text.Length == 0)
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
					Forms.WarningMessage(AH.A(157455));
					return;
				}
			}
		}
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = NG.A.Application.ActivePresentation;
		if (activePresentation.Saved == MsoTriState.msoFalse)
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
			if (activePresentation.Path.Length > 0)
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
				DialogResult dialogResult = Forms.YesNoCancelMessage(AH.A(157488));
				if (dialogResult != System.Windows.Forms.DialogResult.Cancel)
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
					if (dialogResult == System.Windows.Forms.DialogResult.Yes)
					{
						activePresentation.Save();
					}
				}
				else
				{
					flag = true;
				}
			}
		}
		activePresentation = null;
		if (!flag)
		{
			btnRedact.IsEnabled = false;
			btnCancel.Content = AH.A(157646);
			this.m_A = new BackgroundWorker();
			BackgroundWorker a = this.m_A;
			a.WorkerSupportsCancellation = true;
			a.WorkerReportsProgress = false;
			a.DoWork += bgw_DoWork;
			a.RunWorkerCompleted += bgw_RunWorkerCompleted;
			a.RunWorkerAsync();
			_ = null;
		}
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			this.m_A = txtFind.Text.Trim();
		});
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = application.ActivePresentation;
		application.StartNewUndoEntry();
		checked
		{
			try
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = activePresentation.Slides.GetEnumerator();
					IEnumerator enumerator2 = default(IEnumerator);
					IEnumerator enumerator3 = default(IEnumerator);
					IEnumerator enumerator4 = default(IEnumerator);
					while (enumerator.MoveNext())
					{
						Slide slide = (Slide)enumerator.Current;
						if (this.m_A.CancellationPending)
						{
							e.Cancel = true;
							break;
						}
						application.ActiveWindow.View.GotoSlide(slide.SlideIndex);
						try
						{
							enumerator2 = slide.Shapes.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
								if (this.m_A.CancellationPending)
								{
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
										e.Cancel = true;
										break;
									}
									break;
								}
								if (shape.Type == MsoShapeType.msoGroup)
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
										enumerator3 = shape.GroupItems.GetEnumerator();
										while (enumerator3.MoveNext())
										{
											Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current;
											if (shape2.HasTextFrame != MsoTriState.msoTrue)
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
											A(shape2, this.m_A, ref e);
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
								}
								else if (shape.HasTextFrame == MsoTriState.msoTrue)
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
									A(shape, this.m_A, ref e);
								}
								else if (shape.HasTable == MsoTriState.msoTrue)
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
									Table table = shape.Table;
									int count = table.Rows.Count;
									int count2 = table.Columns.Count;
									int num = count;
									for (int num2 = 1; num2 <= num; num2++)
									{
										int num3 = count2;
										int num4 = 1;
										while (true)
										{
											if (num4 <= num3)
											{
												if (this.m_A.CancellationPending)
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
													e.Cancel = true;
													break;
												}
												if (table.Cell(num2, num4).Shape.HasTextFrame == MsoTriState.msoTrue)
												{
													A(table.Cell(num2, num4).Shape, this.m_A, ref e);
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
									table = null;
								}
								else if (shape.HasSmartArt == MsoTriState.msoTrue)
								{
									try
									{
										enumerator4 = shape.SmartArt.AllNodes.GetEnumerator();
										while (true)
										{
											if (enumerator4.MoveNext())
											{
												SmartArtNode smartArtNode = (SmartArtNode)enumerator4.Current;
												if (this.m_A.CancellationPending)
												{
													while (true)
													{
														switch (5)
														{
														case 0:
															continue;
														}
														e.Cancel = true;
														break;
													}
													break;
												}
												if (smartArtNode.TextFrame2.HasText != MsoTriState.msoTrue)
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
													break;
												}
												A(smartArtNode.TextFrame2.TextRange, this.m_A, ref e);
												continue;
											}
											while (true)
											{
												switch (1)
												{
												case 0:
													break;
												default:
													goto end_IL_0320;
												}
												continue;
												end_IL_0320:
												break;
											}
											break;
										}
									}
									finally
									{
										if (enumerator4 is IDisposable)
										{
											while (true)
											{
												switch (1)
												{
												case 0:
													continue;
												}
												(enumerator4 as IDisposable).Dispose();
												break;
											}
										}
									}
								}
								else
								{
									_ = shape.HasChart;
									_ = -1;
								}
							}
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
						foreach (Microsoft.Office.Interop.PowerPoint.Shape shape3 in slide.NotesPage.Shapes)
						{
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
							if (shape3.HasTextFrame != MsoTriState.msoTrue)
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
								break;
							}
							A(shape3, this.m_A, ref e);
						}
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
								continue;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(ex2.Message);
				clsReporting.LogException(ex2);
				this.m_A = true;
				ProjectData.ClearProjectError();
			}
			application = null;
			activePresentation = null;
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, string B, ref DoWorkEventArgs C)
	{
		if (A.TextFrame2.HasText != MsoTriState.msoTrue)
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
			this.A(A.TextFrame2.TextRange, B, ref C);
			return;
		}
	}

	private void A(TextRange2 A, string B, ref DoWorkEventArgs C)
	{
		CG a = default(CG);
		CG CS_0024_003C_003E8__locals4 = new CG(a);
		CS_0024_003C_003E8__locals4.A = this;
		TextRange2 textRange = null;
		int num = 0;
		int length = B.Length;
		textRange = A.Find(B, 0, MsoTriState.msoFalse, MsoTriState.msoTrue);
		checked
		{
			while (textRange != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					if (textRange.Length == length)
					{
						if (this.m_A.CancellationPending)
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
									C.Cancel = true;
									return;
								}
							}
						}
						num = textRange.Start + textRange.Length - 1;
						Redact.RedactWord(textRange);
						this.m_A++;
						CS_0024_003C_003E8__locals4.A = this.m_A.ToString();
						base.Dispatcher.Invoke([SpecialName] () =>
						{
							CS_0024_003C_003E8__locals4.A.lblStatus.Text = AH.A(157757) + CS_0024_003C_003E8__locals4.A + AH.A(170951);
						});
						textRange = A.Find(B, num, MsoTriState.msoFalse, MsoTriState.msoTrue);
						break;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
			}
		}
	}

	private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, AH.A(157655));
		if (!this.m_A)
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
			if (!e.Cancelled)
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
				Forms.InfoMessage(A());
			}
		}
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
			this.m_A.Dispose();
			this.m_A = null;
		}
		if (SelectedSlideRange != null)
		{
			SelectedSlideRange.Select();
			SelectedSlideRange = null;
		}
		Close();
	}

	private void A()
	{
		if (this.m_A != null && this.m_A.IsBusy)
		{
			this.m_A.CancelAsync();
		}
	}

	private void btnCancel_Click(object sender, RoutedEventArgs e)
	{
		if (this.m_A != null)
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
				this.m_A.CancelAsync();
			}
			if (this.m_A > 0)
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
				Forms.InfoMessage(A() + AH.A(157682));
			}
		}
		Close();
	}

	private string A()
	{
		if (this.m_A > 0)
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
					return AH.A(157757) + this.m_A + AH.A(157776) + this.m_A + AH.A(72591);
				}
			}
		}
		return AH.A(157811);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_B)
		{
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(157842), UriKind.Relative);
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
		if (connectionId == 1)
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
					txtFind = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 2:
			lblStatus = (TextBlock)target;
			break;
		case 3:
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				btnRedact = (System.Windows.Controls.Button)target;
				return;
			}
		case 4:
			btnCancel = (System.Windows.Controls.Button)target;
			break;
		default:
			this.m_B = true;
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
	private void B()
	{
		this.m_A = txtFind.Text.Trim();
	}
}
