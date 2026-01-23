using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.ImportExport;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Presentation;

namespace PowerPointAddIn1.Shapes.Templated;

[DesignerGenerated]
public sealed class wpfTemplatedShape : System.Windows.Controls.UserControl, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[CompilerGenerated]
	internal sealed class XD
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		[SpecialName]
		internal void A()
		{
			this.A.Copy();
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private readonly string m_A;

	private Microsoft.Office.Interop.PowerPoint.Presentation m_A;

	private bool m_A;

	[CompilerGenerated]
	private Microsoft.Office.Interop.PowerPoint.Shape m_A;

	private ObservableCollection<BaseInput> m_A;

	private bool m_B;

	private string m_B;

	private Visibility m_A;

	private ObservableCollection<string> m_A;

	private bool m_C;

	private int m_A;

	private string m_C;

	[CompilerGenerated]
	private float m_A;

	[CompilerGenerated]
	private float m_B;

	[CompilerGenerated]
	private int m_B;

	[CompilerGenerated]
	private string m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("scroller")]
	private ScrollViewer m_A;

	[AccessedThroughProperty("cbxTemplates")]
	[CompilerGenerated]
	private System.Windows.Controls.ComboBox m_A;

	[AccessedThroughProperty("radManual")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("radImport")]
	private System.Windows.Controls.RadioButton m_B;

	[AccessedThroughProperty("bdrImportError")]
	[CompilerGenerated]
	private Border m_A;

	[AccessedThroughProperty("tbImportError")]
	[CompilerGenerated]
	private TextBlock m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnTryAgain")]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("txtWorkbook")]
	[CompilerGenerated]
	private System.Windows.Controls.TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtImagesFolder")]
	private System.Windows.Controls.TextBox m_B;

	[AccessedThroughProperty("btnBrowse")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnGenerate")]
	private System.Windows.Controls.Button m_C;

	[AccessedThroughProperty("icInputs")]
	[CompilerGenerated]
	private ItemsControl m_A;

	[AccessedThroughProperty("btnSave")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private System.Windows.Controls.Button m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private System.Windows.Controls.Button m_F;

	[AccessedThroughProperty("btnAnother")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_G;

	[AccessedThroughProperty("btnDismiss")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_H;

	private bool m_D;

	private Microsoft.Office.Interop.PowerPoint.Shape TemplatedShape
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

	public ObservableCollection<BaseInput> Inputs
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(70559));
		}
	}

	public bool IsImageSelected
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(70572));
		}
	}

	public string ErrorText
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			A(AH.A(70603));
			int errorVisibility;
			if (value.Length <= 0)
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
				errorVisibility = 2;
			}
			else
			{
				errorVisibility = 0;
			}
			ErrorVisibility = (Visibility)errorVisibility;
		}
	}

	public Visibility ErrorVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(70622));
		}
	}

	public ObservableCollection<string> ImportValidationResults
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(70653));
		}
	}

	public bool ShowMissingColumns
	{
		get
		{
			return this.m_C;
		}
		set
		{
			this.m_C = value;
			A(AH.A(70700));
		}
	}

	public int ImportRowsCount
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(70737));
			btnGenerate.IsEnabled = value > 0;
		}
	}

	public string ImportErrorMessage
	{
		get
		{
			return this.m_C;
		}
		set
		{
			this.m_C = value;
			A(AH.A(70768));
			if (value.Length > 0)
			{
				bdrImportError.Visibility = Visibility.Visible;
				btnGenerate.Visibility = Visibility.Collapsed;
			}
			else
			{
				bdrImportError.Visibility = Visibility.Collapsed;
				btnGenerate.Visibility = Visibility.Visible;
			}
		}
	}

	private float InsertPointX
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

	private float InsertPointY
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private int LastTemplateIndex
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	private string LastBrowsedPath
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
		[CompilerGenerated]
		set
		{
			this.m_D = value;
		}
	}

	internal virtual ScrollViewer scroller
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

	internal virtual System.Windows.Controls.ComboBox cbxTemplates
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

	internal virtual System.Windows.Controls.RadioButton radManual
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

	internal virtual System.Windows.Controls.RadioButton radImport
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

	internal virtual Border bdrImportError
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

	internal virtual TextBlock tbImportError
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

	internal virtual System.Windows.Controls.Button btnTryAgain
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
			RoutedEventHandler value2 = TryAgainClicked;
			System.Windows.Controls.Button button = this.m_A;
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
			this.m_A = value;
			button = this.m_A;
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

	internal virtual System.Windows.Controls.TextBox txtWorkbook
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

	internal virtual System.Windows.Controls.TextBox txtImagesFolder
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

	internal virtual System.Windows.Controls.Button btnBrowse
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
			RoutedEventHandler value2 = BrowseImagesFolder;
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
				switch (3)
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

	internal virtual System.Windows.Controls.Button btnGenerate
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
			RoutedEventHandler value2 = GenerateShapes;
			System.Windows.Controls.Button button = this.m_C;
			if (button != null)
			{
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

	internal virtual ItemsControl icInputs
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

	internal virtual System.Windows.Controls.Button btnSave
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
			RoutedEventHandler value2 = SaveManualEntryShape;
			System.Windows.Controls.Button button = this.m_D;
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
			this.m_D = value;
			button = this.m_D;
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

	internal virtual System.Windows.Controls.Button btnCancel
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = CancelClicked;
			System.Windows.Controls.Button button = this.m_E;
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
			this.m_E = value;
			button = this.m_E;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnClose
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = CloseClicked;
			System.Windows.Controls.Button button = this.m_F;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_F = value;
			button = this.m_F;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnAnother
	{
		[CompilerGenerated]
		get
		{
			return this.m_G;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = AnotherClicked;
			System.Windows.Controls.Button button = this.m_G;
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
			this.m_G = value;
			button = this.m_G;
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

	internal virtual System.Windows.Controls.Button btnDismiss
	{
		[CompilerGenerated]
		get
		{
			return this.m_H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = DismissError;
			System.Windows.Controls.Button button = this.m_H;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_H = value;
			button = this.m_H;
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
				return;
			}
		}
	}

	public wpfTemplatedShape()
	{
		base.Unloaded += wpfTemplatedShape_Unloaded;
		base.KeyDown += HandleKeyPresses;
		this.m_A = AH.A(73259);
		this.m_A = null;
		this.m_B = false;
		this.m_B = "";
		this.m_A = Visibility.Collapsed;
		this.m_A = 0;
		InitializeComponent();
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
			switch (2)
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

	private void wpfTemplatedShape_Unloaded(object sender, RoutedEventArgs e)
	{
		B();
		C();
		TemplatedShape = null;
	}

	private void PaneSizeChanged(object sender, SizeChangedEventArgs e)
	{
		Panes.PaneSizeChanged(scroller, e);
	}

	public void ShowPane()
	{
		radImport.IsChecked = false;
		radManual.IsChecked = false;
		InsertPointX = 0f;
		InsertPointY = 0f;
		D();
		A();
		if (cbxTemplates.Items.Count > 0)
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
			cbxTemplates.SelectedIndex = 0;
		}
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
		this.m_A = false;
	}

	public void HidePane()
	{
		I();
		cbxTemplates.SelectedIndex = -1;
		B();
		C();
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			H();
		}
		TemplatedShape = null;
	}

	private void A()
	{
		base.SizeChanged += PaneSizeChanged;
		cbxTemplates.SelectionChanged += TemplateChanged;
		radImport.Checked += ImportModeChecked;
		radManual.Checked += ManualModeChecked;
	}

	private void B()
	{
		base.SizeChanged -= PaneSizeChanged;
		cbxTemplates.SelectionChanged -= TemplateChanged;
		radImport.Checked -= ImportModeChecked;
		radManual.Checked -= ManualModeChecked;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(NG.A.Application, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void C()
	{
		try
		{
			PowerPointAddIn1.Presentation.Helpers.CloseQuietly(this.m_A);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		this.m_A = null;
	}

	private void D()
	{
		string text;
		try
		{
			text = CloudStorage.GetShapeTemplatesFolder(KG.A.SettingsXml);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			text = "";
			ProjectData.ClearProjectError();
		}
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			IEnumerator enumerator2 = default(IEnumerator);
			if (text.Length > 0)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
					{
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						List<FileInfo> list = new DirectoryInfo(text + Conversions.ToString(Path.DirectorySeparatorChar)).EnumerateFiles(AH.A(70805), SearchOption.TopDirectoryOnly).ToList();
						if (list.Count == 1)
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
							this.m_A = PowerPointAddIn1.Presentation.Helpers.OpenQuietly(NG.A.Application, list.ElementAt(0).FullName);
							Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape> dictionary = new Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape>();
							try
							{
								enumerator = this.m_A.Slides.GetEnumerator();
								while (enumerator.MoveNext())
								{
									Slide slide = (Slide)enumerator.Current;
									try
									{
										enumerator2 = slide.Shapes.GetEnumerator();
										while (enumerator2.MoveNext())
										{
											Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
											if (shape.Type == MsoShapeType.msoGroup)
											{
												dictionary.Add(shape.Name, shape);
											}
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
													goto end_IL_0142;
												}
												continue;
												end_IL_0142:
												break;
											}
										}
									}
								}
							}
							finally
							{
								if (enumerator is IDisposable)
								{
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											(enumerator as IDisposable).Dispose();
											goto end_IL_0172;
										}
										continue;
										end_IL_0172:
										break;
									}
								}
							}
							if (dictionary.Count > 0)
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
								cbxTemplates.ItemsSource = dictionary;
							}
							else
							{
								B(AH.A(70818));
							}
							dictionary = null;
						}
						else
						{
							B(AH.A(70929));
						}
						list = null;
						return;
					}
					}
				}
			}
			B(AH.A(71115));
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			B(ex4.Message);
			ProjectData.ClearProjectError();
		}
	}

	private void TemplateChanged(object sender, SelectionChangedEventArgs e)
	{
		if (cbxTemplates.SelectedIndex > -1)
		{
			H();
			btnClose.Visibility = Visibility.Collapsed;
			btnAnother.Visibility = Visibility.Collapsed;
			btnCancel.Visibility = Visibility.Visible;
			try
			{
				TemplatedShape = A();
				TemplatedShape.Top = InsertPointY;
				TemplatedShape.Left = InsertPointX;
				A(TemplatedShape);
				Inputs = A(TemplatedShape);
				radManual.IsChecked = true;
				radManual.IsEnabled = true;
				radImport.IsEnabled = true;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(ex2.Message);
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			Focus();
		}
		else
		{
			radManual.IsChecked = false;
			radManual.IsEnabled = false;
			radImport.IsChecked = false;
			radImport.IsEnabled = false;
		}
	}

	private Microsoft.Office.Interop.PowerPoint.Shape A()
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.PowerPoint.Shape A = ((KeyValuePair<string, Microsoft.Office.Interop.PowerPoint.Shape>)cbxTemplates.SelectedItem).Value;
		try
		{
			clsClipboard.CopyWithWait((Action)([SpecialName] () =>
			{
				A.Copy();
			}), 4000);
			try
			{
				application.CommandBars.ExecuteMso(AH.A(58900));
				System.Windows.Forms.Application.DoEvents();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				application.ActiveWindow.Selection.SlideRange[1].Shapes.Paste().Select();
				ProjectData.ClearProjectError();
			}
			clsClipboard.ClearClipboard();
			return application.ActiveWindow.Selection.ShapeRange[1];
		}
		finally
		{
			application = null;
			A = null;
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		this.A(A, B: true);
	}

	private void B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		this.A(A, B: false);
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, bool B)
	{
		try
		{
			NewLateBinding.LateSetComplex(A, null, AH.A(69417), new object[1] { B }, null, null, OptimisticSet: false, RValueBase: true);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private ObservableCollection<BaseInput> A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		ObservableCollection<BaseInput> B = new ObservableCollection<BaseInput>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				this.A(a, ref B);
			}
			return B;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, ref ObservableCollection<BaseInput> B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		if (shape.Type != MsoShapeType.msoGroup)
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
			if (Regex.IsMatch(shape.Name, this.m_A, RegexOptions.IgnoreCase))
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
				string input;
				if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
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
					input = Regex.Replace(shape.TextFrame2.TextRange.Text, AH.A(71202), "");
					input = Regex.Replace(input, AH.A(71209), "");
				}
				else
				{
					input = Regex.Replace(shape.Name, this.m_A, "");
				}
				B.Add(new ImageInput(input, this.A(AH.A(71216)), A));
			}
			else if (shape.HasTextFrame == MsoTriState.msoTrue)
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
				if (shape.TextFrame2.HasText == MsoTriState.msoTrue)
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
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = shape.TextFrame2.TextRange.get_Paragraphs(-1, -1).GetEnumerator();
						IEnumerator enumerator2 = default(IEnumerator);
						while (enumerator.MoveNext())
						{
							TextRange2 textRange = (TextRange2)enumerator.Current;
							{
								enumerator2 = Regex.Matches(textRange.Text, AH.A(71237)).GetEnumerator();
								try
								{
									while (enumerator2.MoveNext())
									{
										Match match = (Match)enumerator2.Current;
										B.Add(new TextInput(match.Groups[1].Value, this.A(AH.A(71262)), textRange.get_Characters(checked(match.Index + 1), match.Length)));
									}
									while (true)
									{
										switch (3)
										{
										case 0:
											break;
										default:
											goto end_IL_01e5;
										}
										continue;
										end_IL_01e5:
										break;
									}
								}
								finally
								{
									IDisposable disposable = enumerator2 as IDisposable;
									if (disposable != null)
									{
										disposable.Dispose();
									}
								}
							}
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
			}
		}
		else
		{
			IEnumerator enumerator3 = default(IEnumerator);
			try
			{
				enumerator3 = shape.GroupItems.GetEnumerator();
				while (enumerator3.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current;
					this.A(a, ref B);
				}
			}
			finally
			{
				if (enumerator3 is IDisposable)
				{
					while (true)
					{
						switch (6)
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
		shape = null;
	}

	private DataTemplate A(string A)
	{
		return (DataTemplate)FindResource(A);
	}

	private void ChooseImage(object sender, RoutedEventArgs e)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Core.FileDialog fileDialog = ((Microsoft.Office.Interop.PowerPoint._Application)application).get_FileDialog(MsoFileDialogType.msoFileDialogFilePicker);
		fileDialog.Title = AH.A(71281);
		fileDialog.Filters.Clear();
		fileDialog.Filters.Add(AH.A(71306), AH.A(71317), RuntimeHelpers.GetObjectValue(Missing.Value));
		fileDialog.AllowMultiSelect = false;
		fileDialog.Show();
		FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
		_ = null;
		if (selectedItems.Count == 1)
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
			try
			{
				if (A(application))
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						D(A());
						break;
					}
				}
				else
				{
					ImageInput imageInput = (ImageInput)((System.Windows.Controls.Button)sender).DataContext;
					ImageInput imageInput2 = imageInput;
					Microsoft.Office.Interop.PowerPoint.Shape shape = application.ActiveWindow.Selection.SlideRange[1].Shapes.AddPicture2(Conversions.ToString(selectedItems.Cast<object>().ElementAtOrDefault(0)), MsoTriState.msoFalse, MsoTriState.msoTrue, imageInput2.Left, imageInput2.Top);
					imageInput2.ZOrder = imageInput2.Placeholder.ZOrderPosition;
					imageInput2.Placeholder.Delete();
					imageInput2.Placeholder = shape;
					imageInput2.IsPopulated = true;
					imageInput2 = null;
					TemplatedShape = A(TemplatedShape, shape, imageInput);
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(ex2.Message);
				ProjectData.ClearProjectError();
			}
			finally
			{
				ImageInput imageInput = null;
			}
		}
		application = null;
	}

	private bool A(Microsoft.Office.Interop.PowerPoint.Application A)
	{
		Microsoft.Office.Interop.PowerPoint.Shapes shapes = A.ActiveWindow.Selection.SlideRange[1].Shapes;
		for (int i = shapes.Count; i >= 1; i = checked(i + -1))
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = shapes[i];
			if (shape.Type == MsoShapeType.msoPlaceholder)
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
				if (ExcelToPowerPoint.IsPictureHolder(shape.PlaceholderFormat.Type))
				{
					return true;
				}
			}
			shape = null;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			shapes = null;
			return false;
		}
	}

	private string A()
	{
		return AH.A(71428);
	}

	private void SelectImage(object sender, RoutedEventArgs e)
	{
		ImageInput imageInput = (ImageInput)((System.Windows.Controls.Button)sender).DataContext;
		ImageInput imageInput2 = imageInput;
		try
		{
			imageInput2.ZOrder = imageInput2.Placeholder.ZOrderPosition;
			imageInput2.Placeholder.Delete();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		imageInput2.Placeholder = NG.A.Application.ActiveWindow.Selection.ShapeRange[1].Duplicate()[1];
		imageInput2.IsPopulated = true;
		TemplatedShape = A(TemplatedShape, imageInput2.Placeholder, imageInput);
		imageInput2 = null;
		imageInput = null;
	}

	private Microsoft.Office.Interop.PowerPoint.Shape A(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B, ImageInput C)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = B;
		shape.LockAspectRatio = MsoTriState.msoTrue;
		if (shape.Width > C.MaxWidth)
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
			shape.Width = C.MaxWidth;
		}
		if (shape.Height > C.MaxHeight)
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
			shape.Height = C.MaxHeight;
		}
		if (C.VerticalAlign == XlVAlign.xlVAlignCenter)
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
			if (shape.Height < C.MaxHeight)
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
				shape.Top = C.Top + (C.MaxHeight - shape.Height) / 2f;
			}
			else
			{
				shape.Top = C.Top;
			}
		}
		else if (C.VerticalAlign == XlVAlign.xlVAlignTop)
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
			shape.Top = C.Top;
		}
		else
		{
			shape.Top = C.Bottom - shape.Height;
		}
		if (C.HorizontalAlign == XlHAlign.xlHAlignCenter)
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
			if (shape.Width < C.MaxWidth)
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
				shape.Left = C.Left + (C.MaxWidth - shape.Width) / 2f;
			}
			else
			{
				shape.Left = C.Left;
			}
		}
		else if (C.HorizontalAlign == XlHAlign.xlHAlignLeft)
		{
			shape.Left = C.Left;
		}
		else
		{
			shape.Left = C.Right - shape.Width;
		}
		shape = null;
		if (radManual.IsChecked == true)
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
			this.B(A);
		}
		A.Ungroup().Select();
		B.Select(MsoTriState.msoFalse);
		A = NG.A.Application.ActiveWindow.Selection.ShapeRange.Group();
		A.Select();
		if (radManual.IsChecked == true)
		{
			this.A(A);
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape2 = B;
		while (shape2.ZOrderPosition < C.ZOrder)
		{
			shape2.ZOrder(MsoZOrderCmd.msoBringForward);
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
		while (shape2.ZOrderPosition > C.ZOrder)
		{
			shape2.ZOrder(MsoZOrderCmd.msoSendBackward);
		}
		shape2 = null;
		return A;
	}

	private void A(Selection A)
	{
		try
		{
			if (A.Type == PpSelectionType.ppSelectionShapes && A.ShapeRange.Count == 1)
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = A.ShapeRange[1];
				IsImageSelected = shape.Type == MsoShapeType.msoPicture || Images.IsGraphic(shape);
			}
			else
			{
				IsImageSelected = false;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			IsImageSelected = false;
			ProjectData.ClearProjectError();
		}
		finally
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		}
	}

	private void BrowseImagesFolder(object sender, RoutedEventArgs e)
	{
		string text = B();
		if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
		{
			txtImagesFolder.Text = text;
		}
	}

	private string B()
	{
		Microsoft.Office.Core.FileDialog fileDialog = ((Microsoft.Office.Interop.PowerPoint._Application)NG.A.Application).get_FileDialog(MsoFileDialogType.msoFileDialogFolderPicker);
		string result = string.Empty;
		fileDialog.Title = AH.A(71584);
		fileDialog.Filters.Clear();
		fileDialog.Show();
		fileDialog.AllowMultiSelect = false;
		fileDialog.InitialFileName = LastBrowsedPath;
		FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
		if (selectedItems.Count > 0)
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
			result = (LastBrowsedPath = selectedItems.Cast<object>().ElementAtOrDefault(0).ToString());
		}
		_ = null;
		selectedItems = null;
		return result;
	}

	private void ImportModeChecked(object sender, RoutedEventArgs e)
	{
		Microsoft.Office.Interop.Excel.Application application = null;
		string importErrorMessage = "";
		int importRowsCount = 0;
		bool showMissingColumns = false;
		ImportValidationResults = new ObservableCollection<string>();
		H();
		if (!A())
		{
			importErrorMessage = AH.A(71615);
		}
		else if (A(NG.A.Application))
		{
			importErrorMessage = A();
		}
		else
		{
			Range range = default(Range);
			try
			{
				application = InstanceManagement.GetExcelInstance(false);
				if (application == null)
				{
					throw new Exception();
				}
				if (application.Selection is Range)
				{
					range = (Range)application.Selection;
					if (Operators.ConditionalCompareObjectLess(range.Rows.CountLarge, 2, TextCompare: false))
					{
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
							importErrorMessage = AH.A(71734);
							break;
						}
					}
					else if (A(range) != 0)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							importErrorMessage = AH.A(71924);
							break;
						}
					}
					else
					{
						IEnumerator<BaseInput> enumerator = default(IEnumerator<BaseInput>);
						try
						{
							enumerator = Inputs.GetEnumerator();
							IEnumerator enumerator2 = default(IEnumerator);
							while (enumerator.MoveNext())
							{
								BaseInput current = enumerator.Current;
								bool flag = false;
								try
								{
									enumerator2 = ((IEnumerable)NewLateBinding.LateGet(range.Rows[1, RuntimeHelpers.GetObjectValue(Missing.Value)], null, AH.A(72007), new object[0], null, null, null)).GetEnumerator();
									while (true)
									{
										if (enumerator2.MoveNext())
										{
											if (!Operators.ConditionalCompareObjectEqual(((Range)enumerator2.Current).Text, current.Label, TextCompare: false))
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
												flag = true;
												break;
											}
											break;
										}
										while (true)
										{
											switch (1)
											{
											case 0:
												break;
											default:
												goto end_IL_01a8;
											}
											continue;
											end_IL_01a8:
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
											switch (4)
											{
											case 0:
												continue;
											}
											(enumerator2 as IDisposable).Dispose();
											break;
										}
									}
								}
								if (flag)
								{
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
								ImportValidationResults.Add(current.Label);
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_0204;
								}
								continue;
								end_IL_0204:
								break;
							}
						}
						finally
						{
							if (enumerator != null)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										continue;
									}
									enumerator.Dispose();
									break;
								}
							}
						}
						if (ImportValidationResults.Count > 0)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								ImportRowsCount = 0;
								importErrorMessage = AH.A(72018);
								showMissingColumns = true;
								break;
							}
						}
						else
						{
							importRowsCount = Conversions.ToInteger(Operators.SubtractObject(range.Rows.CountLarge, 1));
							txtWorkbook.Text = application.ActiveWorkbook.Name;
						}
					}
				}
				else
				{
					importErrorMessage = AH.A(72266);
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				importErrorMessage = AH.A(72327);
				ProjectData.ClearProjectError();
			}
			if (application != null)
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
				JG.A(application);
				JG.A(range);
				application = null;
				range = null;
			}
		}
		ImportRowsCount = importRowsCount;
		ImportErrorMessage = importErrorMessage;
		ShowMissingColumns = showMissingColumns;
	}

	private void ManualModeChecked(object sender, RoutedEventArgs e)
	{
		System.Windows.Controls.ComboBox comboBox = cbxTemplates;
		int selectedIndex = comboBox.SelectedIndex;
		comboBox.SelectionChanged -= TemplateChanged;
		comboBox.SelectedIndex = -1;
		comboBox.SelectionChanged += TemplateChanged;
		comboBox.SelectedIndex = selectedIndex;
		_ = null;
	}

	private int A(Range A)
	{
		int result;
		Range range = default(Range);
		try
		{
			range = A.SpecialCells(XlCellType.xlCellTypeBlanks, RuntimeHelpers.GetObjectValue(Missing.Value));
			result = 0 - ((range != null) ? 1 : 0);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = 0;
			ProjectData.ClearProjectError();
		}
		finally
		{
			JG.A(range);
			range = null;
		}
		return result;
	}

	private void TryAgainClicked(object sender, RoutedEventArgs e)
	{
		E();
	}

	private void E()
	{
		radManual.IsChecked = true;
		radImport.IsChecked = true;
	}

	private bool A()
	{
		return Inputs.Select([SpecialName] (BaseInput A) => A.Label).Distinct().Count() == Inputs.Count;
	}

	private void CopyColumnName(object sender, RoutedEventArgs e)
	{
		string text = Conversions.ToString(((System.Windows.Controls.Button)sender).DataContext);
		try
		{
			System.Windows.Clipboard.SetText(text);
			E(AH.A(72446));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			C(AH.A(72461));
			ProjectData.ClearProjectError();
		}
	}

	private void GenerateShapes(object sender, RoutedEventArgs e)
	{
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		Microsoft.Office.Interop.Excel.Application application2 = null;
		float num = 0f;
		float num2 = 0f;
		int num3 = 0;
		bool flag = false;
		checked
		{
			Range range = default(Range);
			Dictionary<string, int> dictionary;
			ObservableCollection<BaseInput> observableCollection;
			try
			{
				application2 = InstanceManagement.GetExcelInstance(false);
				range = (Range)application2.Selection;
				if (Operators.ConditionalCompareObjectLess(range.Rows.CountLarge, 2, TextCompare: false))
				{
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
						throw new Exception();
					}
				}
				if (A(range) != 0)
				{
					throw new Exception();
				}
				Microsoft.Office.Interop.PowerPoint.Shape shape = A();
				Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
				shape2.Top = num;
				shape2.Left = num2;
				num2 += shape2.Width;
				float width = shape2.Width;
				float height = shape2.Height;
				shape2 = null;
				observableCollection = A(shape);
				dictionary = new Dictionary<string, int>();
				IEnumerator<BaseInput> enumerator = default(IEnumerator<BaseInput>);
				try
				{
					enumerator = observableCollection.GetEnumerator();
					while (enumerator.MoveNext())
					{
						BaseInput current = enumerator.Current;
						bool flag2 = false;
						int count = range.Columns.Count;
						int num4 = 1;
						while (true)
						{
							if (num4 <= count)
							{
								string left = Conversions.ToString(NewLateBinding.LateGet(range.Cells[1, num4], null, AH.A(70464), new object[0], null, null, null));
								if (Operators.CompareString(left, current.Label, TextCompare: false) == 0)
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
									flag2 = true;
									dictionary.Add(current.Label, num4);
									break;
								}
								num4++;
								continue;
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
						if (flag2)
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
							shape.Delete();
							throw new Exception();
						}
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_01cb;
						}
						continue;
						end_IL_01cb:
						break;
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							enumerator.Dispose();
							break;
						}
					}
				}
				try
				{
					int count2 = range.Columns.Count;
					int num4 = 1;
					while (true)
					{
						if (num4 <= count2)
						{
							if (Operators.CompareString(NewLateBinding.LateGet(range.Cells[1, num4], null, AH.A(70464), new object[0], null, null, null).ToString().ToUpper(), AH.A(72486), TextCompare: false) == 0)
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
								dictionary.Add(AH.A(72486), num4);
								break;
							}
							num4++;
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
					PageSetup pageSetup = application.ActivePresentation.PageSetup;
					float slideHeight = pageSetup.SlideHeight;
					float slideWidth = pageSetup.SlideWidth;
					_ = null;
					Slide slide = application.ActiveWindow.Selection.SlideRange[1];
					string text = txtImagesFolder.Text.Trim();
					if (text.Length > 0)
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
						if (!text.EndsWith(Conversions.ToString(Path.DirectorySeparatorChar)))
						{
							text += Conversions.ToString(Path.DirectorySeparatorChar);
						}
					}
					application.StartNewUndoEntry();
					int count3 = range.Rows.Count;
					IEnumerator<BaseInput> enumerator2 = default(IEnumerator<BaseInput>);
					for (int i = 2; i <= count3; i++)
					{
						if (i > 2)
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
							if (num2 + width > slideWidth)
							{
								num2 = 0f;
								num += height;
							}
							if (num + height > slideHeight)
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
								slide = slide.Duplicate()[1];
								slide.Select();
								slide.Shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
								num2 = 0f;
								num = 0f;
							}
							shape = A();
							shape.Left = num2;
							shape.Top = num;
							num2 += width;
							observableCollection = A(shape);
						}
						try
						{
							enumerator2 = observableCollection.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								BaseInput current2 = enumerator2.Current;
								num4 = dictionary[current2.Label];
								string left = Conversions.ToString(NewLateBinding.LateGet(range.Cells[i, num4], null, AH.A(70464), new object[0], null, null, null));
								if (current2 is TextInput)
								{
									((TextInput)current2).Text = left;
								}
								else
								{
									if (!(current2 is ImageInput))
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
									string text2 = text + left.Trim();
									ImageInput imageInput = (ImageInput)current2;
									if (text2.Length <= 0)
									{
										throw new Exception(AH.A(72532) + current2.Label + AH.A(72591));
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
									Microsoft.Office.Interop.PowerPoint.Shape shape3;
									try
									{
										shape3 = slide.Shapes.AddPicture2(text2, MsoTriState.msoFalse, MsoTriState.msoTrue, imageInput.Left, imageInput.Top);
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										if (!File.Exists(text2))
										{
											while (true)
											{
												switch (7)
												{
												case 0:
													break;
												default:
													throw new Exception(AH.A(72493) + text2);
												}
											}
										}
										throw ex2;
									}
									imageInput.ZOrder = imageInput.Placeholder.ZOrderPosition;
									imageInput.Placeholder.Delete();
									imageInput.Placeholder = shape3;
									imageInput = null;
									shape = A(shape, shape3, (ImageInput)current2);
								}
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_05b2;
								}
								continue;
								end_IL_05b2:
								break;
							}
						}
						finally
						{
							if (enumerator2 != null)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										continue;
									}
									enumerator2.Dispose();
									break;
								}
							}
						}
						try
						{
							num4 = dictionary[AH.A(72486)];
							Microsoft.Office.Interop.PowerPoint.Shape shape4;
							(shape4 = shape).Name = Conversions.ToString(Operators.ConcatenateObject(shape4.Name, Operators.ConcatenateObject(AH.A(72596), NewLateBinding.LateGet(range.Cells[i, num4], null, AH.A(70464), new object[0], null, null, null))));
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
						num3++;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						flag = true;
						cbxTemplates.SelectedIndex = -1;
						btnCancel.Visibility = Visibility.Collapsed;
						btnClose.Visibility = Visibility.Visible;
						btnAnother.Visibility = Visibility.Collapsed;
						break;
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					try
					{
						shape.Delete();
					}
					catch (Exception ex7)
					{
						ProjectData.SetProjectError(ex7);
						Exception ex8 = ex7;
						ProjectData.ClearProjectError();
					}
					clsReporting.LogException(ex6);
					C(AH.A(72607) + num3 + AH.A(72674) + ex6.Message);
					ProjectData.ClearProjectError();
				}
				finally
				{
					Slide slide = null;
				}
			}
			catch (Exception ex9)
			{
				ProjectData.SetProjectError(ex9);
				Exception ex10 = ex9;
				E();
				ProjectData.ClearProjectError();
			}
			if (application2 != null)
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
				JG.A(application2);
				JG.A(range);
				application2 = null;
				range = null;
			}
			if (flag)
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
				F(AH.A(72753) + num3 + AH.A(72774));
				Base.LogActivity(AH.A(72791));
			}
			application = null;
			dictionary = null;
			observableCollection = null;
		}
	}

	private void SaveManualEntryShape(object sender, RoutedEventArgs e)
	{
		IEnumerator<BaseInput> enumerator = default(IEnumerator<BaseInput>);
		try
		{
			enumerator = Inputs.GetEnumerator();
			while (enumerator.MoveNext())
			{
				if (enumerator.Current.IsPopulated)
				{
					continue;
				}
				if (Forms.OkCancelMessage2(AH.A(72864)) != DialogResult.Cancel)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
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
					switch (7)
					{
					case 0:
						continue;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
		Base.LogActivity(AH.A(73042));
		B(TemplatedShape);
		TemplatedShape = null;
		LastTemplateIndex = cbxTemplates.SelectedIndex;
		cbxTemplates.SelectedIndex = -1;
		btnCancel.Visibility = Visibility.Collapsed;
		btnClose.Visibility = Visibility.Visible;
		btnAnother.Visibility = Visibility.Visible;
		InsertPointX += 20f;
		InsertPointY += 20f;
	}

	private void AnotherClicked(object sender, RoutedEventArgs e)
	{
		cbxTemplates.SelectedIndex = LastTemplateIndex;
	}

	private void CloseClicked(object sender, RoutedEventArgs e)
	{
		G();
	}

	private void CancelClicked(object sender, RoutedEventArgs e)
	{
		F();
	}

	private void DismissError(object sender, RoutedEventArgs e)
	{
		G();
	}

	private void HandleKeyPresses(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Escape)
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
			if (btnClose.Visibility == Visibility.Visible)
			{
				G();
				return;
			}
			if (btnCancel.Visibility != Visibility.Visible)
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
				F();
				return;
			}
		}
	}

	private void F()
	{
		H();
		G();
	}

	private void G()
	{
		this.m_A = true;
		Pane.A();
		Pane.B();
	}

	private void H()
	{
		if (TemplatedShape != null)
		{
			B(TemplatedShape);
			try
			{
				TemplatedShape.Delete();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			TemplatedShape = null;
		}
	}

	private void B(string A)
	{
		ErrorText = A;
	}

	private void I()
	{
		ErrorText = "";
	}

	private void C(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void E(string A)
	{
		Forms.InfoMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void F(string A)
	{
		Forms.SuccessMessage(System.Windows.Window.GetWindow(this), A);
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
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_D = true;
			Uri resourceLocator = new Uri(AH.A(73115), UriKind.Relative);
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
		if (connectionId == 3)
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
					scroller = (ScrollViewer)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					cbxTemplates = (System.Windows.Controls.ComboBox)target;
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
					radManual = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					radImport = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					bdrImportError = (Border)target;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					tbImportError = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnTryAgain = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtWorkbook = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtImagesFolder = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			btnBrowse = (System.Windows.Controls.Button)target;
			return;
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
					btnGenerate = (System.Windows.Controls.Button)target;
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
					icInputs = (ItemsControl)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnSave = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnCancel = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 18:
			btnClose = (System.Windows.Controls.Button)target;
			break;
		case 19:
			btnAnother = (System.Windows.Controls.Button)target;
			break;
		case 20:
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				btnDismiss = (System.Windows.Controls.Button)target;
				return;
			}
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

	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
			((System.Windows.Controls.Button)target).Click += ChooseImage;
		}
		if (connectionId == 2)
		{
			((System.Windows.Controls.Button)target).Click += SelectImage;
		}
		if (connectionId != 9)
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
			((System.Windows.Controls.Button)target).Click += CopyColumnName;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
