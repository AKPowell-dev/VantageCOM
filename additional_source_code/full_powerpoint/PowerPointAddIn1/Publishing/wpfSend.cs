using System;
using System.CodeDom.Compiler;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Publishing;

[DesignerGenerated]
public sealed class wpfSend : Window, IComponentConnector
{
	private Microsoft.Office.Interop.PowerPoint.Presentation m_A;

	private XmlDocument m_A;

	private bool m_A;

	[AccessedThroughProperty("chkSendPdf")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSendFile")]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("chkSendLink")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox C;

	[CompilerGenerated]
	[AccessedThroughProperty("radScopeEntire")]
	private System.Windows.Controls.RadioButton m_A;

	[AccessedThroughProperty("radScopeSelected")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_B;

	[AccessedThroughProperty("grpAttachment")]
	[CompilerGenerated]
	private System.Windows.Controls.GroupBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtName")]
	private System.Windows.Controls.TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkCompress")]
	private System.Windows.Controls.CheckBox D;

	[CompilerGenerated]
	[AccessedThroughProperty("grpPdf")]
	private System.Windows.Controls.GroupBox m_B;

	[AccessedThroughProperty("chkOpen")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox E;

	[AccessedThroughProperty("chkSaveCopy")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox F;

	[CompilerGenerated]
	[AccessedThroughProperty("txtFolder")]
	private System.Windows.Controls.TextBox m_B;

	[AccessedThroughProperty("chkDuplex")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox G;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private System.Windows.Controls.Button m_B;

	private bool m_B;

	internal virtual System.Windows.Controls.CheckBox chkSendPdf
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

	internal virtual System.Windows.Controls.CheckBox chkSendFile
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

	internal virtual System.Windows.Controls.CheckBox chkSendLink
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
			C = value;
		}
	}

	internal virtual System.Windows.Controls.RadioButton radScopeEntire
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

	internal virtual System.Windows.Controls.RadioButton radScopeSelected
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

	internal virtual System.Windows.Controls.GroupBox grpAttachment
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

	internal virtual System.Windows.Controls.TextBox txtName
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

	internal virtual System.Windows.Controls.CheckBox chkCompress
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

	internal virtual System.Windows.Controls.GroupBox grpPdf
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

	internal virtual System.Windows.Controls.CheckBox chkOpen
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
			E = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSaveCopy
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

	internal virtual System.Windows.Controls.TextBox txtFolder
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
			MouseButtonEventHandler value2 = txtFolder_PreviewMouseDown;
			System.Windows.Controls.TextBox textBox = this.m_B;
			if (textBox != null)
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
				textBox.PreviewMouseDown -= value2;
			}
			this.m_B = value;
			textBox = this.m_B;
			if (textBox != null)
			{
				textBox.PreviewMouseDown += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkDuplex
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

	internal virtual System.Windows.Controls.Button btnOk
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
			RoutedEventHandler value2 = btnOk_Click;
			System.Windows.Controls.Button button = this.m_A;
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
			return this.m_B;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_B = value;
		}
	}

	public wpfSend(Microsoft.Office.Interop.PowerPoint.Presentation pres)
	{
		base.Loaded += wpfSend_Loaded;
		base.Closing += wpfSend_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = pres;
		this.m_A = KG.A.SettingsXml;
	}

	private void wpfSend_Loaded(object sender, RoutedEventArgs e)
	{
		string path = this.m_A.Path;
		string name = this.m_A.Name;
		radScopeEntire.Checked += ScopeChanged;
		radScopeSelected.Checked += ScopeChanged;
		chkSendPdf.Checked += PdfOrFileChecked;
		chkSendPdf.Unchecked += PdfUnchecked;
		chkSendFile.Checked += PdfOrFileChecked;
		chkSendFile.Unchecked += FileUnchecked;
		chkSendLink.Checked += LinkChecked;
		try
		{
			Selection selection = this.m_A.Application.ActiveWindow.Selection;
			if (selection.SlideRange == null)
			{
				throw new Exception();
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
				if (selection.Type == PpSelectionType.ppSelectionSlides)
				{
					radScopeSelected.IsChecked = true;
				}
				else
				{
					radScopeEntire.IsChecked = true;
				}
				this.m_A = true;
				selection = null;
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			radScopeSelected.IsEnabled = false;
			this.m_A = false;
			radScopeEntire.IsChecked = true;
			ProjectData.ClearProjectError();
		}
		clsPublish.PopulateForm(this.m_A, path, name, chkSendPdf, chkSendFile, chkSendLink, txtName, txtFolder);
		XmlNode xmlNode = this.m_A.DocumentElement.SelectSingleNode(AH.A(104910));
		chkCompress.IsChecked = Conversions.ToInteger(xmlNode.SelectSingleNode(AH.A(104927)).InnerText) != 0;
		chkOpen.IsChecked = Conversions.ToInteger(xmlNode.SelectSingleNode(AH.A(104964)).InnerText) != 0;
		chkSaveCopy.IsChecked = Conversions.ToInteger(xmlNode.SelectSingleNode(AH.A(104995)).InnerText) != 0;
		chkDuplex.IsChecked = Conversions.ToInteger(xmlNode.SelectSingleNode(AH.A(105026)).InnerText) != 0;
		xmlNode = null;
	}

	private void ScopeChanged(object sender, RoutedEventArgs e)
	{
		if (radScopeSelected.IsChecked != true)
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
			chkSendLink.IsChecked = false;
			DocumentWindow activeWindow = this.m_A.Application.ActiveWindow;
			if (activeWindow.Selection.Type != PpSelectionType.ppSelectionSlides)
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
				if (activeWindow.Selection.SlideRange != null)
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
					try
					{
						if (activeWindow.Panes[1].ViewType != PpViewType.ppViewThumbnails)
						{
							throw new Exception();
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								continue;
							}
							activeWindow.Panes[1].Activate();
							activeWindow.Selection.SlideRange.Select();
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						radScopeEntire.IsChecked = true;
						radScopeSelected.IsEnabled = false;
						ProjectData.ClearProjectError();
					}
				}
			}
			activeWindow = null;
			return;
		}
	}

	private void PdfOrFileChecked(object sender, RoutedEventArgs e)
	{
		chkSendLink.IsChecked = false;
		radScopeSelected.IsEnabled = this.m_A;
		grpAttachment.Visibility = Visibility.Visible;
		if (chkSendPdf.IsChecked == true)
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
					grpPdf.Visibility = Visibility.Visible;
					return;
				}
			}
		}
		grpPdf.Visibility = Visibility.Collapsed;
	}

	private void PdfUnchecked(object sender, RoutedEventArgs e)
	{
		grpPdf.Visibility = Visibility.Collapsed;
		A();
	}

	private void FileUnchecked(object sender, RoutedEventArgs e)
	{
		A();
	}

	private void A()
	{
		bool? isChecked = chkSendPdf.IsChecked;
		bool? flag;
		if (!isChecked.HasValue)
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
			flag = isChecked;
		}
		else
		{
			flag = isChecked != true;
		}
		bool? flag2 = flag;
		if (flag2.HasValue)
		{
			if (flag2 != true)
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
				break;
			}
		}
		isChecked = chkSendFile.IsChecked;
		if (((!isChecked) ?? isChecked) != true || !flag2.HasValue)
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
			grpAttachment.Visibility = Visibility.Collapsed;
			return;
		}
	}

	private void LinkChecked(object sender, RoutedEventArgs e)
	{
		chkSendPdf.IsChecked = false;
		chkSendFile.IsChecked = false;
		grpAttachment.Visibility = Visibility.Collapsed;
		grpPdf.Visibility = Visibility.Collapsed;
		radScopeEntire.IsChecked = true;
		radScopeSelected.IsEnabled = false;
	}

	private void txtFolder_PreviewMouseDown(object sender, MouseButtonEventArgs e)
	{
		clsPublish.PickSaveCopyFolder(((_Application)NG.A.Application).get_FileDialog(MsoFileDialogType.msoFileDialogFolderPicker), txtFolder, chkSaveCopy);
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void wpfSend_Closing(object sender, CancelEventArgs e)
	{
		bool? obj;
		bool? flag3;
		if (base.DialogResult == true)
		{
			if (clsPublish.HasIllegalCharacters(txtName, (Action<string>)A))
			{
				e.Cancel = true;
				goto IL_039c;
			}
			bool? isChecked = chkSendPdf.IsChecked;
			bool? flag;
			isChecked = (flag = (!isChecked) ?? isChecked);
			if (isChecked.HasValue)
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
				if (flag != true)
				{
					obj = false;
					goto IL_0122;
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
			isChecked = chkSendFile.IsChecked;
			bool? flag2;
			if (!isChecked.HasValue)
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
				flag2 = isChecked;
			}
			else
			{
				flag2 = isChecked != true;
			}
			flag3 = flag2;
			isChecked = flag2;
			if (!isChecked.HasValue)
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
			else if (flag3 != true)
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
				obj = false;
			}
			else
			{
				obj = flag;
			}
			goto IL_0122;
		}
		goto IL_0449;
		IL_0122:
		bool? flag4 = obj;
		if (flag4.HasValue)
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
			if (flag4 != true)
			{
				goto IL_01cc;
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
		flag3 = chkSendLink.IsChecked;
		bool? flag5;
		if (!flag3.HasValue)
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
			flag5 = flag3;
		}
		else
		{
			flag5 = flag3 != true;
		}
		flag3 = flag5;
		if (flag3 == true)
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
			if (flag4.HasValue)
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
				A(AH.A(105075));
				e.Cancel = true;
				goto IL_039c;
			}
		}
		goto IL_01cc;
		IL_039c:
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
			try
			{
				clsPublish.SaveSendSettings(this.m_A, chkSendPdf, chkSendFile, chkSendLink, chkCompress.IsChecked.Value, chkOpen.IsChecked.Value, chkSaveCopy.IsChecked.Value, (bool?)null, (bool?)null, chkDuplex.IsChecked);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		goto IL_0449;
		IL_0449:
		if (e.Cancel)
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
			radScopeEntire.Checked -= ScopeChanged;
			radScopeSelected.Checked -= ScopeChanged;
			chkSendPdf.Checked -= PdfOrFileChecked;
			chkSendPdf.Unchecked -= PdfUnchecked;
			chkSendFile.Checked -= PdfOrFileChecked;
			chkSendFile.Unchecked -= FileUnchecked;
			chkSendLink.Checked -= LinkChecked;
			this.m_A = null;
			this.m_A = null;
			return;
		}
		IL_01cc:
		flag4 = chkSaveCopy.IsChecked;
		if (flag4.HasValue)
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
			if (flag4 != true)
			{
				goto IL_0255;
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
		if (txtFolder.Text.Length == 0)
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
			if (flag4.HasValue)
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
				A(AH.A(105108));
				e.Cancel = true;
				goto IL_039c;
			}
		}
		goto IL_0255;
		IL_0255:
		flag4 = chkSendLink.IsChecked;
		bool? flag6;
		if (!flag4.HasValue)
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
			flag6 = flag4;
		}
		else
		{
			flag6 = flag4 != true;
		}
		flag4 = flag6;
		if (flag4 == true)
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
			flag4 = radScopeEntire.IsChecked;
			if (flag4.HasValue)
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
				if (flag4 != true)
				{
					goto IL_039c;
				}
			}
			if (chkSendFile.IsChecked == true && flag4.HasValue)
			{
				if (this.m_A.Path.Length == 0)
				{
					A(AH.A(105159));
					e.Cancel = true;
				}
				else if (this.m_A.Saved == MsoTriState.msoFalse)
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
					DialogResult dialogResult = System.Windows.Forms.MessageBox.Show(AH.A(105272), AH.A(5874), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
					if (dialogResult != System.Windows.Forms.DialogResult.Cancel)
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
						if (dialogResult == System.Windows.Forms.DialogResult.Yes)
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
							this.m_A.Save();
						}
					}
					else
					{
						e.Cancel = true;
					}
				}
			}
		}
		goto IL_039c;
	}

	private void A(string A)
	{
		Forms.WarningMessage(Window.GetWindow(this), A);
	}

	private void B(string A)
	{
		Forms.ErrorMessage(Window.GetWindow(this), A);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (!this.m_B)
		{
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(105408), UriKind.Relative);
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
				switch (2)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					chkSendPdf = (System.Windows.Controls.CheckBox)target;
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
					chkSendFile = (System.Windows.Controls.CheckBox)target;
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
					chkSendLink = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					radScopeEntire = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					radScopeSelected = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 6)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					grpAttachment = (System.Windows.Controls.GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 7)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					txtName = (System.Windows.Controls.TextBox)target;
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
					chkCompress = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					grpPdf = (System.Windows.Controls.GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkOpen = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			chkSaveCopy = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					txtFolder = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkDuplex = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnOk = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 15)
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
		this.m_B = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
