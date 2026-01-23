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
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Publishing;

[DesignerGenerated]
public sealed class wpfSend : System.Windows.Window, IComponentConnector
{
	private Document m_A;

	private XmlDocument m_A;

	[AccessedThroughProperty("chkSendPdf")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSendFile")]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSendLink")]
	private System.Windows.Controls.CheckBox C;

	[AccessedThroughProperty("radScopeEntire")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("radScopeSelected")]
	private System.Windows.Controls.RadioButton m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("grpAttachment")]
	private System.Windows.Controls.GroupBox m_A;

	[AccessedThroughProperty("txtName")]
	[CompilerGenerated]
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

	[AccessedThroughProperty("txtFolder")]
	[CompilerGenerated]
	private System.Windows.Controls.TextBox m_B;

	[AccessedThroughProperty("btnOk")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnCancel")]
	private System.Windows.Controls.Button m_B;

	private bool m_A;

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

	public wpfSend(Document doc)
	{
		base.Loaded += wpfSend_Loaded;
		base.Closing += wpfSend_Closing;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = doc;
		this.m_A = NC.A.SettingsXml;
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
		if (PC.A.Application.ActiveWindow.Selection.Type == WdSelectionType.wdSelectionNormal)
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
			radScopeSelected.IsChecked = true;
		}
		else
		{
			radScopeEntire.IsChecked = true;
			radScopeSelected.IsEnabled = false;
		}
		clsPublish.PopulateForm(this.m_A, path, name, chkSendPdf, chkSendFile, chkSendLink, txtName, txtFolder);
		XmlNode xmlNode = this.m_A.DocumentElement.SelectSingleNode(XC.A(39833));
		chkCompress.IsChecked = Conversions.ToInteger(xmlNode.SelectSingleNode(XC.A(39850)).InnerText) != 0;
		chkOpen.IsChecked = Conversions.ToInteger(xmlNode.SelectSingleNode(XC.A(39887)).InnerText) != 0;
		chkSaveCopy.IsChecked = Conversions.ToInteger(xmlNode.SelectSingleNode(XC.A(39918)).InnerText) != 0;
		xmlNode = null;
	}

	private void ScopeChanged(object sender, RoutedEventArgs e)
	{
		if (radScopeSelected.IsChecked == true)
		{
			chkSendLink.IsChecked = false;
		}
	}

	private void PdfOrFileChecked(object sender, RoutedEventArgs e)
	{
		chkSendLink.IsChecked = false;
		if (chkSendFile.IsChecked == true)
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
			radScopeSelected.IsEnabled = false;
			radScopeEntire.IsChecked = true;
		}
		else
		{
			radScopeSelected.IsEnabled = true;
		}
		grpAttachment.Visibility = Visibility.Visible;
		if (chkSendPdf.IsChecked == true)
		{
			grpPdf.Visibility = Visibility.Visible;
		}
		else
		{
			grpPdf.Visibility = Visibility.Collapsed;
		}
	}

	private void PdfUnchecked(object sender, RoutedEventArgs e)
	{
		grpPdf.Visibility = Visibility.Collapsed;
		A();
	}

	private void FileUnchecked(object sender, RoutedEventArgs e)
	{
		radScopeSelected.IsEnabled = true;
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
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
		}
		isChecked = chkSendFile.IsChecked;
		if (((!isChecked) ?? isChecked) != true)
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
			if (!flag2.HasValue)
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
				grpAttachment.Visibility = Visibility.Collapsed;
				return;
			}
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
		clsPublish.PickSaveCopyFolder(((_Application)PC.A.Application).get_FileDialog(MsoFileDialogType.msoFileDialogFolderPicker), txtFolder, chkSaveCopy);
	}

	private void btnOk_Click(object sender, RoutedEventArgs e)
	{
		base.DialogResult = true;
	}

	private void wpfSend_Closing(object sender, CancelEventArgs e)
	{
		bool? obj;
		bool? flag4;
		if (base.DialogResult == true)
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
			if (clsPublish.HasIllegalCharacters(txtName, (Action<string>)A))
			{
				e.Cancel = true;
				goto IL_0386;
			}
			bool? isChecked = chkSendPdf.IsChecked;
			bool? flag;
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
				flag = isChecked;
			}
			else
			{
				flag = isChecked != true;
			}
			bool? flag2 = flag;
			isChecked = flag;
			if (isChecked.HasValue)
			{
				if (flag2 != true)
				{
					obj = false;
					goto IL_0130;
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
			bool? flag3;
			if (!isChecked.HasValue)
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
				flag3 = isChecked;
			}
			else
			{
				flag3 = isChecked != true;
			}
			flag4 = flag3;
			isChecked = flag3;
			if (!isChecked.HasValue)
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
				obj = null;
			}
			else if (flag4 != true)
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
				obj = false;
			}
			else
			{
				obj = flag2;
			}
			goto IL_0130;
		}
		goto IL_0433;
		IL_0130:
		bool? flag5 = obj;
		if (flag5.HasValue)
		{
			if (flag5 != true)
			{
				goto IL_01c6;
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
		flag4 = chkSendLink.IsChecked;
		bool? flag6;
		if (!flag4.HasValue)
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
			flag6 = flag4;
		}
		else
		{
			flag6 = flag4 != true;
		}
		flag4 = flag6;
		if (flag4 != true || !flag5.HasValue)
		{
			goto IL_01c6;
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
		A(XC.A(39949));
		e.Cancel = true;
		goto IL_0386;
		IL_0386:
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
			try
			{
				clsPublish.SaveSendSettings(this.m_A, chkSendPdf, chkSendFile, chkSendLink, chkCompress.IsChecked.Value, chkOpen.IsChecked.Value, chkSaveCopy.IsChecked.Value, (bool?)null, (bool?)null, (bool?)null);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		goto IL_0433;
		IL_0433:
		if (e.Cancel)
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
		IL_01c6:
		flag5 = chkSaveCopy.IsChecked;
		if ((flag5 ?? true) && txtFolder.Text.Length == 0)
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
			if (flag5.HasValue)
			{
				A(XC.A(39982));
				e.Cancel = true;
				goto IL_0386;
			}
		}
		flag5 = chkSendLink.IsChecked;
		bool? flag7;
		if (!flag5.HasValue)
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
			flag7 = flag5;
		}
		else
		{
			flag7 = flag5 != true;
		}
		flag5 = flag7;
		if (flag5 == true)
		{
			flag5 = radScopeEntire.IsChecked;
			if (flag5.HasValue)
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
				if (flag5 != true)
				{
					goto IL_0386;
				}
			}
			if (chkSendFile.IsChecked == true)
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
				if (flag5.HasValue)
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
					if (this.m_A.Path.Length == 0)
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
						A(XC.A(40033));
						e.Cancel = true;
					}
					else if (!this.m_A.Saved)
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
						DialogResult dialogResult = System.Windows.Forms.MessageBox.Show(XC.A(40138), XC.A(2438), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
						if (dialogResult != System.Windows.Forms.DialogResult.Cancel)
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
							if (dialogResult == System.Windows.Forms.DialogResult.Yes)
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
		}
		goto IL_0386;
	}

	private void A(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void B(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
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
			switch (7)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = true;
			Uri resourceLocator = new Uri(XC.A(40265), UriKind.Relative);
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
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
		{
			chkSendPdf = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 2)
		{
			chkSendFile = (System.Windows.Controls.CheckBox)target;
			return;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
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
				switch (3)
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
				switch (5)
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
				switch (2)
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
				switch (5)
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
		switch (connectionId)
		{
		case 10:
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				chkOpen = (System.Windows.Controls.CheckBox)target;
				return;
			}
		case 11:
			chkSaveCopy = (System.Windows.Controls.CheckBox)target;
			break;
		case 12:
			txtFolder = (System.Windows.Controls.TextBox)target;
			break;
		case 13:
			btnOk = (System.Windows.Controls.Button)target;
			break;
		case 14:
			btnCancel = (System.Windows.Controls.Button)target;
			break;
		default:
			this.m_A = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}
}
