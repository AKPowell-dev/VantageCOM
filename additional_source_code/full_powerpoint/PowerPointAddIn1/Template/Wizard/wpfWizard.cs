using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
using Foo.Controls;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Agenda;
using PowerPointAddIn1.MasterShapes;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Template.Wizard;

[DesignerGenerated]
public sealed class wpfWizard : Window, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	private Microsoft.Office.Interop.PowerPoint.Presentation m_A;

	private bool m_A;

	private ObservableCollection<TemplateWizardLayout> m_A;

	private ObservableCollection<string> m_A;

	[CompilerGenerated]
	private Settings m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("TabControl1")]
	private System.Windows.Controls.TabControl m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("tabInst")]
	private TabItem m_A;

	[AccessedThroughProperty("tabLayouts")]
	[CompilerGenerated]
	private TabItem m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("lvLayouts")]
	private System.Windows.Controls.ListView m_A;

	[AccessedThroughProperty("gvLayouts")]
	[CompilerGenerated]
	private GridView m_A;

	[AccessedThroughProperty("tabAgenda")]
	[CompilerGenerated]
	private TabItem m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnAgendaGroupsBuild")]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnAgendaGroupsShow")]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnAgendaGroupsHide")]
	private System.Windows.Controls.Button m_C;

	[AccessedThroughProperty("tabShapes")]
	[CompilerGenerated]
	private TabItem m_D;

	[AccessedThroughProperty("lvMasterShapes")]
	[CompilerGenerated]
	private System.Windows.Controls.ListView m_B;

	[AccessedThroughProperty("gvMasterShapes")]
	[CompilerGenerated]
	private GridView m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnMasterShapesShow")]
	private System.Windows.Controls.Button m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnMasterShapesHide")]
	private System.Windows.Controls.Button m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("lvStyles")]
	private System.Windows.Controls.ListView m_C;

	[AccessedThroughProperty("gvStyles")]
	[CompilerGenerated]
	private GridView m_C;

	[AccessedThroughProperty("btnStylesShow")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_F;

	[AccessedThroughProperty("btnStylesHide")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_G;

	[CompilerGenerated]
	[AccessedThroughProperty("tabBrand")]
	private TabItem m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("lbxFonts")]
	private System.Windows.Controls.ListBox m_A;

	[AccessedThroughProperty("btnFontTypeScan")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_H;

	[AccessedThroughProperty("btnFontTypeAdd")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_I;

	[AccessedThroughProperty("btnFontTypeDelete")]
	[CompilerGenerated]
	private System.Windows.Controls.Button J;

	[AccessedThroughProperty("numMinFontSize")]
	[CompilerGenerated]
	private MacNumericUpDown m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("numMaxFontSize")]
	private MacNumericUpDown m_B;

	[AccessedThroughProperty("numTextboxMarginTop")]
	[CompilerGenerated]
	private MacNumericUpDown m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("numTextboxMarginBottom")]
	private MacNumericUpDown m_D;

	[AccessedThroughProperty("numTextboxMarginLeft")]
	[CompilerGenerated]
	private MacNumericUpDown m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("numTextboxMarginRight")]
	private MacNumericUpDown m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("numSlideMarginTop")]
	private MacNumericUpDown m_G;

	[CompilerGenerated]
	[AccessedThroughProperty("numSlideMarginBottom")]
	private MacNumericUpDown m_H;

	[CompilerGenerated]
	[AccessedThroughProperty("numSlideMarginLeft")]
	private MacNumericUpDown m_I;

	[AccessedThroughProperty("numSlideMarginRight")]
	[CompilerGenerated]
	private MacNumericUpDown J;

	[AccessedThroughProperty("tabValidate")]
	[CompilerGenerated]
	private TabItem m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("btnValidate")]
	private System.Windows.Controls.Button K;

	[AccessedThroughProperty("lbxValidate")]
	[CompilerGenerated]
	private System.Windows.Controls.ListBox m_B;

	[AccessedThroughProperty("btnSelectionPane")]
	[CompilerGenerated]
	private System.Windows.Controls.Button L;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private System.Windows.Controls.Button M;

	[AccessedThroughProperty("btnRefresh")]
	[CompilerGenerated]
	private System.Windows.Controls.Button N;

	private bool m_B;

	public ObservableCollection<TemplateWizardLayout> LayoutsCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(122428));
		}
	}

	public ObservableCollection<string> LegalFontTypes
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(122463));
		}
	}

	private Settings BrandSettings
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

	internal virtual System.Windows.Controls.TabControl TabControl1
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
			SelectionChangedEventHandler value2 = TabControl1_SelectionChanged;
			System.Windows.Controls.TabControl tabControl = this.m_A;
			if (tabControl != null)
			{
				tabControl.SelectionChanged -= value2;
			}
			this.m_A = value;
			tabControl = this.m_A;
			if (tabControl == null)
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
				tabControl.SelectionChanged += value2;
				return;
			}
		}
	}

	internal virtual TabItem tabInst
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

	internal virtual TabItem tabLayouts
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

	internal virtual System.Windows.Controls.ListView lvLayouts
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

	internal virtual GridView gvLayouts
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

	internal virtual TabItem tabAgenda
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
			this.m_C = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnAgendaGroupsBuild
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
			RoutedEventHandler value2 = btnAgendaGroupsBuild_Click;
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

	internal virtual System.Windows.Controls.Button btnAgendaGroupsShow
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
			RoutedEventHandler value2 = btnAgendaGroupsShow_Click;
			System.Windows.Controls.Button button = this.m_B;
			if (button != null)
			{
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

	internal virtual System.Windows.Controls.Button btnAgendaGroupsHide
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
			RoutedEventHandler value2 = btnAgendaGroupsHide_Click;
			System.Windows.Controls.Button button = this.m_C;
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
			this.m_C = value;
			button = this.m_C;
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

	internal virtual TabItem tabShapes
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
			this.m_D = value;
		}
	}

	internal virtual System.Windows.Controls.ListView lvMasterShapes
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
			SelectionChangedEventHandler value2 = lvMasterShapes_SelectionChanged;
			System.Windows.Controls.ListView listView = this.m_B;
			if (listView != null)
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
				listView.SelectionChanged -= value2;
			}
			this.m_B = value;
			listView = this.m_B;
			if (listView != null)
			{
				listView.SelectionChanged += value2;
			}
		}
	}

	internal virtual GridView gvMasterShapes
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

	internal virtual System.Windows.Controls.Button btnMasterShapesShow
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
			RoutedEventHandler value2 = btnMasterShapesShow_Click;
			System.Windows.Controls.Button button = this.m_D;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_D = value;
			button = this.m_D;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnMasterShapesHide
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
			RoutedEventHandler value2 = btnMasterShapesHide_Click;
			System.Windows.Controls.Button button = this.m_E;
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
				button.Click -= value2;
			}
			this.m_E = value;
			button = this.m_E;
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

	internal virtual System.Windows.Controls.ListView lvStyles
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
			SelectionChangedEventHandler value2 = lvStyles_SelectionChanged;
			System.Windows.Controls.ListView listView = this.m_C;
			if (listView != null)
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
				listView.SelectionChanged -= value2;
			}
			this.m_C = value;
			listView = this.m_C;
			if (listView == null)
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
				listView.SelectionChanged += value2;
				return;
			}
		}
	}

	internal virtual GridView gvStyles
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
			this.m_C = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnStylesShow
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
			RoutedEventHandler value2 = btnStylesShow_Click;
			System.Windows.Controls.Button button = this.m_F;
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

	internal virtual System.Windows.Controls.Button btnStylesHide
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
			RoutedEventHandler value2 = btnStylesHide_Click;
			System.Windows.Controls.Button button = this.m_G;
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
			this.m_G = value;
			button = this.m_G;
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

	internal virtual TabItem tabBrand
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
			this.m_E = value;
		}
	}

	internal virtual System.Windows.Controls.ListBox lbxFonts
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

	internal virtual System.Windows.Controls.Button btnFontTypeScan
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
			RoutedEventHandler value2 = FontTypeScan;
			System.Windows.Controls.Button button = this.m_H;
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
			this.m_H = value;
			button = this.m_H;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnFontTypeAdd
	{
		[CompilerGenerated]
		get
		{
			return this.m_I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = FontTypeAdd;
			System.Windows.Controls.Button button = this.m_I;
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
			this.m_I = value;
			button = this.m_I;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnFontTypeDelete
	{
		[CompilerGenerated]
		get
		{
			return this.J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = FontTypeDelete;
			System.Windows.Controls.Button button = this.J;
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
			this.J = value;
			button = this.J;
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

	internal virtual MacNumericUpDown numMinFontSize
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

	internal virtual MacNumericUpDown numMaxFontSize
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

	internal virtual MacNumericUpDown numTextboxMarginTop
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
			this.m_C = value;
		}
	}

	internal virtual MacNumericUpDown numTextboxMarginBottom
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
			this.m_D = value;
		}
	}

	internal virtual MacNumericUpDown numTextboxMarginLeft
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
			this.m_E = value;
		}
	}

	internal virtual MacNumericUpDown numTextboxMarginRight
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
			this.m_F = value;
		}
	}

	internal virtual MacNumericUpDown numSlideMarginTop
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
			this.m_G = value;
		}
	}

	internal virtual MacNumericUpDown numSlideMarginBottom
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
			this.m_H = value;
		}
	}

	internal virtual MacNumericUpDown numSlideMarginLeft
	{
		[CompilerGenerated]
		get
		{
			return this.m_I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_I = value;
		}
	}

	internal virtual MacNumericUpDown numSlideMarginRight
	{
		[CompilerGenerated]
		get
		{
			return J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			J = value;
		}
	}

	internal virtual TabItem tabValidate
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
			this.m_F = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnValidate
	{
		[CompilerGenerated]
		get
		{
			return K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnValidate_Click;
			System.Windows.Controls.Button button = K;
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
				button.Click -= value2;
			}
			K = value;
			button = K;
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

	internal virtual System.Windows.Controls.ListBox lbxValidate
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

	internal virtual System.Windows.Controls.Button btnSelectionPane
	{
		[CompilerGenerated]
		get
		{
			return L;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnSelectionPane_Click;
			System.Windows.Controls.Button button = L;
			if (button != null)
			{
				button.Click -= value2;
			}
			L = value;
			button = L;
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

	internal virtual System.Windows.Controls.Button btnClose
	{
		[CompilerGenerated]
		get
		{
			return M;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnClose_Click;
			System.Windows.Controls.Button button = M;
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
			M = value;
			button = M;
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnRefresh
	{
		[CompilerGenerated]
		get
		{
			return N;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnRefresh_Click;
			System.Windows.Controls.Button button = N;
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
				button.Click -= value2;
			}
			N = value;
			button = N;
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
				switch (3)
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

	public wpfWizard(string strPath)
	{
		base.Loaded += wpfTemplateWizard_Loaded;
		base.Closing += wpfTemplateWizard_Closing;
		this.m_A = false;
		this.m_A = null;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
		this.m_A = NG.A.Application;
		try
		{
			try
			{
				this.m_A = this.m_A.Presentations[Path.GetFileName(strPath)];
				if (this.m_A == null)
				{
					throw new Exception();
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				this.m_A = this.m_A.Presentations.Open(strPath);
				if (this.m_A == null)
				{
					throw new Exception();
				}
				this.m_A = true;
				ProjectData.ClearProjectError();
			}
			I();
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			Close();
			ProjectData.ClearProjectError();
		}
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private void wpfTemplateWizard_Loaded(object sender, RoutedEventArgs e)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).AddEventHandler(this.m_A, new EApplication_PresentationCloseFinalEventHandler(A));
		base.MinHeight = base.ActualHeight;
		base.MinWidth = base.ActualWidth;
	}

	private void wpfTemplateWizard_Closing(object sender, CancelEventArgs e)
	{
		if (this.m_A)
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
			Microsoft.Office.Interop.PowerPoint.Presentation a = this.m_A;
			if (a.Saved == MsoTriState.msoFalse)
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
				if (a.ReadOnly == MsoTriState.msoFalse)
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
					if (Forms.YesNoMessage(AH.A(122492)) == System.Windows.Forms.DialogResult.Yes)
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
						a.Save();
						a.Saved = MsoTriState.msoTrue;
					}
				}
			}
			a.Close();
			a = null;
		}
		this.m_A = null;
		lvLayouts.SelectionChanged -= lvLayouts_SelectionChanged;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(10098)).RemoveEventHandler(this.m_A, new EApplication_PresentationCloseFinalEventHandler(A));
		this.m_A = null;
		LayoutsCollection = null;
	}

	private void TabControl1_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (e.Source != TabControl1)
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
			A();
			return;
		}
	}

	private void btnRefresh_Click(object sender, RoutedEventArgs e)
	{
		A();
		lbxValidate.ItemsSource = null;
		lbxValidate.Visibility = Visibility.Hidden;
	}

	private void A()
	{
		if (tabLayouts.IsSelected)
		{
			B();
		}
		else if (tabShapes.IsSelected)
		{
			C();
			D();
		}
		else
		{
			if (!tabBrand.IsSelected)
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
				F();
				return;
			}
		}
	}

	private void btnSelectionPane_Click(object sender, RoutedEventArgs e)
	{
		if (!this.m_A.CommandBars.GetPressedMso(AH.A(91479)))
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
					this.m_A.CommandBars.ExecuteMso(AH.A(91479));
					System.Windows.Forms.Application.DoEvents();
					return;
				}
			}
		}
		E(AH.A(122622));
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		Close();
	}

	private void B()
	{
		ObservableCollection<TemplateWizardLayout> observableCollection = new ObservableCollection<TemplateWizardLayout>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A().GetEnumerator();
			while (enumerator.MoveNext())
			{
				CustomLayout lay = (CustomLayout)enumerator.Current;
				observableCollection.Add(new TemplateWizardLayout(lay));
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
				break;
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
		lvLayouts.SelectionChanged -= lvLayouts_SelectionChanged;
		LayoutsCollection = observableCollection;
		lvLayouts.SelectionChanged += lvLayouts_SelectionChanged;
		JG.A(observableCollection);
		observableCollection = null;
	}

	private void lvLayouts_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (lvLayouts.SelectedIndex <= -1)
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
			I();
			A(checked(lvLayouts.SelectedIndex + 1)).Select();
			Activate();
			return;
		}
	}

	private void OnDropDownOpened(object sender, EventArgs e)
	{
		System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
		lvLayouts.SelectedItem = (TemplateWizardLayout)comboBox.DataContext;
		comboBox.Focus();
		comboBox.IsDropDownOpen = true;
		comboBox = null;
	}

	private void RenameLayout(object sender, RoutedEventArgs e)
	{
		TemplateWizardLayout templateWizardLayout = (TemplateWizardLayout)((System.Windows.Controls.Button)sender).DataContext;
		string name = templateWizardLayout.Name;
		string text = Forms.InputBox(Window.GetWindow(this), AH.A(122685), AH.A(122708), name);
		if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
				if (Operators.CompareString(text, name, TextCompare: false) != 0)
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
					templateWizardLayout.Name = text;
					A(checked(LayoutsCollection.IndexOf(templateWizardLayout) + 1)).Name = text;
					A(gvLayouts);
				}
			}
		}
		templateWizardLayout = null;
	}

	private void LayoutRoleChanged(object sender, SelectionChangedEventArgs e)
	{
		System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
		if (!comboBox.IsLoaded)
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
					comboBox = null;
					return;
				}
			}
		}
		SlideType slideType = default(SlideType);
		switch (comboBox.SelectedIndex)
		{
		case 0:
			slideType = SlideType.Content;
			break;
		case 1:
			slideType = SlideType.Title;
			break;
		case 2:
			slideType = SlideType.TableOfContents;
			break;
		case 3:
			slideType = SlideType.Flysheet;
			break;
		case 4:
			slideType = SlideType.Legal;
			break;
		case 5:
			slideType = SlideType.Contact;
			break;
		case 6:
			slideType = SlideType.Blank;
			break;
		case 7:
			slideType = SlideType.CoverFront;
			break;
		case 8:
			slideType = SlideType.CoverBack;
			break;
		}
		int num = (int)slideType;
		string right = num.ToString();
		if (slideType != SlideType.Content)
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
			using IEnumerator<TemplateWizardLayout> enumerator = LayoutsCollection.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				TemplateWizardLayout current = enumerator.Current;
				{
					enumerator2 = current.Layout.Shapes.GetEnumerator();
					try
					{
						while (true)
						{
							if (enumerator2.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
								try
								{
									Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
									if (shape2.Visible == MsoTriState.msoFalse)
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
										if (Operators.CompareString(shape2.Tags[PowerPointAddIn1.Slides.Helpers.TAG_SLIDE_TYPE], right, TextCompare: false) == 0)
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												shape2.Delete();
												current.Index = 0;
												break;
											}
											break;
										}
									}
									shape2 = null;
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
								continue;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									break;
								default:
									goto end_IL_0161;
								}
								continue;
								end_IL_0161:
								break;
							}
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
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_0191;
				}
				continue;
				end_IL_0191:
				break;
			}
		}
		CustomLayout layout = ((TemplateWizardLayout)comboBox.DataContext).Layout;
		try
		{
			for (int i = layout.Shapes.Count; i >= 1; i = checked(i + -1))
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape3 = layout.Shapes[i];
				if (shape3.Visible == MsoTriState.msoFalse)
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
					if (Operators.CompareString(shape3.Tags[PowerPointAddIn1.Slides.Helpers.TAG_SLIDE_TYPE], "", TextCompare: false) != 0)
					{
						shape3.Delete();
					}
				}
				shape3 = null;
			}
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					goto end_IL_022d;
				}
				continue;
				end_IL_022d:
				break;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		PowerPointAddIn1.Slides.Helpers.DesignateLayoutAsType(layout, slideType);
		comboBox = null;
		layout = null;
	}

	private void btnAgendaGroupsBuild_Click(object sender, RoutedEventArgs e)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		bool flag = false;
		bool flag2 = false;
		CustomLayout customLayout = A();
		IEnumerator enumerator = default(IEnumerator);
		if (customLayout != null)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					int num = customLayout.Shapes.Count;
					while (true)
					{
						if (num < 1)
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
							break;
						}
						Microsoft.Office.Interop.PowerPoint.Shape shape = customLayout.Shapes[num];
						string name = shape.Name;
						if (Operators.CompareString(name, Constants.AGENDA_TITLE_LEVEL_1, TextCompare: false) != 0)
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
							if (Operators.CompareString(name, Constants.AGENDA_TITLE_LEVEL_1_ACTIVE, TextCompare: false) != 0)
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
								if (Operators.CompareString(name, Constants.AGENDA_TITLE_LEVEL_2, TextCompare: false) != 0)
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
									if (Operators.CompareString(name, Constants.AGENDA_TITLE_LEVEL_2_ACTIVE, TextCompare: false) != 0)
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
										if (Operators.CompareString(name, Constants.AGENDA_TITLE_LEVEL_3, TextCompare: false) != 0)
										{
											goto IL_0115;
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
							}
						}
						if (!flag)
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
							if (Forms.OkCancelMessage(AH.A(122775)) == System.Windows.Forms.DialogResult.Cancel)
							{
								flag2 = true;
								break;
							}
							flag = true;
						}
						shape.Delete();
						goto IL_0115;
						IL_0115:
						shape = null;
						num = checked(num + -1);
					}
					if (!flag2)
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
						List<string> list = new List<string>();
						try
						{
							presentation = A();
							try
							{
								enumerator = presentation.Designs[1].SlideMaster.CustomLayouts[2].Shapes.GetEnumerator();
								while (enumerator.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
									if (shape2.Type == MsoShapeType.msoGroup)
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
										if (shape2.Name.StartsWith(AH.A(122951)))
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
											shape2.Copy();
											Microsoft.Office.Interop.PowerPoint.Shape shape3 = customLayout.Shapes.Paste()[1];
											shape3.Top = shape2.Top;
											shape3.Left = shape2.Left;
											shape3.Visible = MsoTriState.msoTrue;
											list.Add(shape3.Name);
											shape3 = null;
										}
									}
								}
								while (true)
								{
									switch (6)
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
							}
							finally
							{
								if (enumerator is IDisposable)
								{
									while (true)
									{
										switch (2)
										{
										case 0:
											break;
										default:
											(enumerator as IDisposable).Dispose();
											goto end_IL_025a;
										}
										continue;
										end_IL_025a:
										break;
									}
								}
							}
							try
							{
								customLayout.Shapes.Range(list.ToArray()).Select();
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								ProjectData.ClearProjectError();
							}
							E(AH.A(122964));
							Activate();
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							try
							{
								presentation.Close();
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								ProjectData.ClearProjectError();
							}
							C(AH.A(123184));
							ProjectData.ClearProjectError();
						}
						finally
						{
							JG.A(presentation);
							presentation = null;
						}
						list = null;
					}
					customLayout = null;
					return;
				}
				}
			}
		}
		D(AH.A(123314));
	}

	private void btnAgendaGroupsHide_Click(object sender, RoutedEventArgs e)
	{
		A(MsoTriState.msoFalse);
	}

	private void btnAgendaGroupsShow_Click(object sender, RoutedEventArgs e)
	{
		A(MsoTriState.msoTrue);
	}

	private void A(MsoTriState A)
	{
		CustomLayout customLayout = this.A();
		IEnumerator enumerator = default(IEnumerator);
		if (customLayout != null)
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
					try
					{
						enumerator = customLayout.Shapes.GetEnumerator();
						while (enumerator.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
							string name = shape.Name;
							if (Operators.CompareString(name, Constants.AGENDA_TITLE_LEVEL_1, TextCompare: false) != 0)
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
								if (Operators.CompareString(name, Constants.AGENDA_TITLE_LEVEL_1_ACTIVE, TextCompare: false) != 0 && Operators.CompareString(name, Constants.AGENDA_TITLE_LEVEL_2, TextCompare: false) != 0)
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
									if (Operators.CompareString(name, Constants.AGENDA_TITLE_LEVEL_2_ACTIVE, TextCompare: false) != 0)
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
										if (Operators.CompareString(name, Constants.AGENDA_TITLE_LEVEL_3, TextCompare: false) != 0)
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
									}
								}
							}
							shape.Visible = A;
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
									goto end_IL_00e0;
								}
								continue;
								end_IL_00e0:
								break;
							}
						}
					}
					customLayout = null;
					return;
				}
			}
		}
		D(AH.A(123314));
	}

	private CustomLayout A()
	{
		CustomLayout result = null;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A().GetEnumerator();
			while (true)
			{
				if (enumerator.MoveNext())
				{
					CustomLayout customLayout = (CustomLayout)enumerator.Current;
					if (!new SlideType[2]
					{
						SlideType.TableOfContents,
						SlideType.Agenda
					}.Contains(PowerPointAddIn1.Slides.Helpers.GetLayoutType(customLayout)))
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						result = customLayout;
						I();
						customLayout.Select();
						Activate();
						break;
					}
					break;
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_0071;
					}
					continue;
					end_IL_0071:
					break;
				}
				break;
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
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		return result;
	}

	private void BuildSectDivPlaceholder(object sender, RoutedEventArgs e)
	{
		bool flag = false;
		bool flag2 = false;
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A().GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			IEnumerator enumerator3 = default(IEnumerator);
			IEnumerator enumerator4 = default(IEnumerator);
			while (true)
			{
				if (enumerator.MoveNext())
				{
					CustomLayout customLayout = (CustomLayout)enumerator.Current;
					if (PowerPointAddIn1.Slides.Helpers.GetLayoutType(customLayout) != SlideType.Flysheet)
					{
						continue;
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
						try
						{
							customLayout.Select();
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							I();
							customLayout.Select();
							ProjectData.ClearProjectError();
						}
						enumerator2 = customLayout.Shapes.Placeholders.GetEnumerator();
						try
						{
							while (true)
							{
								if (enumerator2.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
									if (shape.PlaceholderFormat.Type != PpPlaceholderType.ppPlaceholderBody)
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
										shape.Select();
										flag2 = true;
										break;
									}
									break;
								}
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										goto end_IL_00cd;
									}
									continue;
									end_IL_00cd:
									break;
								}
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
						if (flag2)
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
							D(AH.A(123395));
						}
						else
						{
							try
							{
								presentation = A();
								enumerator3 = presentation.Designs[1].SlideMaster.CustomLayouts[3].Shapes.GetEnumerator();
								try
								{
									while (enumerator3.MoveNext())
									{
										Microsoft.Office.Interop.PowerPoint.Shape shape2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current;
										if (shape2.Type != MsoShapeType.msoPlaceholder)
										{
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
										if (shape2.PlaceholderFormat.Type != PpPlaceholderType.ppPlaceholderBody)
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
										this.m_A.StartNewUndoEntry();
										shape2.Copy();
										Microsoft.Office.Interop.PowerPoint.Shape shape3 = customLayout.Shapes.Paste()[1];
										Microsoft.Office.Interop.PowerPoint.Shape shape4 = shape3;
										shape4.Top = shape2.Top;
										shape4.Left = shape2.Left;
										shape4.Top = (customLayout.Height - shape4.Height) / 2f;
										shape4.Select();
										shape4 = null;
										try
										{
											enumerator4 = A().Shapes.Placeholders.GetEnumerator();
											while (enumerator4.MoveNext())
											{
												Microsoft.Office.Interop.PowerPoint.Shape shape5 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current;
												if (shape5.PlaceholderFormat.Type != PpPlaceholderType.ppPlaceholderBody)
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
													shape3.Left = shape5.Left;
													shape3.Width = shape5.Width;
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
													switch (3)
													{
													case 0:
														continue;
													}
													(enumerator4 as IDisposable).Dispose();
													break;
												}
											}
										}
										shape3 = null;
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_02c3;
										}
										continue;
										end_IL_02c3:
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
								E(AH.A(123522));
								Activate();
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								try
								{
									presentation.Close();
								}
								catch (Exception ex5)
								{
									ProjectData.SetProjectError(ex5);
									Exception ex6 = ex5;
									ProjectData.ClearProjectError();
								}
								C(AH.A(123184));
								ProjectData.ClearProjectError();
							}
							finally
							{
								JG.A(presentation);
								presentation = null;
							}
						}
						flag = true;
						break;
					}
					break;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0353;
					}
					continue;
					end_IL_0353:
					break;
				}
				break;
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
		if (flag)
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
			D(AH.A(123611));
			return;
		}
	}

	private void SelectSectDivPlaceholder(object sender, RoutedEventArgs e)
	{
		bool flag = false;
		bool flag2 = false;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A().GetEnumerator();
			while (enumerator.MoveNext())
			{
				CustomLayout customLayout = (CustomLayout)enumerator.Current;
				if (PowerPointAddIn1.Slides.Helpers.GetLayoutType(customLayout) != SlideType.Flysheet)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					foreach (Microsoft.Office.Interop.PowerPoint.Shape placeholder in customLayout.Shapes.Placeholders)
					{
						if (placeholder.PlaceholderFormat.Type != PpPlaceholderType.ppPlaceholderBody)
						{
							continue;
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							try
							{
								customLayout.Select();
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								I();
								customLayout.Select();
								ProjectData.ClearProjectError();
							}
							placeholder.Select();
							flag2 = true;
							break;
						}
						break;
					}
					flag = true;
					break;
				}
				break;
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
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		if (!flag)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					D(AH.A(123611));
					return;
				}
			}
		}
		if (flag2)
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
			D(AH.A(123710));
			return;
		}
	}

	private void C()
	{
		ObservableCollection<MasterShapeItem> observableCollection = new ObservableCollection<MasterShapeItem>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A().Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (A(shape))
				{
					observableCollection.Add(new MasterShapeItem(shape));
				}
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
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		lvMasterShapes.ItemsSource = observableCollection;
		A(gvMasterShapes);
		JG.A(observableCollection);
		observableCollection = null;
		btnMasterShapesShow.IsEnabled = lvMasterShapes.Items.Count > 0;
		btnMasterShapesHide.IsEnabled = btnMasterShapesShow.IsEnabled;
	}

	private void D()
	{
		ObservableCollection<StyleShapeItem> observableCollection = new ObservableCollection<StyleShapeItem>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A().Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (!B(shape))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				observableCollection.Add(new StyleShapeItem(shape));
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
		lvStyles.ItemsSource = observableCollection;
		A(gvStyles);
		JG.A(observableCollection);
		observableCollection = null;
		btnStylesShow.IsEnabled = lvStyles.Items.Count > 0;
		btnStylesHide.IsEnabled = btnStylesShow.IsEnabled;
	}

	private void lvMasterShapes_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (lvMasterShapes.SelectedIndex == -1)
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
			Microsoft.Office.Interop.PowerPoint.Shape shape = ((MasterShapeItem)lvMasterShapes.SelectedItem).Shape;
			if (shape.Visible == MsoTriState.msoTrue)
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
					A(shape);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					try
					{
						E();
						A(shape);
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					ProjectData.ClearProjectError();
				}
			}
			shape = null;
			return;
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		A.Select();
		lvMasterShapes.Focus();
	}

	private void lvStyles_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (lvStyles.SelectedIndex == -1)
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
			Microsoft.Office.Interop.PowerPoint.Shape shape = ((StyleShapeItem)lvStyles.SelectedItem).Shape;
			if (shape.Visible == MsoTriState.msoTrue)
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
					B(shape);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					try
					{
						E();
						B(shape);
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						ProjectData.ClearProjectError();
					}
					ProjectData.ClearProjectError();
				}
			}
			shape = null;
			return;
		}
	}

	private void B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		A.Select();
		lvStyles.Focus();
	}

	private void btnMasterShapesShow_Click(object sender, RoutedEventArgs e)
	{
		B(MsoTriState.msoTrue);
		E(AH.A(123823));
	}

	private void btnMasterShapesHide_Click(object sender, RoutedEventArgs e)
	{
		B(MsoTriState.msoFalse);
	}

	private void B(MsoTriState A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = this.A().Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (!this.A(shape))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				shape.Visible = A;
			}
			while (true)
			{
				switch (1)
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

	private void btnStylesShow_Click(object sender, RoutedEventArgs e)
	{
		C(MsoTriState.msoTrue);
		E(AH.A(123823));
	}

	private void btnStylesHide_Click(object sender, RoutedEventArgs e)
	{
		C(MsoTriState.msoFalse);
	}

	private void C(MsoTriState A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = this.A().Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				if (!B(shape))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				shape.Visible = A;
			}
			while (true)
			{
				switch (5)
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
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
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

	private bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (PowerPointAddIn1.MasterShapes.Base.A(A))
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
					return A.Type != MsoShapeType.msoPlaceholder;
				}
			}
		}
		return false;
	}

	private bool B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		return Styles.HasStyleShapeName(A);
	}

	private void E()
	{
		this.m_A.ActiveWindow.ViewType = PpViewType.ppViewMasterThumbnails;
	}

	private void F()
	{
		//IL_0032: Unknown result type (might be due to invalid IL or missing references)
		//IL_0037: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_004a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0052: Unknown result type (might be due to invalid IL or missing references)
		if (BrandSettings != null)
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
			Settings settings = new Settings(this.m_A);
			SpinnerProperties spinnerProps = Regional.GetSpinnerProps((bool?)null);
			BrandSettings = settings;
			A(settings);
			A(settings, spinnerProps);
			B(settings, spinnerProps);
			settings = null;
			return;
		}
	}

	private void A(Settings A)
	{
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_007a: Expected O, but got Unknown
		//IL_00ce: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d8: Expected O, but got Unknown
		//IL_00ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f7: Expected O, but got Unknown
		//IL_014c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0156: Expected O, but got Unknown
		LegalFontTypes = new ObservableCollection<string>();
		using (List<string>.Enumerator enumerator = A.LegalFontTypes.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				string current = enumerator.Current;
				LegalFontTypes.Add(current);
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
				break;
			}
		}
		MacNumericUpDown val = numMinFontSize;
		val.ValueChanged -= new MacRangeBaseValueChangedHandler(MinFontSizeChanged);
		if (A.MinFontSize.HasValue)
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
			val.Value = A.MinFontSize.Value;
		}
		else
		{
			val.Value = null;
		}
		val.ValueChanged += new MacRangeBaseValueChangedHandler(MinFontSizeChanged);
		val = null;
		MacNumericUpDown val2 = numMaxFontSize;
		val2.ValueChanged -= new MacRangeBaseValueChangedHandler(MaxFontSizeChanged);
		if (A.MaxFontSize.HasValue)
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
			val2.Value = A.MaxFontSize.Value;
		}
		else
		{
			val2.Value = null;
		}
		val2.ValueChanged += new MacRangeBaseValueChangedHandler(MaxFontSizeChanged);
		val2 = null;
	}

	private void FontTypeAdd(object sender, RoutedEventArgs e)
	{
		string text = Forms.InputBox2(AH.A(123955), AH.A(123984), "");
		if (text == null)
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
			if (text.Length <= 0)
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
				LegalFontTypes.Add(text);
				G();
				return;
			}
		}
	}

	private void FontTypeDelete(object sender, RoutedEventArgs e)
	{
		if (lbxFonts.SelectedIndex <= -1)
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
			LegalFontTypes.RemoveAt(lbxFonts.SelectedIndex);
			G();
			return;
		}
	}

	private void FontTypeScan(object sender, RoutedEventArgs e)
	{
		List<string> B = new List<string>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = this.m_A.Designs.GetEnumerator();
			IEnumerator enumerator3 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Design design = (Design)enumerator.Current;
				foreach (CustomLayout customLayout in design.SlideMaster.CustomLayouts)
				{
					try
					{
						enumerator3 = customLayout.Shapes.GetEnumerator();
						while (enumerator3.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator3.Current;
							A(a, ref B);
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
							break;
						}
					}
					finally
					{
						if (enumerator3 is IDisposable)
						{
							while (true)
							{
								switch (7)
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
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					goto end_IL_00f5;
				}
				continue;
				end_IL_00f5:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		IEnumerator enumerator4 = this.m_A.Slides.GetEnumerator();
		try
		{
			IEnumerator enumerator5 = default(IEnumerator);
			while (enumerator4.MoveNext())
			{
				Slide slide = (Slide)enumerator4.Current;
				try
				{
					enumerator5 = slide.Shapes.GetEnumerator();
					while (enumerator5.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape a2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator5.Current;
						A(a2, ref B);
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_017d;
						}
						continue;
						end_IL_017d:
						break;
					}
				}
				finally
				{
					if (enumerator5 is IDisposable)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							(enumerator5 as IDisposable).Dispose();
							break;
						}
					}
				}
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_01b4;
				}
				continue;
				end_IL_01b4:
				break;
			}
		}
		finally
		{
			IDisposable disposable = enumerator4 as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
		LegalFontTypes = new ObservableCollection<string>();
		B = B.Distinct().ToList();
		using (List<string>.Enumerator enumerator6 = B.GetEnumerator())
		{
			while (enumerator6.MoveNext())
			{
				string current = enumerator6.Current;
				LegalFontTypes.Add(current);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_0220;
				}
				continue;
				end_IL_0220:
				break;
			}
		}
		B = null;
		G();
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, ref List<string> B)
	{
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			if (A.Type != MsoShapeType.msoGroup)
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
						if (A.HasTextFrame == MsoTriState.msoTrue)
						{
							this.A(A.TextFrame2.TextRange, ref B);
						}
						else
						{
							if (A.HasTable == MsoTriState.msoTrue)
							{
								while (true)
								{
									switch (5)
									{
									case 0:
										break;
									default:
									{
										Table table = A.Table;
										int count = table.Rows.Count;
										for (int i = 1; i <= count; i++)
										{
											int count2 = table.Columns.Count;
											for (int j = 1; j <= count2; j++)
											{
												Microsoft.Office.Interop.PowerPoint.TextFrame2 textFrame = table.Cell(i, j).Shape.TextFrame2;
												if (textFrame.HasText == MsoTriState.msoTrue)
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
													this.A(textFrame.TextRange, ref B);
												}
												textFrame = null;
											}
											while (true)
											{
												switch (2)
												{
												case 0:
													break;
												default:
													goto end_IL_00d2;
												}
												continue;
												end_IL_00d2:
												break;
											}
										}
										while (true)
										{
											switch (4)
											{
											case 0:
												break;
											default:
												table = null;
												return;
											}
										}
									}
									}
								}
							}
							if (A.HasSmartArt == MsoTriState.msoTrue)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
									{
										enumerator = A.SmartArt.AllNodes.GetEnumerator();
										try
										{
											while (enumerator.MoveNext())
											{
												SmartArtNode smartArtNode = (SmartArtNode)enumerator.Current;
												this.A(smartArtNode.TextFrame2.TextRange, ref B);
											}
											while (true)
											{
												switch (6)
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
											IDisposable disposable = enumerator as IDisposable;
											if (disposable != null)
											{
												disposable.Dispose();
											}
										}
									}
									}
								}
							}
						}
						return;
					}
				}
			}
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = A.GroupItems.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
					this.A(a, ref B);
				}
				while (true)
				{
					switch (6)
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
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (2)
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
	}

	private void A(TextRange2 A, ref List<string> B)
	{
		IEnumerator enumerator = A.get_Runs(-1, -1).GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				TextRange2 textRange = (TextRange2)enumerator.Current;
				B.Add(textRange.Font.Name);
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
				return;
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
	}

	private void MinFontSizeChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		if (numMinFontSize.Value.HasValue)
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
			BrandSettings.MinFontSize = checked((int)Math.Round(numMinFontSize.Value.Value));
		}
		else
		{
			BrandSettings.MinFontSize = null;
		}
		G();
	}

	private void MaxFontSizeChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		if (numMaxFontSize.Value.HasValue)
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
			BrandSettings.MaxFontSize = checked((int)Math.Round(numMaxFontSize.Value.Value));
		}
		else
		{
			BrandSettings.MaxFontSize = null;
		}
		G();
	}

	private void G()
	{
		BrandSettings.LegalFontTypes = LegalFontTypes.ToList();
		BrandSettings.A(this.m_A);
	}

	private MacNumericUpDown[] A()
	{
		return (MacNumericUpDown[])(object)new MacNumericUpDown[4] { numTextboxMarginTop, numTextboxMarginBottom, numTextboxMarginLeft, numTextboxMarginRight };
	}

	private void A(Settings A, SpinnerProperties B)
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_0042: Unknown result type (might be due to invalid IL or missing references)
		MacNumericUpDown[] array = this.A();
		foreach (MacNumericUpDown val in array)
		{
			this.A(val);
			val.CustomUnit = B.CustomUnits;
			val.NumberDecimalDigits = B.Decimals;
			val.SmallChange = B.SmallChange;
			val.LargeChange = B.LargeChange;
			int num;
			if (Operators.CompareString(val.CustomUnit, AH.A(69068), TextCompare: false) != 0)
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
				num = 5;
			}
			else
			{
				num = 2;
			}
			val.Maximum = num;
			_ = null;
		}
		if (A.TextboxMargins.HasValue)
		{
			Settings.Margins value = A.TextboxMargins.Value;
			if (Operators.CompareString(numTextboxMarginTop.CustomUnit, AH.A(69068), TextCompare: false) == 0)
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
				numTextboxMarginTop.Value = clsPublish.PointsToInches(value.Top);
				numTextboxMarginBottom.Value = clsPublish.PointsToInches(value.Bottom);
				numTextboxMarginLeft.Value = clsPublish.PointsToInches(value.Left);
				numTextboxMarginRight.Value = clsPublish.PointsToInches(value.Right);
			}
			else
			{
				numTextboxMarginTop.Value = clsPublish.PointsToCentimeters(value.Top);
				numTextboxMarginBottom.Value = clsPublish.PointsToCentimeters(value.Bottom);
				numTextboxMarginLeft.Value = clsPublish.PointsToCentimeters(value.Left);
				numTextboxMarginRight.Value = clsPublish.PointsToCentimeters(value.Right);
			}
		}
		else
		{
			MacNumericUpDown[] array2 = this.A();
			for (int j = 0; j < array2.Length; j = checked(j + 1))
			{
				array2[j].Value = null;
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
		}
		MacNumericUpDown[] array3 = this.A();
		foreach (MacNumericUpDown a in array3)
		{
			this.B(a);
		}
		while (true)
		{
			switch (3)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void A(MacNumericUpDown A)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		A.ValueChanged -= new MacRangeBaseValueChangedHandler(TextboxMarginsChanged);
	}

	private void B(MacNumericUpDown A)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		A.ValueChanged += new MacRangeBaseValueChangedHandler(TextboxMarginsChanged);
	}

	private void TextboxMarginsChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		int num = 0;
		double? value = numTextboxMarginTop.Value;
		checked
		{
			float num2 = default(float);
			if (value.HasValue)
			{
				num2 = (float)value.Value;
				num++;
			}
			double? value2 = numTextboxMarginBottom.Value;
			float num3 = default(float);
			if (value2.HasValue)
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
				num3 = (float)value2.Value;
				num++;
			}
			double? value3 = numTextboxMarginLeft.Value;
			float num4 = default(float);
			if (value3.HasValue)
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
				num4 = (float)value3.Value;
				num++;
			}
			double? value4 = numTextboxMarginRight.Value;
			float num5 = default(float);
			if (value4.HasValue)
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
				num5 = (float)value4.Value;
				num++;
			}
			if (num != 4)
			{
				if (num != 0)
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
			if (num == 4)
			{
				int numberDecimalDigits = numTextboxMarginTop.NumberDecimalDigits;
				num2 = (float)Math.Round(num2, numberDecimalDigits);
				num3 = (float)Math.Round(num3, numberDecimalDigits);
				num4 = (float)Math.Round(num4, numberDecimalDigits);
				num5 = (float)Math.Round(num5, numberDecimalDigits);
				if (Operators.CompareString(numTextboxMarginTop.CustomUnit, AH.A(69068), TextCompare: false) == 0)
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
					num2 = clsPublish.InchesToPoints(num2);
					num3 = clsPublish.InchesToPoints(num3);
					num4 = clsPublish.InchesToPoints(num4);
					num5 = clsPublish.InchesToPoints(num5);
				}
				else
				{
					num2 = clsPublish.CentimetersToPoints(num2);
					num3 = clsPublish.CentimetersToPoints(num3);
					num4 = clsPublish.CentimetersToPoints(num4);
					num5 = clsPublish.CentimetersToPoints(num5);
				}
				BrandSettings.TextboxMargins = new Settings.Margins
				{
					Top = num2,
					Bottom = num3,
					Left = num4,
					Right = num5
				};
			}
			else
			{
				BrandSettings.TextboxMargins = null;
			}
			BrandSettings.A(this.m_A);
		}
	}

	private MacNumericUpDown[] B()
	{
		return (MacNumericUpDown[])(object)new MacNumericUpDown[4] { numSlideMarginTop, numSlideMarginBottom, numSlideMarginLeft, numSlideMarginRight };
	}

	private void B(Settings A, SpinnerProperties B)
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_0035: Unknown result type (might be due to invalid IL or missing references)
		//IL_0042: Unknown result type (might be due to invalid IL or missing references)
		MacNumericUpDown[] array = this.B();
		foreach (MacNumericUpDown val in array)
		{
			C(val);
			val.CustomUnit = B.CustomUnits;
			val.NumberDecimalDigits = B.Decimals;
			val.SmallChange = B.SmallChange;
			val.LargeChange = B.LargeChange;
			int num;
			if (Operators.CompareString(val.CustomUnit, AH.A(69068), TextCompare: false) != 0)
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
				num = 5;
			}
			else
			{
				num = 2;
			}
			val.Maximum = num;
			_ = null;
		}
		if (A.SlideMargins.HasValue)
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
			Settings.Margins value = A.SlideMargins.Value;
			if (Operators.CompareString(numSlideMarginTop.CustomUnit, AH.A(69068), TextCompare: false) == 0)
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
				numSlideMarginTop.Value = clsPublish.PointsToInches(value.Top);
				numSlideMarginBottom.Value = clsPublish.PointsToInches(value.Bottom);
				numSlideMarginLeft.Value = clsPublish.PointsToInches(value.Left);
				numSlideMarginRight.Value = clsPublish.PointsToInches(value.Right);
			}
			else
			{
				numSlideMarginTop.Value = clsPublish.PointsToCentimeters(value.Top);
				numSlideMarginBottom.Value = clsPublish.PointsToCentimeters(value.Bottom);
				numSlideMarginLeft.Value = clsPublish.PointsToCentimeters(value.Left);
				numSlideMarginRight.Value = clsPublish.PointsToCentimeters(value.Right);
			}
		}
		else
		{
			MacNumericUpDown[] array2 = this.B();
			for (int j = 0; j < array2.Length; j = checked(j + 1))
			{
				array2[j].Value = null;
			}
		}
		MacNumericUpDown[] array3 = this.B();
		foreach (MacNumericUpDown a in array3)
		{
			D(a);
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				break;
			default:
				return;
			}
		}
	}

	private void C(MacNumericUpDown A)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		A.ValueChanged -= new MacRangeBaseValueChangedHandler(SlideMarginsChanged);
	}

	private void D(MacNumericUpDown A)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		A.ValueChanged += new MacRangeBaseValueChangedHandler(SlideMarginsChanged);
	}

	private void SlideMarginsChanged(object sender, MacRangeBaseValueChangedEventArgs e)
	{
		int num = 0;
		double? value = numSlideMarginTop.Value;
		checked
		{
			float num2 = default(float);
			if (value.HasValue)
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
				num2 = (float)value.Value;
				num++;
			}
			double? value2 = numSlideMarginBottom.Value;
			float num3 = default(float);
			if (value2.HasValue)
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
				num3 = (float)value2.Value;
				num++;
			}
			double? value3 = numSlideMarginLeft.Value;
			float num4 = default(float);
			if (value3.HasValue)
			{
				num4 = (float)value3.Value;
				num++;
			}
			double? value4 = numSlideMarginRight.Value;
			float num5 = default(float);
			if (value4.HasValue)
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
				num5 = (float)value4.Value;
				num++;
			}
			if (num != 4)
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
				if (num != 0)
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
			if (num == 4)
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
				int numberDecimalDigits = numSlideMarginTop.NumberDecimalDigits;
				num2 = (float)Math.Round(num2, numberDecimalDigits);
				num3 = (float)Math.Round(num3, numberDecimalDigits);
				num4 = (float)Math.Round(num4, numberDecimalDigits);
				num5 = (float)Math.Round(num5, numberDecimalDigits);
				if (Operators.CompareString(numSlideMarginTop.CustomUnit, AH.A(69068), TextCompare: false) == 0)
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
					num2 = clsPublish.InchesToPoints(num2);
					num3 = clsPublish.InchesToPoints(num3);
					num4 = clsPublish.InchesToPoints(num4);
					num5 = clsPublish.InchesToPoints(num5);
				}
				else
				{
					num2 = clsPublish.CentimetersToPoints(num2);
					num3 = clsPublish.CentimetersToPoints(num3);
					num4 = clsPublish.CentimetersToPoints(num4);
					num5 = clsPublish.CentimetersToPoints(num5);
				}
				BrandSettings.SlideMargins = new Settings.Margins
				{
					Top = num2,
					Bottom = num3,
					Left = num4,
					Right = num5
				};
			}
			else
			{
				BrandSettings.SlideMargins = null;
			}
			BrandSettings.A(this.m_A);
		}
	}

	private void btnValidate_Click(object sender, RoutedEventArgs e)
	{
		ArrayList arrayList = new ArrayList();
		ArrayList arrayList2 = new ArrayList();
		ArrayList arrayList3 = new ArrayList();
		if (Operators.CompareString(Path.GetExtension(this.m_A.Name), AH.A(69996), TextCompare: false) != 0)
		{
			arrayList.Add(AH.A(124005));
		}
		if (this.m_A.Designs.Count > 1)
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
			arrayList.Add(AH.A(124130));
		}
		if (A().Design.Preserved == MsoTriState.msoTrue)
		{
			arrayList.Add(AH.A(124235));
		}
		if (this.m_A.Slides.Count > 0)
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
			arrayList3.Add(AH.A(124314));
		}
		else
		{
			arrayList2.Add(AH.A(124522));
		}
		if (this.m_A.SectionProperties.Count == 0)
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
			arrayList2.Add(AH.A(124756));
		}
		else if (this.m_A.SectionProperties.Count > 1)
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
			arrayList3.Add(AH.A(124994));
		}
		HeadersFooters headersFooters = A().HeadersFooters;
		if (headersFooters.Footer.Visible != MsoTriState.msoTrue)
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
			arrayList2.Add(AH.A(125206));
		}
		if (headersFooters.SlideNumber.Visible != MsoTriState.msoTrue)
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
			arrayList2.Add(AH.A(125315));
		}
		_ = null;
		bool flag = false;
		bool flag2 = false;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A().Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
				Microsoft.Office.Interop.PowerPoint.Shape shape2;
				if (A(shape))
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
					flag = true;
					shape2 = shape;
					string name = shape2.Name;
					if (shape2.Visible != MsoTriState.msoFalse)
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
						arrayList.Add(AH.A(125436) + name + AH.A(125451));
					}
					if (shape2.HasTextFrame == MsoTriState.msoTrue)
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
						if (shape2.TextFrame2.HasText == MsoTriState.msoTrue)
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
							string text = shape2.TextFrame2.TextRange.Text;
							if (text.Contains(PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_STAMP))
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
								if (name.EndsWith(AH.A(125591)))
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
									arrayList2.Add(AH.A(125596) + name + AH.A(125623));
								}
							}
							if (!text.Contains(PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_SECTION) && !text.Contains(PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_SUBSECTION))
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
								if (!text.Contains(PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_SEC_INDEX) && !text.Contains(PowerPointAddIn1.MasterShapes.Placeholders.PLACEHOLDER_SUBSEC_INDEX))
								{
									goto IL_03f3;
								}
							}
							if (name.EndsWith(AH.A(125859)))
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
								arrayList2.Add(AH.A(125596) + name + AH.A(125864));
							}
							else if (!name.EndsWith(AH.A(126126)))
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
								arrayList2.Add(AH.A(126135) + name + AH.A(126164));
							}
						}
					}
					goto IL_03f3;
				}
				if (B(shape))
				{
					flag2 = true;
					Microsoft.Office.Interop.PowerPoint.Shape shape3 = shape;
					if (shape3.Visible != MsoTriState.msoFalse)
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
						arrayList.Add(AH.A(126336) + shape3.Name + AH.A(126363));
					}
					try
					{
						shape3.PickUp();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						arrayList.Add(AH.A(125436) + shape3.Name + AH.A(126436));
						ProjectData.ClearProjectError();
					}
					shape3 = null;
				}
				else
				{
					if (!PowerPointAddIn1.MasterShapes.Base.A(shape))
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
					if (shape.Type != MsoShapeType.msoPlaceholder)
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
					arrayList.Add(AH.A(125436) + shape.Name + AH.A(126513));
				}
				continue;
				IL_03f3:
				shape2 = null;
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
		if (!flag)
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
			arrayList2.Add(AH.A(126735));
		}
		if (!flag2)
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
			arrayList2.Add(AH.A(126933));
		}
		CustomLayout customLayout = null;
		CustomLayout customLayout2 = null;
		CustomLayout customLayout3 = null;
		CustomLayout customLayout4 = null;
		CustomLayout customLayout5 = null;
		int num = 0;
		int num2 = 0;
		int num3 = 0;
		int num4 = 0;
		int num5 = 0;
		int num6 = 0;
		int num7 = 0;
		int num8 = 0;
		checked
		{
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = A().GetEnumerator();
				while (enumerator2.MoveNext())
				{
					CustomLayout customLayout6 = (CustomLayout)enumerator2.Current;
					switch (PowerPointAddIn1.Slides.Helpers.GetLayoutType(customLayout6))
					{
					case SlideType.Title:
						num++;
						break;
					case SlideType.TableOfContents:
					case SlideType.Agenda:
						num2++;
						customLayout = customLayout6;
						break;
					case SlideType.Flysheet:
						num3++;
						customLayout2 = customLayout6;
						break;
					case SlideType.Legal:
						num4++;
						customLayout3 = customLayout6;
						break;
					case SlideType.Contact:
						num5++;
						customLayout4 = customLayout6;
						break;
					case SlideType.Blank:
						num6++;
						customLayout5 = customLayout6;
						break;
					case SlideType.CoverFront:
						num7++;
						break;
					case SlideType.CoverBack:
						num8++;
						break;
					default:
						if (customLayout6.Shapes.HasTitle == MsoTriState.msoFalse)
						{
							arrayList2.Add(AH.A(127131) + customLayout6.Name + AH.A(127148));
						}
						break;
					}
					if (!Regex.IsMatch(customLayout6.Name, AH.A(127215)))
					{
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
					arrayList2.Add(AH.A(127131) + customLayout6.Name + AH.A(127226));
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_06f3;
					}
					continue;
					end_IL_06f3:
					break;
				}
			}
			finally
			{
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							continue;
						}
						(enumerator2 as IDisposable).Dispose();
						break;
					}
				}
			}
			if (num == 0)
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
				arrayList.Add(AH.A(127293));
			}
			else if (num > 1)
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
				arrayList.Add(AH.A(127396));
			}
			if (num2 == 0)
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
				arrayList.Add(AH.A(127513));
			}
			else if (num2 > 1)
			{
				arrayList.Add(AH.A(127610));
			}
			if (num3 == 0)
			{
				arrayList.Add(AH.A(127721));
			}
			else if (num3 > 1)
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
				arrayList.Add(AH.A(128197));
			}
			if (num4 == 0)
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
				arrayList2.Add(AH.A(128324));
			}
			else if (num4 > 1)
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
				arrayList.Add(AH.A(128478));
			}
			if (num5 == 0)
			{
				arrayList2.Add(AH.A(128612));
			}
			else if (num5 > 1)
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
				arrayList.Add(AH.A(128778));
			}
			if (num6 == 0)
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
				arrayList2.Add(AH.A(128924));
			}
			else if (num6 > 1)
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
				arrayList.Add(AH.A(129102));
			}
			if (num7 > 1)
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
				arrayList.Add(AH.A(129250));
			}
			if (num8 > 1)
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
				arrayList.Add(AH.A(129369));
			}
			IEnumerator enumerator3 = default(IEnumerator);
			try
			{
				enumerator3 = A().GetEnumerator();
				IEnumerator enumerator4 = default(IEnumerator);
				while (enumerator3.MoveNext())
				{
					CustomLayout customLayout7 = (CustomLayout)enumerator3.Current;
					try
					{
						enumerator4 = customLayout7.Shapes.GetEnumerator();
						while (enumerator4.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape4 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator4.Current;
							try
							{
								_ = shape4.Tags[PowerPointAddIn1.Shapes.Helpers.TAG_SHAPE_TYPE];
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								ProjectData.ClearProjectError();
							}
							MsoShapeType type = shape4.Type;
							if (type <= MsoShapeType.msoEmbeddedOLEObject)
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
								if (type != MsoShapeType.msoChart)
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
									if (type != MsoShapeType.msoEmbeddedOLEObject)
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
										continue;
									}
								}
							}
							else if (unchecked((uint)(type - 10)) > 2u)
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
								if (type != MsoShapeType.msoMedia)
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
							}
							arrayList2.Add(AH.A(127131) + customLayout7.Name + AH.A(129486) + shape4.Name + AH.A(129523));
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0a23;
							}
							continue;
							end_IL_0a23:
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
			}
			finally
			{
				if (enumerator3 is IDisposable)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						(enumerator3 as IDisposable).Dispose();
						break;
					}
				}
			}
			IEnumerator enumerator5 = default(IEnumerator);
			try
			{
				enumerator5 = A().GetEnumerator();
				IEnumerator enumerator6 = default(IEnumerator);
				while (enumerator5.MoveNext())
				{
					CustomLayout customLayout8 = (CustomLayout)enumerator5.Current;
					try
					{
						enumerator6 = customLayout8.Shapes.GetEnumerator();
						while (enumerator6.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape5 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator6.Current;
							try
							{
								_ = shape5.Tags[PowerPointAddIn1.Shapes.Helpers.TAG_SHAPE_TYPE];
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
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_0afc;
							}
							continue;
							end_IL_0afc:
							break;
						}
					}
					finally
					{
						if (enumerator6 is IDisposable)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								(enumerator6 as IDisposable).Dispose();
								break;
							}
						}
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_0b34;
					}
					continue;
					end_IL_0b34:
					break;
				}
			}
			finally
			{
				if (enumerator5 is IDisposable)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						(enumerator5 as IDisposable).Dispose();
						break;
					}
				}
			}
			if (customLayout != null)
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
				Microsoft.Office.Interop.PowerPoint.Shape shape6 = null;
				Microsoft.Office.Interop.PowerPoint.Shape shape7 = null;
				Microsoft.Office.Interop.PowerPoint.Shape shape8 = null;
				Microsoft.Office.Interop.PowerPoint.Shape shape9 = null;
				Microsoft.Office.Interop.PowerPoint.Shape shape10 = null;
				IEnumerator enumerator7 = default(IEnumerator);
				try
				{
					enumerator7 = customLayout.Shapes.GetEnumerator();
					IEnumerator enumerator8 = default(IEnumerator);
					while (enumerator7.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape11 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator7.Current;
						if (shape11.Type != MsoShapeType.msoGroup)
						{
							continue;
						}
						bool flag3 = false;
						string name2 = shape11.Name;
						if (Operators.CompareString(name2, Constants.AGENDA_TITLE_LEVEL_1, TextCompare: false) == 0)
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
							shape6 = shape11;
							flag3 = true;
						}
						else if (Operators.CompareString(name2, Constants.AGENDA_TITLE_LEVEL_1_ACTIVE, TextCompare: false) == 0)
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
							shape7 = shape11;
							flag3 = true;
						}
						else if (Operators.CompareString(name2, Constants.AGENDA_TITLE_LEVEL_2, TextCompare: false) == 0)
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
							shape8 = shape11;
							flag3 = true;
						}
						else if (Operators.CompareString(name2, Constants.AGENDA_TITLE_LEVEL_2_ACTIVE, TextCompare: false) == 0)
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
							shape9 = shape11;
							flag3 = true;
						}
						else if (Operators.CompareString(name2, Constants.AGENDA_TITLE_LEVEL_3, TextCompare: false) == 0)
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
							shape10 = shape11;
							flag3 = true;
						}
						if (!flag3)
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
						bool flag4 = false;
						bool flag5 = false;
						bool flag6 = false;
						{
							enumerator8 = shape11.GroupItems.GetEnumerator();
							try
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape12;
								for (; enumerator8.MoveNext(); shape12 = null)
								{
									shape12 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator8.Current;
									if (shape12.HasTextFrame != MsoTriState.msoTrue)
									{
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
									if (shape12.TextFrame2.HasText != MsoTriState.msoTrue)
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
									string text2 = shape12.TextFrame2.TextRange.Text.ToUpper();
									uint num9 = YG.A(text2);
									if (num9 <= 211792314)
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
										if (num9 != 7274958)
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
											if (num9 != 149826459)
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
												if (num9 != 211792314)
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
													continue;
												}
												if (Operators.CompareString(text2, AH.A(7299), TextCompare: false) != 0)
												{
													continue;
												}
											}
											else if (Operators.CompareString(text2, AH.A(7342), TextCompare: false) != 0)
											{
												continue;
											}
										}
										else if (Operators.CompareString(text2, AH.A(7312), TextCompare: false) != 0)
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
											continue;
										}
									}
									else
									{
										if (num9 <= 2063731641)
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
											if (num9 != 2050578930)
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
												if (num9 != 2063731641)
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
												}
												else if (Operators.CompareString(text2, AH.A(7365), TextCompare: false) != 0)
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
												}
												else
												{
													flag6 = true;
												}
											}
											else if (Operators.CompareString(text2, AH.A(7277), TextCompare: false) == 0)
											{
												flag4 = true;
											}
											continue;
										}
										if (num9 != 2381896112u)
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
											if (num9 != 3278918548u)
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
												continue;
											}
											if (Operators.CompareString(text2, AH.A(7327), TextCompare: false) != 0)
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
												continue;
											}
										}
										else if (Operators.CompareString(text2, AH.A(7284), TextCompare: false) != 0)
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
											continue;
										}
									}
									flag5 = true;
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_0ee3;
									}
									continue;
									end_IL_0ee3:
									break;
								}
							}
							finally
							{
								IDisposable disposable = enumerator8 as IDisposable;
								if (disposable != null)
								{
									disposable.Dispose();
								}
							}
						}
						if (!flag5)
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
							arrayList.Add(AH.A(129604) + shape11.Name + AH.A(129649));
						}
						if (!flag6)
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
							arrayList2.Add(AH.A(129604) + shape11.Name + AH.A(129746));
						}
						if (!flag4 && Operators.CompareString(shape11.Name, Constants.AGENDA_TITLE_LEVEL_3, TextCompare: false) != 0)
						{
							arrayList.Add(AH.A(129604) + shape11.Name + AH.A(129813));
						}
						if (shape11.Visible != MsoTriState.msoTrue)
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
						arrayList.Add(AH.A(129604) + shape11.Name + AH.A(129912));
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							goto end_IL_1021;
						}
						continue;
						end_IL_1021:
						break;
					}
				}
				finally
				{
					if (enumerator7 is IDisposable)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							(enumerator7 as IDisposable).Dispose();
							break;
						}
					}
				}
				if (shape6 == null)
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
					arrayList.Add(AH.A(129604) + Constants.AGENDA_TITLE_LEVEL_1 + AH.A(130094));
				}
				if (shape7 == null)
				{
					arrayList2.Add(AH.A(129604) + Constants.AGENDA_TITLE_LEVEL_1_ACTIVE + AH.A(130169));
				}
				if (shape8 == null)
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
					arrayList.Add(AH.A(129604) + Constants.AGENDA_TITLE_LEVEL_2 + AH.A(130094));
				}
				if (shape9 == null)
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
					arrayList2.Add(AH.A(129604) + Constants.AGENDA_TITLE_LEVEL_2_ACTIVE + AH.A(130447));
				}
				if (shape10 == null)
				{
					arrayList2.Add(AH.A(129604) + Constants.AGENDA_TITLE_LEVEL_3 + AH.A(130727));
				}
			}
			if (customLayout2 != null)
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
				int num10 = 0;
				{
					IEnumerator enumerator9 = customLayout2.Shapes.Placeholders.GetEnumerator();
					try
					{
						while (enumerator9.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape13 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator9.Current;
							if (shape13.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody && shape13.HasTextFrame == MsoTriState.msoTrue)
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
								TextRange2 textRange = shape13.TextFrame2.TextRange;
								if (textRange.get_Paragraphs(-1, -1).Count == 2)
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
									if (textRange.get_Paragraphs(1, 1).ParagraphFormat.Bullet.Type == MsoBulletType.msoBulletNumbered)
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
										if (textRange.get_Paragraphs(2, 1).ParagraphFormat.Bullet.Type == MsoBulletType.msoBulletNumbered)
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
											if (textRange.get_Paragraphs(1, 1).ParagraphFormat.IndentLevel == 1 && textRange.get_Paragraphs(2, 1).ParagraphFormat.IndentLevel == 2)
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
												num10++;
											}
										}
									}
								}
								textRange = null;
							}
							shape13 = null;
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_12a5;
							}
							continue;
							end_IL_12a5:
							break;
						}
					}
					finally
					{
						IDisposable disposable2 = enumerator9 as IDisposable;
						if (disposable2 != null)
						{
							disposable2.Dispose();
						}
					}
				}
				if (num10 != 1)
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
					arrayList.Add(AH.A(131145));
				}
				if (customLayout2.Shapes.HasTitle == MsoTriState.msoTrue)
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
					arrayList2.Add(AH.A(131771));
				}
			}
			if (customLayout3 != null)
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
				bool flag7 = false;
				bool flag8 = false;
				{
					IEnumerator enumerator10 = customLayout3.Shapes.GetEnumerator();
					try
					{
						while (enumerator10.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape14 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator10.Current;
							if (shape14.HasTextFrame != MsoTriState.msoTrue)
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
							if (shape14.TextFrame.HasText != MsoTriState.msoTrue)
							{
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
							if (shape14.TextFrame.TextRange.Characters().Count <= 50)
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
							flag7 = true;
							if (shape14.Type != MsoShapeType.msoTextBox)
							{
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
							flag8 = true;
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_13de;
							}
							continue;
							end_IL_13de:
							break;
						}
					}
					finally
					{
						IDisposable disposable3 = enumerator10 as IDisposable;
						if (disposable3 != null)
						{
							disposable3.Dispose();
						}
					}
				}
				if (!flag7)
				{
					arrayList2.Add(AH.A(131921));
				}
				if (!flag8)
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
					arrayList2.Add(AH.A(132347));
				}
			}
			if (customLayout4 != null)
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
				bool flag9 = false;
				IEnumerator enumerator11 = default(IEnumerator);
				try
				{
					enumerator11 = customLayout4.Shapes.GetEnumerator();
					while (enumerator11.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape15 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator11.Current;
						if (shape15.HasTextFrame != MsoTriState.msoTrue || shape15.TextFrame.HasText != MsoTriState.msoTrue || shape15.TextFrame.TextRange.Characters().Count <= 10)
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
						flag9 = true;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_14cd;
						}
						continue;
						end_IL_14cd:
						break;
					}
				}
				finally
				{
					if (enumerator11 is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator11 as IDisposable).Dispose();
							break;
						}
					}
				}
				if (!flag9)
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
					arrayList2.Add(AH.A(132689));
				}
			}
			if (customLayout5 != null)
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
				IEnumerator enumerator12 = default(IEnumerator);
				try
				{
					enumerator12 = customLayout5.Shapes.GetEnumerator();
					while (true)
					{
						if (enumerator12.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape16 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator12.Current;
							if (shape16.Type != MsoShapeType.msoPlaceholder)
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
							if (shape16.PlaceholderFormat.Type != PpPlaceholderType.ppPlaceholderSlideNumber)
							{
								continue;
							}
							while (true)
							{
								switch (3)
								{
								case 0:
									continue;
								}
								arrayList.Add(AH.A(133155));
								break;
							}
							break;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_159e;
							}
							continue;
							end_IL_159e:
							break;
						}
						break;
					}
				}
				finally
				{
					if (enumerator12 is IDisposable)
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							(enumerator12 as IDisposable).Dispose();
							break;
						}
					}
				}
				if (customLayout5.Shapes.HasTitle == MsoTriState.msoTrue)
				{
					arrayList.Add(AH.A(133317));
				}
			}
			List<ValidationMessage> list = new List<ValidationMessage>();
			List<ValidationMessage> list2 = list;
			IEnumerator enumerator13 = arrayList.GetEnumerator();
			try
			{
				while (enumerator13.MoveNext())
				{
					string strMessage = Conversions.ToString(enumerator13.Current);
					list2.Add(new ValidationMessage(strMessage, ValidationType.ErrorLevel));
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_1633;
					}
					continue;
					end_IL_1633:
					break;
				}
			}
			finally
			{
				IDisposable disposable4 = enumerator13 as IDisposable;
				if (disposable4 != null)
				{
					disposable4.Dispose();
				}
			}
			IEnumerator enumerator14 = default(IEnumerator);
			try
			{
				enumerator14 = arrayList2.GetEnumerator();
				while (enumerator14.MoveNext())
				{
					string strMessage2 = Conversions.ToString(enumerator14.Current);
					list2.Add(new ValidationMessage(strMessage2, ValidationType.WarningLevel));
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						goto end_IL_168e;
					}
					continue;
					end_IL_168e:
					break;
				}
			}
			finally
			{
				if (enumerator14 is IDisposable)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						(enumerator14 as IDisposable).Dispose();
						break;
					}
				}
			}
			if (!list2.Any())
			{
				list2.Add(new ValidationMessage(AH.A(133477), ValidationType.SuccessLevel));
			}
			foreach (object item in arrayList3)
			{
				string strMessage3 = Conversions.ToString(item);
				list2.Add(new ValidationMessage(strMessage3, ValidationType.InfoLevel));
			}
			list2 = null;
			lbxValidate.ItemsSource = list;
			lbxValidate.Visibility = Visibility.Visible;
			customLayout = null;
			customLayout2 = null;
			customLayout4 = null;
			customLayout3 = null;
			customLayout5 = null;
			list = null;
		}
	}

	private bool A(Microsoft.Office.Interop.PowerPoint.Shape A, string B)
	{
		string left = "";
		try
		{
			left = A.Tags[PowerPointAddIn1.Shapes.Helpers.TAG_SHAPE_TYPE];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return Operators.CompareString(left, B, TextCompare: false) == 0;
	}

	private void TemplateDocs(object sender, MouseButtonEventArgs e)
	{
		B(AH.A(133611));
	}

	private void MasterShapesDocs(object sender, MouseButtonEventArgs e)
	{
		B(AH.A(133722));
	}

	private void StylesDocs(object sender, MouseButtonEventArgs e)
	{
		B(AH.A(133815));
	}

	private void B(string A)
	{
		clsSupport.OnlineDocs(A);
	}

	private void H()
	{
		A(PpViewType.ppViewSlide, AH.A(133896));
	}

	private void I()
	{
		A(PpViewType.ppViewSlideMaster, AH.A(133927));
	}

	private void A(PpViewType A, string B)
	{
		try
		{
			this.m_A.Windows[1].Activate();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		try
		{
			if (this.m_A.ActiveWindow.Panes[2].ViewType == A)
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
				this.m_A.CommandBars.ExecuteMso(B);
				System.Windows.Forms.Application.DoEvents();
				return;
			}
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			if (A == PpViewType.ppViewSlideMaster)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						try
						{
							this.m_A.ActiveWindow.ViewType = PpViewType.ppViewMasterThumbnails;
						}
						catch (Exception ex5)
						{
							ProjectData.SetProjectError(ex5);
							Exception ex6 = ex5;
							throw;
						}
						ProjectData.ClearProjectError();
						return;
					}
				}
			}
			throw;
		}
	}

	private void A(Selection A)
	{
		if (this.m_A.ActivePresentation != this.m_A)
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
			if (this.A().Count == lvLayouts.Items.Count)
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
				B();
				return;
			}
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		if (Operators.CompareString(A.FullName, this.m_A.FullName, TextCompare: false) != 0)
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
			Close();
			return;
		}
	}

	private void C(string A)
	{
		Forms.ErrorMessage(Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.WarningMessage(Window.GetWindow(this), A);
	}

	private void E(string A)
	{
		Forms.InfoMessage(Window.GetWindow(this), A);
	}

	private Master A()
	{
		return this.m_A.Designs[1].SlideMaster;
	}

	private CustomLayouts A()
	{
		return A().CustomLayouts;
	}

	private CustomLayout A(int A)
	{
		return this.A()[A];
	}

	private Microsoft.Office.Interop.PowerPoint.Presentation A()
	{
		return this.m_A.Presentations.Open(Path.Combine(clsEnvironment.CommonAppDataPath, AH.A(133966), AH.A(133995), AH.A(134010), AH.A(134023)), MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);
	}

	private void A(GridView A)
	{
		Forms.AutoResizeGridView(A);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_B)
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
			this.m_B = true;
			Uri resourceLocator = new Uri(AH.A(134046), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
			return;
		}
	}

	void IComponentConnector.InitializeComponent()
	{
		//ILSpy generated this explicit interface implementation from .override directive in InitializeComponent
		this.InitializeComponent();
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		//IL_0329: Unknown result type (might be due to invalid IL or missing references)
		//IL_0333: Expected O, but got Unknown
		//IL_0345: Unknown result type (might be due to invalid IL or missing references)
		//IL_034f: Expected O, but got Unknown
		//IL_0361: Unknown result type (might be due to invalid IL or missing references)
		//IL_036b: Expected O, but got Unknown
		//IL_038f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0399: Expected O, but got Unknown
		//IL_037d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0387: Expected O, but got Unknown
		//IL_03bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_03c7: Expected O, but got Unknown
		//IL_03ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_03b5: Expected O, but got Unknown
		//IL_03d9: Unknown result type (might be due to invalid IL or missing references)
		//IL_03e3: Expected O, but got Unknown
		//IL_03f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_03ff: Expected O, but got Unknown
		//IL_0411: Unknown result type (might be due to invalid IL or missing references)
		//IL_041b: Expected O, but got Unknown
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
					TabControl1 = (System.Windows.Controls.TabControl)target;
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
					tabInst = (TabItem)target;
					return;
				}
			}
		}
		if (connectionId == 3)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					((TextBlock)target).MouseLeftButtonUp += TemplateDocs;
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
					tabLayouts = (TabItem)target;
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
					lvLayouts = (System.Windows.Controls.ListView)target;
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
					gvLayouts = (GridView)target;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					tabAgenda = (TabItem)target;
					return;
				}
			}
		}
		if (connectionId == 10)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnAgendaGroupsBuild = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			btnAgendaGroupsShow = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 12)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnAgendaGroupsHide = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					((System.Windows.Controls.Button)target).Click += BuildSectDivPlaceholder;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					((System.Windows.Controls.Button)target).Click += SelectSectDivPlaceholder;
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
					tabShapes = (TabItem)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			lvMasterShapes = (System.Windows.Controls.ListView)target;
			return;
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					gvMasterShapes = (GridView)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnMasterShapesShow = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnMasterShapesHide = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					((TextBlock)target).MouseLeftButtonUp += MasterShapesDocs;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					lvStyles = (System.Windows.Controls.ListView)target;
					return;
				}
			}
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					gvStyles = (GridView)target;
					return;
				}
			}
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnStylesShow = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnStylesHide = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			((TextBlock)target).MouseLeftButtonUp += StylesDocs;
			return;
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					tabBrand = (TabItem)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			lbxFonts = (System.Windows.Controls.ListBox)target;
			return;
		}
		if (connectionId == 28)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnFontTypeScan = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 29)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnFontTypeAdd = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnFontTypeDelete = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 31)
		{
			numMinFontSize = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 32)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					numMaxFontSize = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 33)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					numTextboxMarginTop = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 34)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					numTextboxMarginBottom = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 35)
		{
			numTextboxMarginLeft = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 36)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					numTextboxMarginRight = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 37)
		{
			numSlideMarginTop = (MacNumericUpDown)target;
			return;
		}
		if (connectionId == 38)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					numSlideMarginBottom = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 39)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					numSlideMarginLeft = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 40)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					numSlideMarginRight = (MacNumericUpDown)target;
					return;
				}
			}
		}
		if (connectionId == 41)
		{
			tabValidate = (TabItem)target;
			return;
		}
		if (connectionId == 42)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnValidate = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 43)
		{
			lbxValidate = (System.Windows.Controls.ListBox)target;
			return;
		}
		if (connectionId == 44)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnSelectionPane = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 45)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnClose = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 46)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					btnRefresh = (System.Windows.Controls.Button)target;
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

	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 7)
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
			((System.Windows.Controls.Button)target).Click += RenameLayout;
		}
		if (connectionId != 8)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += LayoutRoleChanged;
			((System.Windows.Controls.ComboBox)target).DropDownOpened += OnDropDownOpened;
			return;
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}
}
