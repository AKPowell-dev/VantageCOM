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
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Libraries;
using MacabacusMacros.Libraries.Manage.Publish;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.Libraries.Pane.Filters;
using MacabacusMacros.Libraries.Pane.UI;
using MacabacusMacros.Libraries.Tags;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.UI;
using MacabacusMacros.Xaml;
using Macabacus_Word.Library2.Versioning;
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Library2.UI;

[DesignerGenerated]
public sealed class wpfLibrary : System.Windows.Controls.UserControl, INotifyPropertyChanged, IUIWithPreviewWin, IComponentConnector, IStyleConnector
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<ContentGroup, FiltersGroup> A;

		public static Func<FiltersGroup, bool> A;

		public static Func<ContentItem, bool> A;

		public static Func<ContentGroup, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal FiltersGroup A(ContentGroup A)
		{
			return A.FiltersGroup;
		}

		[SpecialName]
		internal bool A(FiltersGroup A)
		{
			return A != null;
		}

		[SpecialName]
		internal bool A(ContentItem A)
		{
			return A.Visibility != Visibility.Visible;
		}

		[SpecialName]
		internal int A(ContentGroup A)
		{
			if (A.AllContentItems != null)
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
						return ((Collection<ContentItem>)(object)A.AllContentItems).Count;
					}
				}
			}
			return 0;
		}
	}

	[CompilerGenerated]
	internal sealed class W
	{
		public int A;

		[SpecialName]
		internal void A()
		{
			this.A = 1;
		}
	}

	[CompilerGenerated]
	internal sealed class X
	{
		public Microsoft.Office.Interop.Excel.Chart A;

		[SpecialName]
		internal void A()
		{
			this.A.ChartArea.Copy();
		}
	}

	[CompilerGenerated]
	internal sealed class Y
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public Action A;

		public Y(Y A)
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
			this.A.Copy();
		}
	}

	[CompilerGenerated]
	internal sealed class Z
	{
		public ContentType A;

		[SpecialName]
		internal bool A(ContentGroup A)
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0007: Unknown result type (might be due to invalid IL or missing references)
			return A.ContentType == this.A;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	public List<string> HiddenFiles;

	public Microsoft.Office.Interop.PowerPoint.Application ppApp;

	public bool blnQuitPP;

	public Microsoft.Office.Interop.Excel.Application xlApp;

	public bool blnQuitXL;

	private bool m_A;

	private PreviewWinHandler m_A;

	private bool m_B;

	private bool m_C;

	private XmlDocument m_A;

	private bool m_D;

	[CompilerGenerated]
	private bool m_E;

	private ICollectionView m_A;

	[CompilerGenerated]
	private ContentGroupsCollection m_A;

	[CompilerGenerated]
	private wpfDragDrop m_A;

	private TagEditor m_A;

	private ObservableCollection<FiltersGroup> m_A;

	[CompilerGenerated]
	private ItemsPanelTemplate m_A;

	[CompilerGenerated]
	private ItemsPanelTemplate m_B;

	[CompilerGenerated]
	private System.Windows.Controls.ListView m_A;

	[CompilerGenerated]
	private bool m_F;

	private Search m_A;

	private ContentGroupsListener m_A;

	private string m_A;

	private string m_B;

	private bool m_G;

	private bool m_H;

	private bool m_I;

	private bool m_J;

	private object m_A;

	private System.Windows.Point m_A;

	private ObservableCollection<ContentItem> m_A;

	private int m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("popSuggest")]
	private Popup m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("rtbSearch")]
	private System.Windows.Controls.RichTextBox m_A;

	[AccessedThroughProperty("chkShared")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPersonal")]
	private System.Windows.Controls.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chk3rdParty")]
	private System.Windows.Controls.CheckBox m_C;

	[AccessedThroughProperty("chkPublic")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_D;

	[AccessedThroughProperty("scroller")]
	[CompilerGenerated]
	private ScrollViewer m_A;

	[AccessedThroughProperty("chkFilters")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("stkFilters")]
	private StackPanel m_A;

	[AccessedThroughProperty("chkContentFilters")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("grdContentFilters")]
	private DockPanel m_A;

	[AccessedThroughProperty("chkCharts")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_G;

	[AccessedThroughProperty("chkShapes")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_H;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImages")]
	private System.Windows.Controls.CheckBox m_I;

	[CompilerGenerated]
	[AccessedThroughProperty("chkText")]
	private System.Windows.Controls.CheckBox m_J;

	[AccessedThroughProperty("chkPdfs")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_K;

	[AccessedThroughProperty("chkDocs")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_L;

	[AccessedThroughProperty("chkFavorites")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_M;

	[AccessedThroughProperty("chkImageTypeFilters")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_N;

	[AccessedThroughProperty("grdImageTypeFilters")]
	[CompilerGenerated]
	private StackPanel m_B;

	[AccessedThroughProperty("chkImagesSvg")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_O;

	[AccessedThroughProperty("chkImagesPng")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_P;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesJpg")]
	private System.Windows.Controls.CheckBox m_Q;

	[AccessedThroughProperty("chkImagesEmf")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox R;

	[AccessedThroughProperty("chkImagesWmf")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox S;

	[AccessedThroughProperty("chkImagesGif")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox T;

	[AccessedThroughProperty("chkImagesBmp")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox U;

	[AccessedThroughProperty("chkImagesTiff")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox V;

	[CompilerGenerated]
	[AccessedThroughProperty("icFilters")]
	private ItemsControl m_A;

	[AccessedThroughProperty("icContent")]
	[CompilerGenerated]
	private ItemsControl m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPreview")]
	private System.Windows.Controls.CheckBox m_W;

	[CompilerGenerated]
	[AccessedThroughProperty("chkStars")]
	private System.Windows.Controls.CheckBox m_X;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImageTypeBadge")]
	private System.Windows.Controls.CheckBox m_Y;

	[AccessedThroughProperty("btnInsert")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkKeep")]
	private System.Windows.Controls.CheckBox m_Z;

	[AccessedThroughProperty("popRemoveGroup")]
	[CompilerGenerated]
	private Popup m_B;

	[AccessedThroughProperty("btnHideContent")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	private bool m_K;

	internal bool HiddenOnBackstaging
	{
		[CompilerGenerated]
		get
		{
			return this.m_E;
		}
		[CompilerGenerated]
		set
		{
			this.m_E = value;
		}
	}

	public ICollectionView SourceCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(XC.A(3391));
		}
	}

	internal ContentGroupsCollection AllGroups
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

	public wpfDragDrop DragDropOverlay
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

	public TagEditor TagEditor
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(XC.A(9841));
		}
	}

	public ObservableCollection<FiltersGroup> ContentFilters
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(XC.A(9860));
		}
	}

	private ItemsPanelTemplate ItemsPanelWrap
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

	private ItemsPanelTemplate ItemsPanelStack
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

	private System.Windows.Controls.ListView ActiveListView
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

	private bool IsAdmin
	{
		[CompilerGenerated]
		get
		{
			return this.m_F;
		}
		[CompilerGenerated]
		set
		{
			this.m_F = value;
		}
	}

	public ContentGroupsListener GroupsListener
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(XC.A(9889));
		}
	}

	public string SelectedCountStr
	{
		get
		{
			return this.m_A;
		}
		set
		{
			if (object.Equals(this.m_A, value))
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
				this.m_A = value;
				A(XC.A(9918));
				return;
			}
		}
	}

	public string SelActionText
	{
		get
		{
			return this.m_B;
		}
		set
		{
			if (!object.Equals(this.m_B, value))
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
				this.m_B = value;
				A(XC.A(9958));
				return;
			}
		}
	}

	private FrameworkElement PreviewParentUIElem => this;

	private bool PreviewSetting
	{
		get
		{
			return global::A.N.Settings.LibraryPaneShowPreview;
		}
		set
		{
			global::A.N.Settings.LibraryPaneShowPreview = value;
		}
	}

	public bool IsFilteringContent
	{
		get
		{
			return this.m_I;
		}
		set
		{
			this.m_I = value;
			A(XC.A(10246));
		}
	}

	public bool IsRunningAsyncLoad
	{
		get
		{
			return this.m_J;
		}
		set
		{
			this.m_J = value;
			A(XC.A(10283));
		}
	}

	internal virtual Popup popSuggest
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

	internal virtual System.Windows.Controls.RichTextBox rtbSearch
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

	internal virtual System.Windows.Controls.CheckBox chkShared
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

	internal virtual System.Windows.Controls.CheckBox chkPersonal
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

	internal virtual System.Windows.Controls.CheckBox chk3rdParty
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

	internal virtual System.Windows.Controls.CheckBox chkPublic
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

	internal virtual System.Windows.Controls.CheckBox chkFilters
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

	internal virtual StackPanel stkFilters
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

	internal virtual System.Windows.Controls.CheckBox chkContentFilters
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
			RoutedEventHandler value2 = ContentFiltersChecked;
			RoutedEventHandler value3 = ContentFiltersUnchecked;
			System.Windows.Controls.CheckBox checkBox = this.m_F;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_F = value;
			checkBox = this.m_F;
			if (checkBox != null)
			{
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
			}
		}
	}

	internal virtual DockPanel grdContentFilters
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

	internal virtual System.Windows.Controls.CheckBox chkCharts
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

	internal virtual System.Windows.Controls.CheckBox chkShapes
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

	internal virtual System.Windows.Controls.CheckBox chkImages
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

	internal virtual System.Windows.Controls.CheckBox chkText
	{
		[CompilerGenerated]
		get
		{
			return this.m_J;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_J = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkPdfs
	{
		[CompilerGenerated]
		get
		{
			return this.m_K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_K = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkDocs
	{
		[CompilerGenerated]
		get
		{
			return this.m_L;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_L = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFavorites
	{
		[CompilerGenerated]
		get
		{
			return this.m_M;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_M = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImageTypeFilters
	{
		[CompilerGenerated]
		get
		{
			return this.m_N;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = ImageTypeFiltersChecked;
			RoutedEventHandler value3 = ImageTypeFiltersUnchecked;
			System.Windows.Controls.CheckBox checkBox = this.m_N;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_N = value;
			checkBox = this.m_N;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual StackPanel grdImageTypeFilters
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

	internal virtual System.Windows.Controls.CheckBox chkImagesSvg
	{
		[CompilerGenerated]
		get
		{
			return this.m_O;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_O = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesPng
	{
		[CompilerGenerated]
		get
		{
			return this.m_P;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_P = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesJpg
	{
		[CompilerGenerated]
		get
		{
			return this.m_Q;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_Q = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesEmf
	{
		[CompilerGenerated]
		get
		{
			return R;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			R = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesWmf
	{
		[CompilerGenerated]
		get
		{
			return S;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			S = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesGif
	{
		[CompilerGenerated]
		get
		{
			return T;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			T = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesBmp
	{
		[CompilerGenerated]
		get
		{
			return U;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			U = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesTiff
	{
		[CompilerGenerated]
		get
		{
			return V;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			V = value;
		}
	}

	internal virtual ItemsControl icFilters
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

	internal virtual ItemsControl icContent
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

	internal virtual System.Windows.Controls.CheckBox chkPreview
	{
		[CompilerGenerated]
		get
		{
			return this.m_W;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_W = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkStars
	{
		[CompilerGenerated]
		get
		{
			return this.m_X;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_X = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImageTypeBadge
	{
		[CompilerGenerated]
		get
		{
			return this.m_Y;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_Y = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnInsert
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
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkKeep
	{
		[CompilerGenerated]
		get
		{
			return this.m_Z;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_Z = value;
		}
	}

	internal virtual Popup popRemoveGroup
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

	internal virtual System.Windows.Controls.Button btnHideContent
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

	public wpfLibrary()
	{
		//IL_00b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ba: Expected O, but got Unknown
		//IL_00e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fc: Expected O, but got Unknown
		//IL_00f7: Unknown result type (might be due to invalid IL or missing references)
		//IL_0101: Expected O, but got Unknown
		//IL_0131: Unknown result type (might be due to invalid IL or missing references)
		//IL_0142: Expected O, but got Unknown
		//IL_0149: Unknown result type (might be due to invalid IL or missing references)
		//IL_0153: Expected O, but got Unknown
		base.Loaded += wpfLibrary_Loaded;
		base.Unloaded += wpfLibrary_Unloaded;
		blnQuitPP = false;
		blnQuitXL = false;
		this.m_B = false;
		this.m_A = null;
		this.m_A = "";
		this.m_B = XC.A(10224);
		this.m_G = false;
		this.m_H = false;
		this.m_I = false;
		this.m_J = false;
		this.m_A = RuntimeHelpers.GetObjectValue(new object());
		this.m_A = new ObservableCollection<ContentItem>();
		this.m_A = -1;
		InitializeComponent();
		ContextMenus.FixMenuAlignment();
		this.m_A = new PreviewWinHandler((IUIWithPreviewWin)(object)this);
		GroupsListener = new ContentGroupsListener((Action)F, (Action)D, (Action<ContentType, bool>)A, new ShowsLibrariesChangedDelegate(E), (Action<ImageType?, bool>)A);
		TagEditor tagEditor = TagEditor;
		Search val = new Search(ref tagEditor, GroupsListener, rtbSearch, popSuggest, (Func<object, object>)base.FindResource);
		TagEditor = tagEditor;
		this.m_A = val;
		ContentGroup.GroupFiltered += new GroupFilteredEventHandler(A);
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private void A()
	{
		System.Windows.Controls.ListView activeListView = ActiveListView;
		int? obj;
		if (activeListView != null)
		{
			obj = activeListView.SelectedItems?.Count;
		}
		else
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
			obj = null;
		}
		int? num = obj;
		int valueOrDefault = num.GetValueOrDefault();
		SelectedCountStr = ((valueOrDefault < 2) ? "" : string.Format(XC.A(9951), valueOrDefault));
	}

	private void wpfLibrary_Loaded(object sender, RoutedEventArgs e)
	{
		this.m_A.OnLoad();
		chkPreview.IsChecked = global::A.N.Settings.LibraryPaneShowPreview;
		chkPreview.Checked += PreviewToggle;
		chkPreview.Unchecked += PreviewToggle;
		chkStars.IsChecked = global::A.N.Settings.LibraryPaneShowStars;
		chkStars.Checked += StarsToggle;
		chkStars.Unchecked += StarsToggle;
		chkImageTypeBadge.IsChecked = global::A.N.Settings.LibraryPaneShowImageTypeBadge;
		chkImageTypeBadge.Checked += ImageTypeBadgeToggle;
		chkImageTypeBadge.Unchecked += ImageTypeBadgeToggle;
		chkKeep.IsChecked = global::A.N.Settings.LibraryPaneKeepSourceFormat;
		chkKeep.Checked += KeepSourceFormatToggle;
		chkKeep.Unchecked += KeepSourceFormatToggle;
	}

	private void wpfLibrary_Unloaded(object sender, RoutedEventArgs e)
	{
		//IL_0052: Unknown result type (might be due to invalid IL or missing references)
		//IL_005c: Expected O, but got Unknown
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		//IL_006d: Expected O, but got Unknown
		this.m_G = true;
		CleanUp();
		Pane.A(this);
		if (AllGroups != null)
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
			IEnumerator<ContentGroup> enumerator = default(IEnumerator<ContentGroup>);
			try
			{
				enumerator = ((Collection<ContentGroup>)(object)AllGroups).GetEnumerator();
				while (enumerator.MoveNext())
				{
					ContentGroup current = enumerator.Current;
					current.LoadedItemsInfoChanged -= new LoadedItemsInfoChangedEventHandler(this.B);
					current.HasFavoritesChanged -= new HasFavoritesChangedEventHandler(A);
				}
			}
			finally
			{
				if (enumerator != null)
				{
					while (true)
					{
						switch (2)
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
		if (GroupsListener != null)
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
			GroupsListener.Disconnect();
			GroupsListener = null;
		}
		Search.Release(ref this.m_A);
		HiddenFiles = null;
		PreviewWinHandler.ClearReferences(ref this.m_A);
		DragDropOverlay = null;
	}

	private void A(ContentGroupsCollection A)
	{
		MySettings settings = global::A.N.Settings;
		A.ShowsPersonalLibrary = settings.ContentInsertShowPersonal;
		A.ShowsSharedLibraries = settings.ContentInsertShowShared;
		A.ShowsPublicLibrary = settings.ContentInsertShowPublic;
		A.Shows3rdPartyLibraries = settings.ContentInsertShow3rdParty;
		A[(ContentType)3] = settings.ContentInsertShowShapes;
		A[(ContentType)4] = settings.ContentInsertShowImages;
		A[(ContentType)5] = settings.ContentInsertShowCharts;
		A[(ContentType)7] = settings.ContentInsertShowText;
		A[(ContentType)15] = settings.ContentInsertShowPDFs && Pane.PDFsContentIsEnabled;
		int num;
		if (settings.ContentInsertShowDocs)
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
			num = (Pane.DocsContentIsEnabled ? 1 : 0);
		}
		else
		{
			num = 0;
		}
		A[(ContentType)14] = (byte)num != 0;
		A.ImgTypesFilter.FromListString(settings.ContentInsertExcludeImageTypes);
		settings = null;
		G();
	}

	public void LoadContent()
	{
		if (this.m_G)
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
			this.m_A = NC.A.SettingsXml;
			this.m_D = global::A.N.Settings.LibraryPaneShowStars;
			HiddenFiles = new List<string>();
			IsAdmin = Base.IsUserAdmin();
			ItemsPanelWrap = (ItemsPanelTemplate)FindResource(XC.A(9985));
			ItemsPanelStack = (ItemsPanelTemplate)FindResource(XC.A(10014));
			AllGroups = Load.GetLibraryContent((Func<XmlDocument, string, LibraryItem, bool, ContentGroup>)A, (Action<ContentGroupsCollection>)A);
			B();
			GroupsListener.Groups = AllGroups;
			SourceCollection = CollectionViewSource.GetDefaultView(AllGroups);
			ICollectionView sourceCollection = SourceCollection;
			if (!sourceCollection.GroupDescriptions.Any())
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
				sourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(XC.A(8561)));
			}
			sourceCollection.Filter = A;
			sourceCollection = null;
			chkContentFilters.IsChecked = true;
			chkImageTypeFilters.IsChecked = true;
			this.m_A = null;
			return;
		}
	}

	private ContentGroup A(XmlDocument A, string B, LibraryItem C, bool D)
	{
		//IL_01d6: Unknown result type (might be due to invalid IL or missing references)
		//IL_01da: Invalid comparison between Unknown and I4
		//IL_0042: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Invalid comparison between Unknown and I4
		//IL_008a: Unknown result type (might be due to invalid IL or missing references)
		//IL_008d: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a7: Expected I4, but got Unknown
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ab: Invalid comparison between Unknown and I4
		//IL_012d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0138: Unknown result type (might be due to invalid IL or missing references)
		//IL_013b: Invalid comparison between Unknown and I4
		//IL_00b7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00bb: Invalid comparison between Unknown and I4
		//IL_015d: Unknown result type (might be due to invalid IL or missing references)
		//IL_017a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0181: Expected O, but got Unknown
		//IL_018a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0194: Expected O, but got Unknown
		//IL_019d: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a7: Expected O, but got Unknown
		ContentGroup val = null;
		XmlNodeList childNodes = A.DocumentElement.ChildNodes;
		if (childNodes != null && childNodes.Count > 0)
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
			string directoryName = Path.GetDirectoryName(B);
			ContentType contentType = Manifests.GetContentType(A);
			Base.ProcessManifestNodes(ref childNodes, directoryName, contentType);
			if ((int)contentType == 14)
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
				if (!Access.IsEnterprisePlanOrTrialMode() && !Access.IsLegacyPlan())
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							childNodes = null;
							return null;
						}
					}
				}
			}
			string a;
			switch (contentType - 3)
			{
			default:
				if ((int)contentType != 14)
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
					if ((int)contentType != 15)
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
						goto case 3;
					}
					a = XC.A(10119);
					break;
				}
				a = XC.A(10134);
				break;
			case 0:
				a = XC.A(10045);
				break;
			case 1:
				a = XC.A(10064);
				break;
			case 2:
				a = XC.A(10083);
				break;
			case 4:
				a = XC.A(10102);
				break;
			case 3:
				childNodes = null;
				return null;
			}
			bool flag = Favorites.ManifestContainsFavorite(A, contentType);
			ItemsPanelTemplate itemsPanelTemplate;
			if ((int)contentType != 7)
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
				itemsPanelTemplate = ItemsPanelWrap;
			}
			else
			{
				itemsPanelTemplate = ItemsPanelStack;
			}
			try
			{
				val = new ContentGroup(C, contentType, directoryName, A, D, IsAdmin, this.A(a), itemsPanelTemplate, this.m_D, flag);
				val.HasFavoritesChanged += new HasFavoritesChangedEventHandler(this.A);
				val.LoadedItemsInfoChanged += new LoadedItemsInfoChangedEventHandler(this.B);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				Forms.ErrorMessage(XC.A(10159) + directoryName);
				ProjectData.ClearProjectError();
			}
			itemsPanelTemplate = null;
			if (val != null && (int)contentType == 14)
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
				XmlNode schemaNodeById = CustomFields.GetSchemaNodeById(this.m_A, val.MetadataSchemaId);
				if (schemaNodeById != null)
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
					if (schemaNodeById.ChildNodes.Count > 0)
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
						this.A(val, schemaNodeById);
					}
				}
				schemaNodeById = null;
			}
		}
		childNodes = null;
		return val;
	}

	private void B()
	{
		this.m_C = true;
		try
		{
			if (ContentFilters == null)
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
					ContentFilters = new ObservableCollection<FiltersGroup>();
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					BindingOperations.ClearBinding(icFilters, ItemsControl.ItemsSourceProperty);
					icFilters.SetBinding(ItemsControl.ItemsSourceProperty, XC.A(9860));
					ProjectData.ClearProjectError();
				}
			}
			ContentGroupsCollection allGroups = AllGroups;
			Func<ContentGroup, FiltersGroup> selector;
			if (_Closure_0024__.A == null)
			{
				selector = (_Closure_0024__.A = [SpecialName] (ContentGroup A) => A.FiltersGroup);
			}
			else
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
				selector = _Closure_0024__.A;
			}
			IEnumerable<FiltersGroup> source = ((IEnumerable<ContentGroup>)allGroups).Select(selector);
			Func<FiltersGroup, bool> predicate;
			if (_Closure_0024__.A == null)
			{
				predicate = (_Closure_0024__.A = [SpecialName] (FiltersGroup A) => A != null);
			}
			else
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
				predicate = _Closure_0024__.A;
			}
			IEnumerable<FiltersGroup> enumerable = source.Where(predicate);
			if (ContentFilters.SequenceEqual(enumerable))
			{
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
			try
			{
				ContentFilters.Clear();
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				BindingOperations.ClearBinding(icFilters, ItemsControl.ItemsSourceProperty);
				icFilters.SetBinding(ItemsControl.ItemsSourceProperty, XC.A(9860));
				ProjectData.ClearProjectError();
			}
			IEnumerator<FiltersGroup> enumerator = default(IEnumerator<FiltersGroup>);
			try
			{
				enumerator = enumerable.GetEnumerator();
				while (enumerator.MoveNext())
				{
					FiltersGroup current = enumerator.Current;
					ContentFilters.Add(current);
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
		finally
		{
			this.m_C = false;
		}
	}

	public void CleanUp()
	{
		this.m_A.ClosePreview();
		if (DragDropOverlay != null)
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
			((System.Windows.Window)(object)DragDropOverlay).Close();
			DragDropOverlay = null;
		}
		PowerPointApp.CleanUp(ppApp, HiddenFiles, blnQuitPP);
		ExcelApp.CleanUp(xlApp, HiddenFiles, blnQuitXL);
		MC.A(xlApp);
		xlApp = null;
		HiddenFiles.Clear();
	}

	private DataTemplate A(string A)
	{
		return (DataTemplate)FindResource(A);
	}

	private void ExpandGroup(object sender, RoutedEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		//IL_0050: Unknown result type (might be due to invalid IL or missing references)
		//IL_0055: Unknown result type (might be due to invalid IL or missing references)
		//IL_0057: Unknown result type (might be due to invalid IL or missing references)
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		//IL_005c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0076: Expected I4, but got Unknown
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_007a: Invalid comparison between Unknown and I4
		//IL_007f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0083: Invalid comparison between Unknown and I4
		ContentGroup val = (ContentGroup)((System.Windows.Controls.CheckBox)sender).DataContext;
		ContentGroup val2 = val;
		object allContentItemsLock = val2.AllContentItemsLock;
		ObjectFlowControl.CheckForSyncLockOnValueType(allContentItemsLock);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(allContentItemsLock, ref lockTaken);
			if (val2.AllContentItems == null)
			{
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
					bool flag = false;
					try
					{
						ContentType contentType = val2.ContentType;
						switch (contentType - 3)
						{
						default:
							if ((int)contentType != 14)
							{
								if ((int)contentType == 15)
								{
									val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)E);
								}
							}
							else
							{
								val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)F);
								flag = true;
							}
							break;
						case 0:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)A);
							break;
						case 1:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)B);
							break;
						case 2:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)C);
							break;
						case 4:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)D);
							break;
						case 3:
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
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
						val2.AllContentItems = val2.FiltersGroup.InitialSort();
						val2.ApplyItemsCriteria();
						val2.PopulateFilterFields();
					}
					else if (Conversions.ToBoolean(Operators.NotObject(val.ItemsCriteriaAllowsAll())))
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
						val2.ApplyItemsCriteria();
					}
					if (((Collection<ContentItem>)(object)val2.AllContentItems).Count > 10)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								continue;
							}
							val2.ExpandContent();
							SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
							A().RunWorkerAsync(val);
							break;
						}
					}
					else
					{
						ContentGroup obj = val2;
						BackgroundWorker backgroundWorker = null;
						obj.LoadContentDetail(ref backgroundWorker);
						val2.ExpandContent();
					}
					break;
				}
			}
			else
			{
				val2.ExpandContent();
			}
		}
		finally
		{
			if (lockTaken)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					Monitor.Exit(allContentItemsLock);
					break;
				}
			}
		}
		val2 = null;
		val = null;
	}

	private BackgroundWorker A()
	{
		BackgroundWorker backgroundWorker = new BackgroundWorker();
		backgroundWorker.WorkerSupportsCancellation = true;
		backgroundWorker.WorkerReportsProgress = false;
		backgroundWorker.DoWork += bgw_DoWork;
		backgroundWorker.RunWorkerCompleted += bgw_RunWorkerCompleted;
		_ = null;
		return backgroundWorker;
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		ContentGroup val = (ContentGroup)e.Argument;
		BackgroundWorker backgroundWorker = (BackgroundWorker)sender;
		val.LoadContentDetail(ref backgroundWorker);
		sender = backgroundWorker;
	}

	private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
	}

	private void ShowMoreMenu(object sender, RoutedEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		ContentGroup dataContext = (ContentGroup)((System.Windows.Controls.Button)sender).DataContext;
		popRemoveGroup.DataContext = dataContext;
		popRemoveGroup.PlacementTarget = (UIElement)sender;
		popRemoveGroup.IsOpen = true;
		btnHideContent.Focus();
		dataContext = null;
	}

	private void CloseMorePopup(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Escape)
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
			((Popup)sender).IsOpen = false;
			e.Handled = true;
			return;
		}
	}

	private void MorePopupOpened(object sender, EventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		((ContentGroup)((Popup)sender).DataContext).IsPopupOpen = true;
	}

	private void MorePopupClosed(object sender, EventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		((ContentGroup)((Popup)sender).DataContext).IsPopupOpen = false;
	}

	private void HideContentGroup(object sender, RoutedEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		ContentGroup val = (ContentGroup)((System.Windows.Controls.Button)sender).DataContext;
		PopupMenus.HideGroup(ref val);
		val = null;
	}

	private ContentItem A(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new ShapeItem(A, B);
	}

	private ContentItem B(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new ImageItem(A, B);
	}

	private ContentItem C(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new ChartItem(A, B);
	}

	private ContentItem D(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new TextItem(A, B);
	}

	private ContentItem E(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new PdfItem(A, B);
	}

	private ContentItem F(ContentGroup A, XmlNode B)
	{
		return (ContentItem)(object)new DocumentItem(A, B);
	}

	private bool PreviewIsEnabled()
	{
		return chkPreview.IsChecked.Value;
	}

	private void ListViewItemSelected(object sender, RoutedEventArgs e)
	{
		//IL_000f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0015: Expected O, but got Unknown
		//IL_00ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b4: Invalid comparison between Unknown and I4
		//IL_00c8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d1: Invalid comparison between Unknown and I4
		System.Windows.Controls.ListViewItem listViewItem = (System.Windows.Controls.ListViewItem)sender;
		ContentItem val = (ContentItem)listViewItem.DataContext;
		System.Windows.Controls.ListView listView = (System.Windows.Controls.ListView)ItemsControl.ItemsControlFromItemContainer(listViewItem);
		Base.ScrollToItem(scroller, listViewItem);
		if (ActiveListView != null)
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
			if (listView != ActiveListView)
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
				if (ActiveListView.SelectedItems.Count > 0)
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
					ActiveListView.SelectedItems.Clear();
				}
			}
		}
		ActiveListView = listView;
		A();
		if ((int)val.Group.ContentType != 14)
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
			if ((int)val.Group.ContentType != 15)
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
				SelActionText = XC.A(10224);
				goto IL_0101;
			}
		}
		SelActionText = XC.A(10237);
		goto IL_0101;
		IL_0101:
		Base.IsKeepSourceFormattingAvailable(val, chkKeep);
		listViewItem = null;
		val = null;
		listView = null;
	}

	private void MouseEnterListViewItem(object sender, System.Windows.Input.MouseEventArgs e)
	{
		this.m_A.MouseEnterListViewItem(RuntimeHelpers.GetObjectValue(sender), e, this.m_B);
	}

	private void MouseLeaveListViewItem(object sender, System.Windows.Input.MouseEventArgs e)
	{
		this.m_A.MouseLeaveListViewItem(RuntimeHelpers.GetObjectValue(sender), e);
	}

	private void PreviewToggle(object sender, RoutedEventArgs e)
	{
		this.m_A.PreviewToggle(RuntimeHelpers.GetObjectValue(sender), e);
	}

	private void StarsToggle(object sender, RoutedEventArgs e)
	{
		bool value = chkStars.IsChecked.Value;
		global::A.N.Settings.LibraryPaneShowStars = value;
		IEnumerator<ContentGroup> enumerator = default(IEnumerator<ContentGroup>);
		try
		{
			enumerator = ((Collection<ContentGroup>)(object)AllGroups).GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.ShowFavorites = value;
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
				return;
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
	}

	private void ImageTypeBadgeToggle(object sender, RoutedEventArgs e)
	{
		global::A.N.Settings.LibraryPaneShowImageTypeBadge = chkImageTypeBadge.IsChecked.Value;
	}

	private void OnRequestBringIntoView(object sender, RequestBringIntoViewEventArgs e)
	{
		e.Handled = true;
	}

	private void lstView_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		C();
	}

	private void C()
	{
		if (this.m_H)
		{
			return;
		}
		try
		{
			this.m_H = true;
			IEnumerable<ContentItem> source = ActiveListView.SelectedItems.Cast<ContentItem>();
			Func<ContentItem, bool> predicate;
			if (_Closure_0024__.A == null)
			{
				predicate = (_Closure_0024__.A = [SpecialName] (ContentItem A) => A.Visibility != Visibility.Visible);
			}
			else
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
				predicate = _Closure_0024__.A;
			}
			source.Where(predicate).ToList().ForEach([SpecialName] (ContentItem A) =>
			{
				ActiveListView.SelectedItems.Remove(A);
			});
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			this.m_H = false;
		}
		A();
	}

	private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
	{
		this.m_A.ScrollViewer_PreviewMouseWheel(RuntimeHelpers.GetObjectValue(sender), e);
	}

	private void D()
	{
		IEnumerator<ContentGroup> enumerator = default(IEnumerator<ContentGroup>);
		try
		{
			enumerator = ((Collection<ContentGroup>)(object)AllGroups).GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.ApplyItemsCriteria();
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					enumerator.Dispose();
					break;
				}
			}
		}
	}

	public void DeleteTag(object sender, RoutedEventArgs e)
	{
		this.m_A.DeleteTag(sender as System.Windows.Controls.Button);
	}

	private void E(bool? A, bool? B, bool? C, bool? D)
	{
		if (A.HasValue)
		{
			global::A.N.Settings.ContentInsertShowPersonal = A.Value;
		}
		if (B.HasValue)
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
			global::A.N.Settings.ContentInsertShowShared = B.Value;
		}
		if (C.HasValue)
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
			global::A.N.Settings.ContentInsertShowPublic = C.Value;
		}
		if (D.HasValue)
		{
			global::A.N.Settings.ContentInsertShow3rdParty = D.Value;
		}
		this.A(A: false);
	}

	private void A(ContentType A, bool B)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_001e: Expected I4, but got Unknown
		//IL_001e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0021: Invalid comparison between Unknown and I4
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_003c: Invalid comparison between Unknown and I4
		switch (A - 3)
		{
		default:
			if ((int)A != 14)
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
				if ((int)A != 15)
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
				else
				{
					if (!Pane.PDFsContentIsEnabled)
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
					global::A.N.Settings.ContentInsertShowPDFs = B;
				}
			}
			else
			{
				if (!Pane.DocsContentIsEnabled)
				{
					break;
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
				global::A.N.Settings.ContentInsertShowDocs = B;
			}
			break;
		case 0:
			global::A.N.Settings.ContentInsertShowShapes = B;
			break;
		case 1:
			global::A.N.Settings.ContentInsertShowImages = B;
			break;
		case 2:
			global::A.N.Settings.ContentInsertShowCharts = B;
			break;
		case 4:
			global::A.N.Settings.ContentInsertShowText = B;
			break;
		case 3:
			break;
		}
		this.A(A: false);
	}

	private void F()
	{
		A(A: false);
	}

	private void A(ImageType? A, bool B)
	{
		if (AllGroups != null)
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
			if (AllGroups.ImgTypesFilter != null)
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
				global::A.N.Settings.ContentInsertExcludeImageTypes = AllGroups.ImgTypesFilter.ToListString();
			}
		}
		this.A(A: false);
	}

	internal void A(bool A = false)
	{
		if (SourceCollection == null)
		{
			return;
		}
		IEnumerator<ContentGroup> enumerator = default(IEnumerator<ContentGroup>);
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
			if (GroupsListener.IsChangingGroups)
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
				if (A)
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
					this.A(AllGroups);
				}
				IsFilteringContent = true;
				try
				{
					ContentGroupsCollection allGroups = AllGroups;
					Func<ContentGroup, int> selector;
					if (_Closure_0024__.A == null)
					{
						selector = (_Closure_0024__.A = [SpecialName] (ContentGroup val) =>
						{
							if (val.AllContentItems != null)
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
										return ((Collection<ContentItem>)(object)val.AllContentItems).Count;
									}
								}
							}
							return 0;
						});
					}
					else
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
						selector = _Closure_0024__.A;
					}
					if (((IEnumerable<ContentGroup>)allGroups).Sum(selector) > 100)
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
						XamlUtilities.XamlDoEvents();
					}
					SourceCollection.Refresh();
					try
					{
						enumerator = ((Collection<ContentGroup>)(object)AllGroups).GetEnumerator();
						while (enumerator.MoveNext())
						{
							ContentGroup current = enumerator.Current;
							this.A(current, (bool?)null);
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
					finally
					{
						if (enumerator != null)
						{
							while (true)
							{
								switch (6)
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
				finally
				{
					IsFilteringContent = false;
				}
			}
		}
	}

	internal void G()
	{
		GroupsListener.Update3rdPartyInfo();
	}

	private void A(ContentGroup A)
	{
		if (AllGroups.ShowFavoritesOnly)
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
					this.A(A: false);
					return;
				}
			}
		}
		this.A(A, (bool?)null);
	}

	private void A(ContentGroup A, bool? B = null)
	{
		if (A.FiltersGroup == null)
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
			bool num;
			if (!B.HasValue)
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
				num = this.B(A);
			}
			else
			{
				num = B == true;
			}
			int num2;
			if (!num)
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
				num2 = 2;
			}
			else
			{
				num2 = 0;
			}
			Visibility visibility = (Visibility)num2;
			if (A.FiltersGroup.GroupVisibility == visibility)
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
				A.FiltersGroup.GroupVisibility = visibility;
				return;
			}
		}
	}

	private bool A(object A)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_000c: Expected O, but got Unknown
		return this.A((ContentGroup)A);
	}

	private bool A(ContentGroup A)
	{
		bool flag = B(A);
		if (flag)
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
			A.ApplyItemsCriteria();
		}
		this.A(A, (bool?)flag);
		return flag;
	}

	private bool B(ContentGroup A)
	{
		if (AllGroups == null)
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
					return false;
				}
			}
		}
		if (AllGroups.ShowFavoritesOnly)
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
			if (!A.HasFavorites)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return false;
					}
				}
			}
		}
		return Filter.ApplyLibraryFilter(A, global::A.N.Settings.ContentInsertShowShared, global::A.N.Settings.ContentInsertShowPersonal, global::A.N.Settings.ContentInsertShowPublic, global::A.N.Settings.ContentInsertShow3rdParty, (Func<ContentGroup, bool>)C);
	}

	private bool C(ContentGroup A)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0009: Invalid comparison between Unknown and I4
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0045: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Invalid comparison between Unknown and I4
		//IL_006c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0072: Invalid comparison between Unknown and I4
		//IL_0094: Unknown result type (might be due to invalid IL or missing references)
		//IL_0099: Unknown result type (might be due to invalid IL or missing references)
		//IL_009c: Invalid comparison between Unknown and I4
		//IL_00bb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c4: Invalid comparison between Unknown and I4
		//IL_00f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fd: Invalid comparison between Unknown and I4
		if ((int)A.ContentType == 3)
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
			if (global::A.N.Settings.ContentInsertShowShapes)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						return true;
					}
				}
			}
		}
		if ((int)A.ContentType == 4)
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
			if (global::A.N.Settings.ContentInsertShowImages)
			{
				return true;
			}
		}
		if ((int)A.ContentType == 5)
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
			if (global::A.N.Settings.ContentInsertShowCharts)
			{
				return true;
			}
		}
		if ((int)A.ContentType == 7)
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
			if (global::A.N.Settings.ContentInsertShowText)
			{
				return true;
			}
		}
		if ((int)A.ContentType == 15 && Pane.PDFsContentIsEnabled)
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
			if (global::A.N.Settings.ContentInsertShowPDFs)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						return true;
					}
				}
			}
		}
		if ((int)A.ContentType == 14 && Pane.DocsContentIsEnabled && global::A.N.Settings.ContentInsertShowDocs)
		{
			return true;
		}
		return false;
	}

	private void ContentFiltersChecked(object sender, RoutedEventArgs e)
	{
		grdContentFilters.Visibility = Visibility.Visible;
	}

	private void ContentFiltersUnchecked(object sender, RoutedEventArgs e)
	{
		grdContentFilters.Visibility = Visibility.Collapsed;
	}

	private void ImageTypeFiltersChecked(object sender, RoutedEventArgs e)
	{
		grdImageTypeFilters.Visibility = Visibility.Visible;
	}

	private void ImageTypeFiltersUnchecked(object sender, RoutedEventArgs e)
	{
		grdImageTypeFilters.Visibility = Visibility.Collapsed;
	}

	private void B(ContentGroup A)
	{
		object a = this.m_A;
		ObjectFlowControl.CheckForSyncLockOnValueType(a);
		bool lockTaken = false;
		try
		{
			Monitor.Enter(a, ref lockTaken);
			int A2;
			System.Windows.Application.Current.Dispatcher.Invoke([SpecialName] () =>
			{
				A2 = 1;
			});
		}
		finally
		{
			if (lockTaken)
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
					Monitor.Exit(a);
					break;
				}
			}
		}
	}

	private void A(ContentGroup A, XmlNode B)
	{
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_000e: Expected O, but got Unknown
		A.FiltersGroup = new FiltersGroup((FrameworkElement)this, A, B);
	}

	private void FilterListSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		((System.Windows.Controls.ComboBox)sender).SelectedItem = null;
	}

	private void ListItemCheckChanged(object sender, RoutedEventArgs e)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		ListField field = ((ListOption)((System.Windows.Controls.CheckBox)sender).DataContext).Field;
		field.ComboBoxText = Core.ListOptionText(field.Options);
		A((BaseField)(object)field);
		field = null;
	}

	private void FilterBooleanSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_0032: Expected O, but got Unknown
		System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
		if (comboBox.SelectedIndex == 0)
		{
			comboBox.SelectedIndex = -1;
			e.Handled = true;
		}
		else
		{
			A((BaseField)comboBox.DataContext);
		}
		comboBox = null;
	}

	private void FilterMinDateChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		A((BaseField)((DatePicker)sender).DataContext);
	}

	private void FilterMaxDateChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		A((BaseField)((DatePicker)sender).DataContext);
	}

	private void FilterTextChanged(object sender, TextChangedEventArgs e)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		A((BaseField)((System.Windows.Controls.TextBox)sender).DataContext);
	}

	private void FilterMinValueChanged(object sender, TextChangedEventArgs e)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0016: Expected O, but got Unknown
		A((BaseField)((System.Windows.Controls.TextBox)sender).DataContext);
	}

	private void FilterMaxValueChanged(object sender, TextChangedEventArgs e)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
		A((BaseField)((System.Windows.Controls.TextBox)sender).DataContext);
	}

	private void A(BaseField A)
	{
		A.FiltersChanged();
	}

	private void ResetFilters(object sender, MouseButtonEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		((FiltersGroup)((TextBlock)sender).DataContext).ResetFilters(ref this.m_C);
	}

	private void A(ContentGroup A, bool B)
	{
		try
		{
			if (!B)
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
				C();
				return;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
	}

	private void SortFieldChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		((FiltersGroup)((System.Windows.Controls.ComboBox)sender).DataContext).SortFieldChanged();
	}

	private void SortOrderChanged(object sender, MouseButtonEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		((FiltersGroup)((TextBlock)sender).DataContext).SortOrderChanged();
	}

	private void MenuOpening(object sender, ContextMenuEventArgs e)
	{
		//IL_000c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0012: Expected O, but got Unknown
		System.Windows.Controls.ListViewItem obj = (System.Windows.Controls.ListViewItem)sender;
		ContentItem val = (ContentItem)obj.DataContext;
		System.Windows.Controls.ContextMenu contextMenu = obj.ContextMenu;
		D(val.Group);
		this.m_B = true;
		Base.ShowHideContextMenuOptionForContentTypes(contextMenu, val, XC.A(10320), true, (ContentType[])(object)new ContentType[1] { (ContentType)4 });
		val = null;
	}

	private void MenuClosing(object sender, ContextMenuEventArgs e)
	{
		this.m_B = false;
	}

	private System.Windows.Controls.ListViewItem A(object A)
	{
		return (System.Windows.Controls.ListViewItem)((System.Windows.Controls.ContextMenu)((System.Windows.Controls.MenuItem)A).CommandParameter).PlacementTarget;
	}

	private ContentItem A(object A)
	{
		//IL_0017: Unknown result type (might be due to invalid IL or missing references)
		//IL_001d: Expected O, but got Unknown
		return (ContentItem)this.A(RuntimeHelpers.GetObjectValue(A)).DataContext;
	}

	private void KeepSourceFormatToggle(object sender, RoutedEventArgs e)
	{
		global::A.N.Settings.LibraryPaneKeepSourceFormat = chkKeep.IsChecked.Value;
	}

	private void btnInsert_Click(object sender, RoutedEventArgs e)
	{
		H();
	}

	private void InsertContent(object sender, RoutedEventArgs e)
	{
		H();
	}

	private void MouseDblClickListViewItem(object sender, MouseButtonEventArgs e)
	{
		H();
	}

	private void H()
	{
		//IL_0071: Unknown result type (might be due to invalid IL or missing references)
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_007b: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Unknown result type (might be due to invalid IL or missing references)
		//IL_007e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_009a: Expected I4, but got Unknown
		//IL_009a: Unknown result type (might be due to invalid IL or missing references)
		//IL_009d: Invalid comparison between Unknown and I4
		//IL_00a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ac: Invalid comparison between Unknown and I4
		if (ActiveListView == null || ActiveListView.SelectedItems.Count <= 0)
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
			UndoRecord undoRecord = PC.A.Application.UndoRecord;
			undoRecord.StartCustomRecord(XC.A(10343));
			ContentType contentType = ((ContentGroup)ActiveListView.DataContext).ContentType;
			switch (contentType - 3)
			{
			default:
				if ((int)contentType != 14)
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
					if ((int)contentType != 15)
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
					}
					else
					{
						M();
					}
				}
				else
				{
					N();
				}
				break;
			case 2:
				I();
				break;
			case 0:
				J();
				break;
			case 1:
				K();
				break;
			case 4:
				L();
				break;
			case 3:
				break;
			}
			undoRecord.EndCustomRecord();
			undoRecord = null;
			UsageLogger.LogInsertion(ActiveListView.SelectedItems.Cast<ContentItem>(), (OfficeApp)3);
			return;
		}
	}

	private void B(string A)
	{
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, A);
	}

	private void I()
	{
		//IL_00f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f9: Expected O, but got Unknown
		//IL_0102: Unknown result type (might be due to invalid IL or missing references)
		//IL_010e: Expected O, but got Unknown
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		if (!Access.AllowWordOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			Workbook workbook = null;
			InlineShape inlineShape = null;
			bool value = chkKeep.IsChecked.Value;
			string c = Content.RightNow();
			string d = A();
			workbook = ExcelApp.GetWorkbook(ref xlApp, ref HiddenFiles, ref blnQuitXL, ((ContentItem)ActiveListView.SelectedItems[0]).ContentPath);
			if (workbook != null)
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
				Microsoft.Office.Interop.Word.Application application = PC.A.Application;
				application.ScreenUpdating = false;
				bool copyObjectsWithCells = xlApp.CopyObjectsWithCells;
				xlApp.CopyObjectsWithCells = true;
				Workbook workbook2;
				Microsoft.Office.Interop.Excel.Chart chart;
				try
				{
					try
					{
						enumerator = ActiveListView.SelectedItems.GetEnumerator();
						while (enumerator.MoveNext())
						{
							ContentItem val = (ContentItem)enumerator.Current;
							workbook2 = ExcelApp.CreateTargetWorkbook(xlApp, workbook, (ChartItem)val, value);
							chart = ((ChartObject)((Worksheet)workbook2.Worksheets[1]).ChartObjects(1)).Chart;
							try
							{
								if (!xlApp.Visible)
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
									Forms.InfoMessage(XC.A(10388));
								}
								((ChartObject)chart.Parent).Select(RuntimeHelpers.GetObjectValue(Missing.Value));
								clsClipboard.CopyWithWait((Action)([SpecialName] () =>
								{
									xlApp.CommandBars.ExecuteMso(XC.A(11557));
									System.Windows.Forms.Application.DoEvents();
								}), 4000);
								if (value)
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
									application.CommandBars.ExecuteMso(XC.A(10694));
								}
								else
								{
									application.CommandBars.ExecuteMso(XC.A(10757));
								}
								application.CommandBars.ReleaseFocus();
								System.Windows.Forms.Application.DoEvents();
								try
								{
									Microsoft.Office.Interop.Word.Selection selection = application.Selection;
									object Unit = WdUnits.wdCharacter;
									object Count = 1;
									object Extend = WdMovementType.wdExtend;
									selection.MoveLeft(ref Unit, ref Count, ref Extend);
									inlineShape = application.Selection.InlineShapes[1];
									inlineShape.Chart.ChartData.ActivateChartDataWindow();
								}
								catch (Exception ex)
								{
									ProjectData.SetProjectError(ex);
									Exception ex2 = ex;
									ProjectData.ClearProjectError();
								}
								try
								{
									if (value)
									{
										string text = Path.Combine(Interaction.Environ(XC.A(10820)), XC.A(10829));
										chart.SaveChartTemplate(text);
										inlineShape.Chart.ApplyChartTemplate(text);
										File.Delete(text);
									}
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									ProjectData.ClearProjectError();
								}
								A(inlineShape, val, c, d);
								inlineShape = null;
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								clsReporting.LogException(ex6);
								A(chart, application);
								ProjectData.ClearProjectError();
							}
							try
							{
								workbook2.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							catch (Exception ex7)
							{
								ProjectData.SetProjectError(ex7);
								Exception ex8 = ex7;
								ProjectData.ClearProjectError();
							}
						}
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								goto end_IL_0328;
							}
							continue;
							end_IL_0328:
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								(enumerator as IDisposable).Dispose();
								break;
							}
						}
					}
					B(XC.A(10856));
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					Forms.ErrorMessage(XC.A(10899) + ex10.Message);
					clsReporting.LogException(ex10);
					ProjectData.ClearProjectError();
				}
				xlApp.CopyObjectsWithCells = copyObjectsWithCells;
				try
				{
					workbook.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				}
				catch (Exception ex11)
				{
					ProjectData.SetProjectError(ex11);
					Exception ex12 = ex11;
					ProjectData.ClearProjectError();
				}
				application.ScreenUpdating = true;
				workbook2 = null;
				workbook = null;
				chart = null;
				application = null;
			}
			if (xlApp == null)
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
				xlApp.DisplayAlerts = true;
				xlApp.EnableEvents = true;
				xlApp.ScreenUpdating = true;
				return;
			}
		}
	}

	private void A(Microsoft.Office.Interop.Excel.Chart A, Microsoft.Office.Interop.Word.Application B)
	{
		clsClipboard.CopyWithWait((Action)([SpecialName] () =>
		{
			A.ChartArea.Copy();
		}), 4000);
		Microsoft.Office.Interop.Word.Range range = B.Selection.Range;
		object IconIndex = RuntimeHelpers.GetObjectValue(Missing.Value);
		object Link = false;
		object Placement = WdOLEPlacement.wdInLine;
		object DisplayAsIcon = RuntimeHelpers.GetObjectValue(Missing.Value);
		object DataType = WdPasteDataType.wdPasteOLEObject;
		object IconFileName = RuntimeHelpers.GetObjectValue(Missing.Value);
		object IconLabel = RuntimeHelpers.GetObjectValue(Missing.Value);
		range.PasteSpecial(ref IconIndex, ref Link, ref Placement, ref DisplayAsIcon, ref DataType, ref IconFileName, ref IconLabel);
	}

	private void J()
	{
		//IL_009c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a3: Expected O, but got Unknown
		//IL_0065: Unknown result type (might be due to invalid IL or missing references)
		//IL_00eb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f6: Expected O, but got Unknown
		//IL_00b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00df: Expected O, but got Unknown
		Y a = default(Y);
		Y CS_0024_003C_003E8__locals7 = new Y(a);
		Presentation presentation = null;
		string c = Content.RightNow();
		string d = A();
		if (Insert.AreShapesLegacyPublished(ActiveListView))
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
			presentation = PowerPointApp.GetPresentation(ref ppApp, ref HiddenFiles, ref blnQuitPP, ((ContentItem)ActiveListView.SelectedItems[0]).ContentPath);
		}
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ActiveListView.SelectedItems.GetEnumerator();
				while (enumerator.MoveNext())
				{
					ContentItem val = (ContentItem)enumerator.Current;
					if (presentation == null)
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
						CS_0024_003C_003E8__locals7.A = PowerPointApp.GetSourceShape((ShapeItem)val, PowerPointApp.GetPresentation(ref ppApp, ref HiddenFiles, ref blnQuitPP, val.ContentPath));
					}
					else
					{
						CS_0024_003C_003E8__locals7.A = PowerPointApp.GetSourceShape((ShapeItem)val, presentation);
					}
					Action action;
					if (CS_0024_003C_003E8__locals7.A != null)
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
						action = CS_0024_003C_003E8__locals7.A;
					}
					else
					{
						action = (CS_0024_003C_003E8__locals7.A = [SpecialName] () =>
						{
							CS_0024_003C_003E8__locals7.A.Copy();
						});
					}
					clsClipboard.CopyWithWait(action, 4000);
					try
					{
						Microsoft.Office.Interop.Word.Application application = PC.A.Application;
						application.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);
						Microsoft.Office.Interop.Word.ShapeRange shapeRange = application.Selection.ShapeRange;
						object Index = 1;
						A(shapeRange[ref Index], val, c, d);
						application = null;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
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
						goto end_IL_019e;
					}
					continue;
					end_IL_019e:
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
			B(XC.A(10994));
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			Forms.ErrorMessage(ex4.Message);
			ProjectData.ClearProjectError();
		}
		presentation = null;
		CS_0024_003C_003E8__locals7.A = null;
	}

	private void K()
	{
		//IL_0047: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Expected O, but got Unknown
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		string c = Content.RightNow();
		string d = A();
		application.ScreenUpdating = false;
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ActiveListView.SelectedItems.GetEnumerator();
				while (enumerator.MoveNext())
				{
					ContentItem val = (ContentItem)enumerator.Current;
					InlineShapes inlineShapes = application.Selection.Range.InlineShapes;
					string contentPath = val.ContentPath;
					object LinkToFile = false;
					object SaveWithDocument = true;
					object Range = RuntimeHelpers.GetObjectValue(Missing.Value);
					InlineShape inlineShape = inlineShapes.AddPicture(contentPath, ref LinkToFile, ref SaveWithDocument, ref Range);
					PageSetup pageSetup = application.ActiveDocument.PageSetup;
					float num = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;
					float num2 = pageSetup.PageHeight - pageSetup.TopMargin - pageSetup.BottomMargin;
					pageSetup = null;
					InlineShape inlineShape2 = inlineShape;
					inlineShape2.LockAspectRatio = MsoTriState.msoTrue;
					if (inlineShape2.Width / inlineShape2.Height > num / num2)
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
						if (inlineShape2.Width > num)
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
							inlineShape2.Width = num;
						}
					}
					else if (inlineShape2.Height > num2)
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
						inlineShape2.Height = num2;
					}
					inlineShape2 = null;
					A(inlineShape, val, c, d);
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
			B(XC.A(11037));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		application = null;
	}

	private void L()
	{
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.Word.Application application = PC.A.Application;
		UndoRecord undoRecord = application.UndoRecord;
		undoRecord.StartCustomRecord(XC.A(11080));
		application.ScreenUpdating = false;
		Microsoft.Office.Interop.Word.Selection selection;
		try
		{
			selection = application.Selection;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ActiveListView.SelectedItems.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				IEnumerator enumerator3 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					string text = File.ReadAllText(((ContentItem)enumerator.Current).ContentPath);
					WdSelectionType type = selection.Type;
					if (type <= WdSelectionType.wdSelectionNormal)
					{
						if (type != WdSelectionType.wdSelectionIP)
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
							if (type == WdSelectionType.wdSelectionNormal)
							{
								selection.Text = text;
							}
						}
						else
						{
							selection.InsertAfter(text);
						}
						continue;
					}
					if (type != WdSelectionType.wdSelectionInlineShape)
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
						if (type != WdSelectionType.wdSelectionShape)
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
							continue;
						}
						try
						{
							enumerator2 = Helpers.SelectedShapes(selection).GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Microsoft.Office.Interop.Word.Shape a = (Microsoft.Office.Interop.Word.Shape)enumerator2.Current;
								A(a, text);
							}
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_016c;
								}
								continue;
								end_IL_016c:
								break;
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
						continue;
					}
					try
					{
						enumerator3 = selection.ChildShapeRange.GetEnumerator();
						while (enumerator3.MoveNext())
						{
							Microsoft.Office.Interop.Word.Shape a2 = (Microsoft.Office.Interop.Word.Shape)enumerator3.Current;
							A(a2, text);
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
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_01a6;
					}
					continue;
					end_IL_01a6:
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
			B(XC.A(11103));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		application.ScreenUpdating = true;
		undoRecord.EndCustomRecord();
		application = null;
		selection = null;
		undoRecord = null;
	}

	private void A(Microsoft.Office.Interop.Word.Shape A, string B)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			A.TextFrame.TextRange.Text = B;
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape a = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
				this.A(a, B);
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

	private void M()
	{
		B(XC.A(11142));
	}

	private void N()
	{
		//IL_00ad: Unknown result type (might be due to invalid IL or missing references)
		if (ActiveListView.SelectedItems.Count > 5)
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
			if (System.Windows.Forms.MessageBox.Show(XC.A(11159) + ActiveListView.SelectedItems.Count + XC.A(11228), XC.A(2438), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
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
				break;
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ActiveListView.SelectedItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Documents.OpenFile(((ContentItem)enumerator.Current).ContentPath);
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_00c6;
				}
				continue;
				end_IL_00c6:
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
		B(XC.A(11269));
	}

	private void O()
	{
		H();
	}

	private void lstView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
	{
		this.m_A = e.GetPosition(null);
	}

	private void lstView_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
	{
		DragDrop.ListViewMouseMove(RuntimeHelpers.GetObjectValue(sender), e, this.m_A, ref this.m_A);
	}

	private void lstView_DragLeave(object sender, System.Windows.DragEventArgs e)
	{
		//IL_004b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0055: Expected O, but got Unknown
		if (DragDropOverlay != null)
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
			try
			{
				DragDropOverlay = new wpfDragDrop((Action)O, (IntPtr)PC.A.Application.ActiveWindow.Hwnd);
				wpfDragDrop dragDropOverlay = DragDropOverlay;
				((System.Windows.Window)(object)dragDropOverlay).Closed += DragDropOverlayClosed;
				((System.Windows.Window)(object)dragDropOverlay).ShowActivated = false;
				((System.Windows.Window)(object)dragDropOverlay).Show();
				_ = null;
				return;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	private void DragDropOverlayClosed(object sender, EventArgs e)
	{
		DragDropOverlay = null;
	}

	private void DownloadContent(object sender, RoutedEventArgs e)
	{
		Base.DoContentDownload(A(RuntimeHelpers.GetObjectValue(sender)));
	}

	private void A(Microsoft.Office.Interop.Word.Shape A, ContentItem B, string C, string D)
	{
		Tagging.A(A, B, C, D);
	}

	private void A(InlineShape A, ContentItem B, string C, string D)
	{
		Tagging.A(A, B, C, D);
	}

	private string A()
	{
		return Core.GetAuthor(PC.A.Application.ActiveDocument);
	}

	private bool D(ContentGroup A)
	{
		return Access.UserHasAccess(A.Library, (AccessType)1, IsAdmin);
	}

	private void P()
	{
		Forms.WarningMessage(XC.A(11314));
	}

	private ContentGroup A(ContentType A)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		return ((IEnumerable<ContentGroup>)AllGroups).First([SpecialName] (ContentGroup val) => val.ContentType == A);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_K)
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
			this.m_K = true;
			Uri resourceLocator = new Uri(XC.A(11450), UriKind.Relative);
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
		if (connectionId == 13)
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
					((System.Windows.Controls.MenuItem)target).Click += InsertContent;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					((System.Windows.Controls.MenuItem)target).Click += DownloadContent;
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
					popSuggest = (Popup)target;
					return;
				}
			}
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					rtbSearch = (System.Windows.Controls.RichTextBox)target;
					return;
				}
			}
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
					chkShared = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkPersonal = (System.Windows.Controls.CheckBox)target;
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
					chk3rdParty = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkPublic = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			scroller = (ScrollViewer)target;
			scroller.PreviewMouseWheel += ScrollViewer_PreviewMouseWheel;
			return;
		}
		if (connectionId == 22)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkFilters = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					stkFilters = (StackPanel)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			chkContentFilters = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 25)
		{
			grdContentFilters = (DockPanel)target;
			return;
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkCharts = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkShapes = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 28)
		{
			chkImages = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 29)
		{
			chkText = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 30)
		{
			chkPdfs = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 31)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkDocs = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 32)
		{
			chkFavorites = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 33)
		{
			chkImageTypeFilters = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 34)
		{
			grdImageTypeFilters = (StackPanel)target;
			return;
		}
		if (connectionId == 35)
		{
			chkImagesSvg = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 36)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkImagesPng = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 37)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkImagesJpg = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 38)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkImagesEmf = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 39)
		{
			chkImagesWmf = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 40)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkImagesGif = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 41)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkImagesBmp = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 42)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkImagesTiff = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 43)
		{
			icFilters = (ItemsControl)target;
			return;
		}
		if (connectionId == 47)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					icContent = (ItemsControl)target;
					return;
				}
			}
		}
		if (connectionId == 51)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkPreview = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 52)
		{
			chkStars = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 53)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkImageTypeBadge = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 54)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnInsert = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 55)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkKeep = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		switch (connectionId)
		{
		case 56:
			while (true)
			{
				switch (4)
				{
				case 0:
					continue;
				}
				popRemoveGroup = (Popup)target;
				popRemoveGroup.Opened += MorePopupOpened;
				popRemoveGroup.Closed += MorePopupClosed;
				popRemoveGroup.PreviewKeyDown += CloseMorePopup;
				return;
			}
		case 57:
			btnHideContent = (System.Windows.Controls.Button)target;
			btnHideContent.Click += HideContentGroup;
			break;
		default:
			this.m_K = true;
			break;
		}
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
			((System.Windows.Controls.Button)target).Click += DeleteTag;
		}
		if (connectionId == 2)
		{
			((System.Windows.Controls.Button)target).Click += ShowMoreMenu;
		}
		if (connectionId == 3)
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
			((System.Windows.Controls.TextBox)target).TextChanged += FilterTextChanged;
		}
		if (connectionId == 4)
		{
			((System.Windows.Controls.ComboBox)target).SelectionChanged += FilterListSelectionChanged;
		}
		if (connectionId == 5)
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
			((System.Windows.Controls.CheckBox)target).Checked += ListItemCheckChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += ListItemCheckChanged;
		}
		if (connectionId == 6)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += FilterBooleanSelectionChanged;
		}
		if (connectionId == 7)
		{
			((DatePicker)target).SelectedDateChanged += FilterMinDateChanged;
		}
		if (connectionId == 8)
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
			((DatePicker)target).SelectedDateChanged += FilterMaxDateChanged;
		}
		if (connectionId == 9)
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
			((System.Windows.Controls.TextBox)target).TextChanged += FilterMinValueChanged;
		}
		if (connectionId == 10)
		{
			((System.Windows.Controls.TextBox)target).TextChanged += FilterMaxValueChanged;
		}
		if (connectionId == 11)
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
			((System.Windows.Controls.TextBox)target).TextChanged += FilterMinValueChanged;
		}
		if (connectionId == 12)
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
			((System.Windows.Controls.TextBox)target).TextChanged += FilterMaxValueChanged;
		}
		if (connectionId == 44)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += SortFieldChanged;
		}
		if (connectionId == 45)
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
			((TextBlock)target).MouseUp += SortOrderChanged;
		}
		if (connectionId == 46)
		{
			((TextBlock)target).MouseUp += ResetFilters;
		}
		if (connectionId == 48)
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
			((System.Windows.Controls.CheckBox)target).Checked += ExpandGroup;
		}
		if (connectionId == 49)
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
			((System.Windows.Controls.ListView)target).PreviewMouseLeftButtonDown += lstView_PreviewMouseLeftButtonDown;
			((System.Windows.Controls.ListView)target).MouseMove += lstView_MouseMove;
			((System.Windows.Controls.ListView)target).DragLeave += lstView_DragLeave;
			((System.Windows.Controls.ListView)target).SelectionChanged += lstView_SelectionChanged;
		}
		if (connectionId == 50)
		{
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = FrameworkElement.ContextMenuOpeningEvent;
			eventSetter.Handler = new ContextMenuEventHandler(MenuOpening);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = FrameworkElement.ContextMenuClosingEvent;
			eventSetter.Handler = new ContextMenuEventHandler(MenuClosing);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = ListBoxItem.SelectedEvent;
			eventSetter.Handler = new RoutedEventHandler(ListViewItemSelected);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseEnterEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(MouseEnterListViewItem);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseLeaveEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(MouseLeaveListViewItem);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = System.Windows.Controls.Control.MouseDoubleClickEvent;
			eventSetter.Handler = new MouseButtonEventHandler(MouseDblClickListViewItem);
			((Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = FrameworkElement.RequestBringIntoViewEvent;
			eventSetter.Handler = new RequestBringIntoViewEventHandler(OnRequestBringIntoView);
			((Style)target).Setters.Add(eventSetter);
		}
	}

	void IStyleConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IStyleConnector_Connect
		this.System_Windows_Markup_IStyleConnector_Connect(connectionId, target);
	}

	[SpecialName]
	[CompilerGenerated]
	private void A(ContentItem A)
	{
		ActiveListView.SelectedItems.Remove(A);
	}

	[SpecialName]
	[CompilerGenerated]
	private void Q()
	{
		xlApp.CommandBars.ExecuteMso(XC.A(11557));
		System.Windows.Forms.Application.DoEvents();
	}
}
