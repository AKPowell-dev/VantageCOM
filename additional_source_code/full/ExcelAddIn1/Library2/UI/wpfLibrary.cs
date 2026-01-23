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
using ExcelAddIn1.Library2.Insert;
using ExcelAddIn1.Library2.Versioning;
using ExcelAddIn1.Sheets;
using ExcelAddIn1.Workbook;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.Libraries;
using MacabacusMacros.Libraries.Manage.Publish;
using MacabacusMacros.Libraries.Pane;
using MacabacusMacros.Libraries.Pane.Filters;
using MacabacusMacros.Libraries.Pane.UI;
using MacabacusMacros.Libraries.Tags;
using MacabacusMacros.Libraries.Versioning;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using MacabacusMacros.UI.FormsExtensions;
using MacabacusMacros.Xaml;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Library2.UI;

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
			return A.Visibility != System.Windows.Visibility.Visible;
		}

		[SpecialName]
		internal int A(ContentGroup A)
		{
			RangeObservableCollection<ContentItem> allContentItems = A.AllContentItems;
			if (allContentItems == null)
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
						return 0;
					}
				}
			}
			return ((Collection<ContentItem>)(object)allContentItems).Count;
		}
	}

	[CompilerGenerated]
	internal sealed class BF
	{
		public int A;

		[SpecialName]
		internal void A()
		{
			this.A = 1;
		}
	}

	[CompilerGenerated]
	internal sealed class CF
	{
		public ContentType A;

		[SpecialName]
		internal bool A(ContentGroup A)
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0006: Unknown result type (might be due to invalid IL or missing references)
			//IL_0009: Unknown result type (might be due to invalid IL or missing references)
			return A.ContentType == this.A;
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	public List<string> HiddenFiles;

	public Microsoft.Office.Interop.PowerPoint.Application ppApp;

	public bool blnQuitPP;

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

	[AccessedThroughProperty("popSuggest")]
	[CompilerGenerated]
	private Popup m_A;

	[AccessedThroughProperty("rtbSearch")]
	[CompilerGenerated]
	private System.Windows.Controls.RichTextBox m_A;

	[AccessedThroughProperty("chkShared")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("chkPersonal")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("chk3rdParty")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPublic")]
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

	[AccessedThroughProperty("grdContentFilters")]
	[CompilerGenerated]
	private DockPanel m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkCharts")]
	private System.Windows.Controls.CheckBox m_G;

	[AccessedThroughProperty("chkTables")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_H;

	[AccessedThroughProperty("chkShapes")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_I;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImages")]
	private System.Windows.Controls.CheckBox m_J;

	[AccessedThroughProperty("chkText")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_K;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPdfs")]
	private System.Windows.Controls.CheckBox m_L;

	[AccessedThroughProperty("chkModels")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_M;

	[AccessedThroughProperty("chkFavorites")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_N;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImageTypeFilters")]
	private System.Windows.Controls.CheckBox m_O;

	[AccessedThroughProperty("grdImageTypeFilters")]
	[CompilerGenerated]
	private StackPanel m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesSvg")]
	private System.Windows.Controls.CheckBox m_P;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesPng")]
	private System.Windows.Controls.CheckBox m_Q;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesJpg")]
	private System.Windows.Controls.CheckBox R;

	[AccessedThroughProperty("chkImagesEmf")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox S;

	[AccessedThroughProperty("chkImagesWmf")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox T;

	[AccessedThroughProperty("chkImagesGif")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox U;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesBmp")]
	private System.Windows.Controls.CheckBox V;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesTiff")]
	private System.Windows.Controls.CheckBox W;

	[CompilerGenerated]
	[AccessedThroughProperty("icFilters")]
	private ItemsControl m_A;

	[AccessedThroughProperty("icContent")]
	[CompilerGenerated]
	private ItemsControl m_B;

	[AccessedThroughProperty("chkPreview")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox X;

	[CompilerGenerated]
	[AccessedThroughProperty("chkStars")]
	private System.Windows.Controls.CheckBox Y;

	[AccessedThroughProperty("chkImageTypeBadge")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox Z;

	[AccessedThroughProperty("btnInsert")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkKeep")]
	private System.Windows.Controls.CheckBox AB;

	[AccessedThroughProperty("popRemoveGroup")]
	[CompilerGenerated]
	private Popup m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnHideContent")]
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
			A(VH.A(36261));
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
			A(VH.A(87682));
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
			A(VH.A(87701));
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
			A(VH.A(87730));
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
				switch (5)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.m_A = value;
				A(VH.A(87759));
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
				switch (4)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.m_B = value;
				A(VH.A(87792));
				return;
			}
		}
	}

	private FrameworkElement PreviewParentUIElem => this;

	private bool PreviewSetting
	{
		get
		{
			return global::A.K.Settings.LibraryPaneShowPreview;
		}
		set
		{
			global::A.K.Settings.LibraryPaneShowPreview = value;
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
			A(VH.A(88080));
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
			A(VH.A(88117));
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_F = value;
			checkBox = this.m_F;
			if (checkBox == null)
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
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
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

	internal virtual System.Windows.Controls.CheckBox chkTables
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

	internal virtual System.Windows.Controls.CheckBox chkShapes
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

	internal virtual System.Windows.Controls.CheckBox chkImages
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

	internal virtual System.Windows.Controls.CheckBox chkText
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

	internal virtual System.Windows.Controls.CheckBox chkPdfs
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

	internal virtual System.Windows.Controls.CheckBox chkModels
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

	internal virtual System.Windows.Controls.CheckBox chkFavorites
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
			this.m_N = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImageTypeFilters
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
			RoutedEventHandler value2 = ImageTypeFiltersChecked;
			RoutedEventHandler value3 = ImageTypeFiltersUnchecked;
			System.Windows.Controls.CheckBox checkBox = this.m_O;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_O = value;
			checkBox = this.m_O;
			if (checkBox == null)
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
			return this.m_P;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_P = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesPng
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

	internal virtual System.Windows.Controls.CheckBox chkImagesJpg
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

	internal virtual System.Windows.Controls.CheckBox chkImagesEmf
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

	internal virtual System.Windows.Controls.CheckBox chkImagesWmf
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

	internal virtual System.Windows.Controls.CheckBox chkImagesGif
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

	internal virtual System.Windows.Controls.CheckBox chkImagesBmp
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

	internal virtual System.Windows.Controls.CheckBox chkImagesTiff
	{
		[CompilerGenerated]
		get
		{
			return W;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			W = value;
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
			return X;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			X = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkStars
	{
		[CompilerGenerated]
		get
		{
			return Y;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Y = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImageTypeBadge
	{
		[CompilerGenerated]
		get
		{
			return Z;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Z = value;
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

	internal virtual System.Windows.Controls.CheckBox chkKeep
	{
		[CompilerGenerated]
		get
		{
			return AB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			AB = value;
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

	public wpfLibrary()
	{
		//IL_00a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b3: Expected O, but got Unknown
		//IL_00df: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f5: Expected O, but got Unknown
		//IL_00f0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fa: Expected O, but got Unknown
		//IL_012a: Unknown result type (might be due to invalid IL or missing references)
		//IL_013b: Expected O, but got Unknown
		//IL_0142: Unknown result type (might be due to invalid IL or missing references)
		//IL_014c: Expected O, but got Unknown
		base.Loaded += wpfLibrary_Loaded;
		base.Unloaded += wpfLibrary_Unloaded;
		blnQuitPP = false;
		this.m_B = false;
		this.m_A = null;
		this.m_A = "";
		this.m_B = VH.A(57289);
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
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(A));
			return;
		}
	}

	private void A()
	{
		System.Windows.Controls.ListView activeListView = ActiveListView;
		int? obj;
		if (activeListView == null)
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
			obj = null;
		}
		else
		{
			IList selectedItems = activeListView.SelectedItems;
			if (selectedItems == null)
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
			else
			{
				obj = selectedItems.Count;
			}
		}
		int? num = obj;
		int valueOrDefault = num.GetValueOrDefault();
		SelectedCountStr = ((valueOrDefault < 2) ? "" : string.Format(VH.A(48282), valueOrDefault));
	}

	private void wpfLibrary_Loaded(object sender, RoutedEventArgs e)
	{
		clsPanes.EliminateTaskPaneFlicker(this);
		this.m_A.OnLoad();
		chkPreview.IsChecked = global::A.K.Settings.LibraryPaneShowPreview;
		chkPreview.Checked += PreviewToggle;
		chkPreview.Unchecked += PreviewToggle;
		chkStars.IsChecked = global::A.K.Settings.LibraryPaneShowStars;
		chkStars.Checked += StarsToggle;
		chkStars.Unchecked += StarsToggle;
		chkImageTypeBadge.IsChecked = global::A.K.Settings.LibraryPaneShowImageTypeBadge;
		chkImageTypeBadge.Checked += ImageTypeBadgeToggle;
		chkImageTypeBadge.Unchecked += ImageTypeBadgeToggle;
		chkKeep.IsChecked = global::A.K.Settings.LibraryPaneKeepSourceFormat;
		chkKeep.Checked += KeepSourceFormatToggle;
		chkKeep.Unchecked += KeepSourceFormatToggle;
	}

	private void wpfLibrary_Unloaded(object sender, RoutedEventArgs e)
	{
		//IL_003b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0045: Expected O, but got Unknown
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Expected O, but got Unknown
		this.m_G = true;
		CleanUp();
		Pane.A(this);
		if (AllGroups != null)
		{
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
					break;
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
		if (GroupsListener != null)
		{
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
		MySettings settings = global::A.K.Settings;
		A.ShowsPersonalLibrary = settings.ContentInsertShowPersonal;
		A.ShowsSharedLibraries = settings.ContentInsertShowShared;
		A.ShowsPublicLibrary = settings.ContentInsertShowPublic;
		A.Shows3rdPartyLibraries = settings.ContentInsertShow3rdParty;
		A[(ContentType)6] = settings.ContentInsertShowTables;
		A[(ContentType)3] = settings.ContentInsertShowShapes;
		A[(ContentType)4] = settings.ContentInsertShowImages;
		A[(ContentType)5] = settings.ContentInsertShowCharts;
		A[(ContentType)7] = settings.ContentInsertShowText;
		int num;
		if (settings.ContentInsertShowPDFs)
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
			num = (Pane.PDFsContentIsEnabled ? 1 : 0);
		}
		else
		{
			num = 0;
		}
		A[(ContentType)15] = (byte)num != 0;
		int num2;
		if (settings.ContentInsertShowModels)
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
			num2 = (Pane.ModelsContentIsEnabled ? 1 : 0);
		}
		else
		{
			num2 = 0;
		}
		A[(ContentType)13] = (byte)num2 != 0;
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
			switch (3)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			this.m_A = KH.A.SettingsXml;
			this.m_D = global::A.K.Settings.LibraryPaneShowStars;
			HiddenFiles = new List<string>();
			IsAdmin = Base.IsUserAdmin();
			ItemsPanelWrap = (ItemsPanelTemplate)FindResource(VH.A(87819));
			ItemsPanelStack = (ItemsPanelTemplate)FindResource(VH.A(87848));
			AllGroups = Load.GetLibraryContent((Func<XmlDocument, string, LibraryItem, bool, ContentGroup>)A, (Action<ContentGroupsCollection>)A);
			B();
			GroupsListener.Groups = AllGroups;
			SourceCollection = CollectionViewSource.GetDefaultView(AllGroups);
			ICollectionView sourceCollection = SourceCollection;
			if (!sourceCollection.GroupDescriptions.Any())
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
				sourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(VH.A(84167)));
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
		//IL_0202: Unknown result type (might be due to invalid IL or missing references)
		//IL_0206: Invalid comparison between Unknown and I4
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0063: Invalid comparison between Unknown and I4
		//IL_009e: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00db: Expected I4, but got Unknown
		//IL_014a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0155: Unknown result type (might be due to invalid IL or missing references)
		//IL_0158: Invalid comparison between Unknown and I4
		//IL_017c: Unknown result type (might be due to invalid IL or missing references)
		//IL_019d: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a4: Expected O, but got Unknown
		//IL_01ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b7: Expected O, but got Unknown
		//IL_01c0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ca: Expected O, but got Unknown
		ContentGroup val = null;
		XmlNodeList childNodes = A.DocumentElement.ChildNodes;
		if (childNodes != null)
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
			if (childNodes.Count > 0)
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
				string directoryName = Path.GetDirectoryName(B);
				ContentType contentType = Manifests.GetContentType(A);
				Base.ProcessManifestNodes(ref childNodes, directoryName, contentType);
				if ((int)contentType == 13)
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
					if (!Access.IsEnterprisePlanOrTrialMode())
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
						if (!Access.IsLegacyPlan())
						{
							while (true)
							{
								switch (3)
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
				}
				string a;
				switch (contentType - 3)
				{
				case 3:
					a = VH.A(87879);
					break;
				case 0:
					a = VH.A(87898);
					break;
				case 1:
					a = VH.A(87917);
					break;
				case 2:
					a = VH.A(87936);
					break;
				case 4:
					a = VH.A(87955);
					break;
				case 12:
					a = VH.A(87972);
					break;
				case 10:
					a = VH.A(87987);
					break;
				default:
					childNodes = null;
					return null;
				}
				bool flag = Favorites.ManifestContainsFavorite(A, contentType);
				ItemsPanelTemplate itemsPanelTemplate;
				if ((int)contentType != 7)
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
					this.C(VH.A(88006) + directoryName);
					ProjectData.ClearProjectError();
				}
				itemsPanelTemplate = null;
				if (val != null)
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
					if ((int)contentType == 13)
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
						XmlNode schemaNodeById = CustomFields.GetSchemaNodeById(this.m_A, val.MetadataSchemaId);
						if (schemaNodeById != null)
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
							if (schemaNodeById.ChildNodes.Count > 0)
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
								this.A(val, schemaNodeById);
							}
						}
						schemaNodeById = null;
					}
				}
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
				try
				{
					ContentFilters = new ObservableCollection<FiltersGroup>();
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					BindingOperations.ClearBinding(icFilters, ItemsControl.ItemsSourceProperty);
					icFilters.SetBinding(ItemsControl.ItemsSourceProperty, VH.A(87701));
					ProjectData.ClearProjectError();
				}
			}
			IEnumerable<FiltersGroup> enumerable = from A in (IEnumerable<ContentGroup>)AllGroups
				select A.FiltersGroup into A
				where A != null
				select A;
			if (ContentFilters.SequenceEqual(enumerable))
			{
				while (true)
				{
					switch (4)
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
				icFilters.SetBinding(ItemsControl.ItemsSourceProperty, VH.A(87701));
				ProjectData.ClearProjectError();
			}
			using IEnumerator<FiltersGroup> enumerator = enumerable.GetEnumerator();
			while (enumerator.MoveNext())
			{
				FiltersGroup current = enumerator.Current;
				ContentFilters.Add(current);
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
			((System.Windows.Window)(object)DragDropOverlay).Close();
			DragDropOverlay = null;
		}
		PowerPointApp.CleanUp(ppApp, HiddenFiles, blnQuitPP);
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
		//IL_0096: Expected I4, but got Unknown
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
					switch (7)
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
						case 3:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)A);
							break;
						case 0:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)B);
							break;
						case 1:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)C);
							break;
						case 2:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)D);
							break;
						case 4:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)E);
							break;
						case 12:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)F);
							break;
						case 10:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)G);
							flag = true;
							break;
						case 5:
						case 6:
						case 7:
						case 8:
						case 9:
						case 11:
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
							switch (3)
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
						val2.ApplyItemsCriteria();
					}
					if (((Collection<ContentItem>)(object)val2.AllContentItems).Count > 10)
					{
						val2.ExpandContent();
						SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
						A().RunWorkerAsync(val);
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
				Monitor.Exit(allContentItemsLock);
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
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Expected O, but got Unknown
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
			switch (6)
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
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		((ContentGroup)((Popup)sender).DataContext).IsPopupOpen = true;
	}

	private void MorePopupClosed(object sender, EventArgs e)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		((ContentGroup)((Popup)sender).DataContext).IsPopupOpen = false;
	}

	private void HideContentGroup(object sender, RoutedEventArgs e)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Expected O, but got Unknown
		ContentGroup val = (ContentGroup)((System.Windows.Controls.Button)sender).DataContext;
		PopupMenus.HideGroup(ref val);
		val = null;
	}

	private ContentItem A(ContentGroup A, XmlNode B)
	{
		return (ContentItem)(object)new TableItem(A, B);
	}

	private ContentItem B(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new ShapeItem(A, B);
	}

	private ContentItem C(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new ImageItem(A, B);
	}

	private ContentItem D(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new ChartItem(A, B);
	}

	private ContentItem E(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new TextItem(A, B);
	}

	private ContentItem F(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new PdfItem(A, B);
	}

	private ContentItem G(ContentGroup A, XmlNode B)
	{
		return (ContentItem)(object)new ModelItem(A, B);
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
		//IL_00c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cf: Invalid comparison between Unknown and I4
		System.Windows.Controls.ListViewItem listViewItem = (System.Windows.Controls.ListViewItem)sender;
		ContentItem val = (ContentItem)listViewItem.DataContext;
		System.Windows.Controls.ListView listView = (System.Windows.Controls.ListView)ItemsControl.ItemsControlFromItemContainer(listViewItem);
		Base.ScrollToItem(scroller, listViewItem);
		if (ActiveListView != null)
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
			if (listView != ActiveListView)
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
				if (ActiveListView.SelectedItems.Count > 0)
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
					ActiveListView.SelectedItems.Clear();
				}
			}
		}
		ActiveListView = listView;
		A();
		if ((int)val.Group.ContentType != 13)
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
			if ((int)val.Group.ContentType != 15)
			{
				SelActionText = VH.A(57289);
				goto IL_00f7;
			}
		}
		SelActionText = VH.A(88071);
		goto IL_00f7;
		IL_00f7:
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
		global::A.K.Settings.LibraryPaneShowStars = value;
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
	}

	private void ImageTypeBadgeToggle(object sender, RoutedEventArgs e)
	{
		global::A.K.Settings.LibraryPaneShowImageTypeBadge = chkImageTypeBadge.IsChecked.Value;
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
				predicate = (_Closure_0024__.A = [SpecialName] (ContentItem A) => A.Visibility != System.Windows.Visibility.Visible);
			}
			else
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
					switch (2)
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
			global::A.K.Settings.ContentInsertShowPersonal = A.Value;
		}
		if (B.HasValue)
		{
			global::A.K.Settings.ContentInsertShowShared = B.Value;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			global::A.K.Settings.ContentInsertShowPublic = C.Value;
		}
		if (D.HasValue)
		{
			global::A.K.Settings.ContentInsertShow3rdParty = D.Value;
		}
		this.A(A: false);
	}

	private void A(ContentType A, bool B)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_003e: Expected I4, but got Unknown
		switch (A - 3)
		{
		case 3:
			global::A.K.Settings.ContentInsertShowTables = B;
			break;
		case 0:
			global::A.K.Settings.ContentInsertShowShapes = B;
			break;
		case 1:
			global::A.K.Settings.ContentInsertShowImages = B;
			break;
		case 2:
			global::A.K.Settings.ContentInsertShowCharts = B;
			break;
		case 4:
			global::A.K.Settings.ContentInsertShowText = B;
			break;
		case 12:
			if (!Pane.PDFsContentIsEnabled)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			global::A.K.Settings.ContentInsertShowPDFs = B;
			break;
		case 10:
			if (!Pane.ModelsContentIsEnabled)
			{
				break;
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
			global::A.K.Settings.ContentInsertShowModels = B;
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
		ContentGroupsCollection allGroups = AllGroups;
		object obj;
		if (allGroups == null)
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
		else
		{
			obj = allGroups.ImgTypesFilter;
		}
		if (obj != null)
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
			global::A.K.Settings.ContentInsertExcludeImageTypes = AllGroups.ImgTypesFilter.ToListString();
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
			switch (7)
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
				switch (1)
				{
				case 0:
					continue;
				}
				if (A)
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
							RangeObservableCollection<ContentItem> allContentItems = val.AllContentItems;
							if (allContentItems == null)
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
										return 0;
									}
								}
							}
							return ((Collection<ContentItem>)(object)allContentItems).Count;
						});
					}
					else
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
						selector = _Closure_0024__.A;
					}
					if (((IEnumerable<ContentGroup>)allGroups).Sum(selector) > 100)
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
			this.A(A: false);
		}
		else
		{
			this.A(A, (bool?)null);
		}
	}

	private void A(ContentGroup A, bool? B = null)
	{
		if (A.FiltersGroup == null)
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
			bool num;
			if (!B.HasValue)
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
					switch (3)
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
			System.Windows.Visibility visibility = (System.Windows.Visibility)num2;
			if (A.FiltersGroup.GroupVisibility == visibility)
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
				switch (3)
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
				switch (3)
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
					switch (5)
					{
					case 0:
						break;
					default:
						return false;
					}
				}
			}
		}
		return Filter.ApplyLibraryFilter(A, global::A.K.Settings.ContentInsertShowShared, global::A.K.Settings.ContentInsertShowPersonal, global::A.K.Settings.ContentInsertShowPublic, global::A.K.Settings.ContentInsertShow3rdParty, (Func<ContentGroup, bool>)C);
	}

	private bool C(ContentGroup A)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0009: Invalid comparison between Unknown and I4
		//IL_0040: Unknown result type (might be due to invalid IL or missing references)
		//IL_0045: Unknown result type (might be due to invalid IL or missing references)
		//IL_0048: Invalid comparison between Unknown and I4
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		//IL_007b: Unknown result type (might be due to invalid IL or missing references)
		//IL_007e: Invalid comparison between Unknown and I4
		//IL_00a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00aa: Invalid comparison between Unknown and I4
		//IL_00d8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00de: Invalid comparison between Unknown and I4
		//IL_010a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0111: Invalid comparison between Unknown and I4
		//IL_0141: Unknown result type (might be due to invalid IL or missing references)
		//IL_0146: Unknown result type (might be due to invalid IL or missing references)
		//IL_014a: Invalid comparison between Unknown and I4
		if ((int)A.ContentType == 6)
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
			if (global::A.K.Settings.ContentInsertShowTables)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						return true;
					}
				}
			}
		}
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
			if (global::A.K.Settings.ContentInsertShowShapes)
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
		if ((int)A.ContentType == 4 && global::A.K.Settings.ContentInsertShowImages)
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
		if ((int)A.ContentType == 5)
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
			if (global::A.K.Settings.ContentInsertShowCharts)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						return true;
					}
				}
			}
		}
		if ((int)A.ContentType == 7)
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
			if (global::A.K.Settings.ContentInsertShowText)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						return true;
					}
				}
			}
		}
		if ((int)A.ContentType == 15)
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
			if (Pane.PDFsContentIsEnabled)
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
				if (global::A.K.Settings.ContentInsertShowPDFs)
				{
					return true;
				}
			}
		}
		if ((int)A.ContentType == 13)
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
			if (Pane.ModelsContentIsEnabled)
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
				if (global::A.K.Settings.ContentInsertShowModels)
				{
					return true;
				}
			}
		}
		return false;
	}

	private void ContentFiltersChecked(object sender, RoutedEventArgs e)
	{
		grdContentFilters.Visibility = System.Windows.Visibility.Visible;
	}

	private void ContentFiltersUnchecked(object sender, RoutedEventArgs e)
	{
		grdContentFilters.Visibility = System.Windows.Visibility.Collapsed;
	}

	private void ImageTypeFiltersChecked(object sender, RoutedEventArgs e)
	{
		grdImageTypeFilters.Visibility = System.Windows.Visibility.Visible;
	}

	private void ImageTypeFiltersUnchecked(object sender, RoutedEventArgs e)
	{
		grdImageTypeFilters.Visibility = System.Windows.Visibility.Collapsed;
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
					switch (7)
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
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		System.Windows.Controls.ComboBox comboBox = (System.Windows.Controls.ComboBox)sender;
		if (comboBox.SelectedIndex == 0)
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
			if (B)
			{
				C();
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
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0014: Expected O, but got Unknown
		System.Windows.Controls.ListViewItem obj = (System.Windows.Controls.ListViewItem)sender;
		ContentItem val = (ContentItem)obj.DataContext;
		System.Windows.Controls.ContextMenu contextMenu = obj.ContextMenu;
		D(val.Group);
		this.m_B = true;
		Base.ShowHideContextMenuOptionForContentTypes(contextMenu, val, VH.A(88154), true, (ContentType[])(object)new ContentType[1] { (ContentType)4 });
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
		global::A.K.Settings.LibraryPaneKeepSourceFormat = chkKeep.IsChecked.Value;
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
		//IL_0047: Unknown result type (might be due to invalid IL or missing references)
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_0056: Unknown result type (might be due to invalid IL or missing references)
		//IL_0090: Expected I4, but got Unknown
		if (ActiveListView == null || ActiveListView.SelectedItems.Count <= 0)
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
			ContentType contentType = ((ContentGroup)ActiveListView.DataContext).ContentType;
			switch (contentType - 3)
			{
			case 2:
				I();
				break;
			case 3:
				J();
				break;
			case 0:
				K();
				break;
			case 1:
				L();
				break;
			case 4:
				M();
				break;
			case 12:
				N();
				break;
			case 10:
				O();
				break;
			}
			UsageLogger.LogInsertion(ActiveListView.SelectedItems.Cast<ContentItem>(), (OfficeApp)1);
			return;
		}
	}

	private void B(string A)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)6, A);
	}

	private void I()
	{
		//IL_0132: Unknown result type (might be due to invalid IL or missing references)
		//IL_013c: Expected O, but got Unknown
		//IL_01d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_01de: Expected O, but got Unknown
		//IL_01e8: Unknown result type (might be due to invalid IL or missing references)
		if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		IEnumerator enumerator2 = default(IEnumerator);
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
			Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
			Microsoft.Office.Interop.Excel.Workbook workbook = null;
			Microsoft.Office.Interop.Excel.Workbook workbook2 = null;
			bool value = chkKeep.IsChecked.Value;
			string c = Content.RightNow();
			string d = A();
			try
			{
				workbook2 = application.ActiveWorkbook;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			Worksheet worksheet;
			if (workbook2 != null)
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
				if (Miscellaneous.DisplayObjects(workbook2))
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
					if (!Workbooks.IsShared(workbook2, true, (System.Windows.Window)null))
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
						application.ScreenUpdating = false;
						bool copyObjectsWithCells = application.CopyObjectsWithCells;
						application.CopyObjectsWithCells = true;
						try
						{
							int num = Conversions.ToInteger(NewLateBinding.LateGet(workbook2.ActiveSheet, null, VH.A(48135), new object[0], null, null, null));
							if (workbook2.Path.Length == 0)
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
								workbook2.Saved = false;
							}
							workbook = A((ContentItem)ActiveListView.SelectedItems[0]);
							if (Operators.ConditionalCompareObjectLess(((Worksheet)workbook2.Worksheets[1]).Columns.CountLarge, ((Worksheet)workbook.Worksheets[1]).Columns.CountLarge, TextCompare: false))
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
								D(VH.A(88177));
							}
							else
							{
								try
								{
									enumerator = ActiveListView.SelectedItems.GetEnumerator();
									while (enumerator.MoveNext())
									{
										ContentItem val = (ContentItem)enumerator.Current;
										worksheet = (Worksheet)workbook.Sheets[((ChartItem)val).SheetIndex];
										worksheet.Copy(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(workbook2.Sheets[num]));
										ChartObject chartObject = (ChartObject)((Worksheet)workbook2.ActiveSheet).ChartObjects(1);
										num = checked(num + 1);
										if (value)
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
											Charts.RetainChartTheme(((ChartObject)worksheet.ChartObjects(1)).Chart, chartObject.Chart);
										}
										chartObject.Name = Base.RandomChartName();
										try
										{
											enumerator2 = ((Worksheet)workbook2.ActiveSheet).Shapes.GetEnumerator();
											while (true)
											{
												if (enumerator2.MoveNext())
												{
													Microsoft.Office.Interop.Excel.Shape shape = (Microsoft.Office.Interop.Excel.Shape)enumerator2.Current;
													if (shape.HasChart != MsoTriState.msoTrue || shape.Chart != chartObject.Chart)
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
														A(shape, val, c, d);
														break;
													}
													break;
												}
												while (true)
												{
													switch (2)
													{
													case 0:
														break;
													default:
														goto end_IL_0308;
													}
													continue;
													end_IL_0308:
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
										chartObject = null;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_0345;
										}
										continue;
										end_IL_0345:
										break;
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
							B(VH.A(88611));
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							C(VH.A(88654) + ex4.Message);
							clsReporting.LogException(ex4);
							ProjectData.ClearProjectError();
						}
						application.CopyObjectsWithCells = copyObjectsWithCells;
						if (workbook != null)
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
							application.DisplayAlerts = false;
							try
							{
								workbook.Close(false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
							}
							catch (Exception ex5)
							{
								ProjectData.SetProjectError(ex5);
								Exception ex6 = ex5;
								ProjectData.ClearProjectError();
							}
							application.DisplayAlerts = true;
						}
						application.ScreenUpdating = true;
					}
				}
			}
			application = null;
			workbook2 = null;
			workbook = null;
			worksheet = null;
			return;
		}
	}

	private void J()
	{
		VE vE = null;
		try
		{
			if (!Access.AllowExcelOperation((PlanType)4, (Restriction)1, false))
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
						return;
					}
				}
			}
			vE = new VE(MH.A.Application, A);
			if (!vE.A(ActiveListView.SelectedItems.Cast<ContentItem>().ToList()))
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
				B(VH.A(88749));
				return;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			C(string.Format(VH.A(88792), ex2.Message));
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		finally
		{
			vE?.A();
		}
	}

	private void K()
	{
		//IL_0102: Unknown result type (might be due to invalid IL or missing references)
		//IL_0136: Unknown result type (might be due to invalid IL or missing references)
		//IL_013d: Expected O, but got Unknown
		//IL_017f: Unknown result type (might be due to invalid IL or missing references)
		//IL_018a: Expected O, but got Unknown
		//IL_014c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0178: Expected O, but got Unknown
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		Presentation presentation = null;
		string c = Content.RightNow();
		string d = A();
		Microsoft.Office.Interop.PowerPoint.Shape sourceShape;
		Microsoft.Office.Interop.Excel.Shape a;
		if (Miscellaneous.DisplayObjects(application.ActiveWorkbook) && !Workbooks.IsShared(application.ActiveWorkbook, true, (System.Windows.Window)null))
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
			if (application.ActiveWindow.SelectedSheets.Count == 1)
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
				if (application.ActiveSheet is Worksheet)
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
					Worksheet worksheet = (Worksheet)application.ActiveSheet;
					ExcelAddIn1.Sheets.Protection.Unprotect(worksheet);
					if (!worksheet.ProtectContents)
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
						if (Insert.AreShapesLegacyPublished(ActiveListView))
						{
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
											switch (1)
											{
											case 0:
												continue;
											}
											break;
										}
										sourceShape = PowerPointApp.GetSourceShape((ShapeItem)val, PowerPointApp.GetPresentation(ref ppApp, ref HiddenFiles, ref blnQuitPP, val.ContentPath));
									}
									else
									{
										sourceShape = PowerPointApp.GetSourceShape((ShapeItem)val, presentation);
									}
									a = ExcelAddIn1.Library2.Insert.Shapes.A(sourceShape, worksheet);
									A(a, val, c, d);
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
							clsClipboard.ClearClipboard();
							B(VH.A(88889));
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							C(ex2.Message);
							clsReporting.LogException(ex2);
							ProjectData.ClearProjectError();
						}
						presentation = null;
					}
					else
					{
						D(VH.A(88932));
					}
					worksheet = null;
				}
			}
			else
			{
				D(VH.A(89023));
			}
		}
		application = null;
		sourceShape = null;
		a = null;
	}

	private void L()
	{
		//IL_0297: Unknown result type (might be due to invalid IL or missing references)
		//IL_029e: Expected O, but got Unknown
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		string c = Content.RightNow();
		string d = A();
		Microsoft.Office.Interop.Excel.Shape shape;
		if (Miscellaneous.DisplayObjects(application.ActiveWorkbook))
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
			if (application.ActiveWindow.SelectedSheets.Count == 1)
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
				if (application.ActiveSheet is Worksheet)
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
					Worksheet obj = (Worksheet)application.ActiveSheet;
					ExcelAddIn1.Sheets.Protection.Unprotect(obj);
					bool flag = false;
					bool flag2 = false;
					float num = 0f;
					float num2 = 0f;
					if (!obj.ProtectContents)
					{
						float height = default(float);
						float width = default(float);
						if (application.Selection is Range)
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
							Range activeCell = application.ActiveCell;
							num = Conversions.ToSingle(activeCell.Top);
							num2 = Conversions.ToSingle(activeCell.Left);
							_ = null;
							Range range = (Range)application.Selection;
							if (!Operators.ConditionalCompareObjectEqual(range.Cells.CountLarge, 1, TextCompare: false) && Operators.ConditionalCompareObjectLess(range.Rows.CountLarge, range.Worksheet.Rows.CountLarge, TextCompare: false))
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
								if (Operators.ConditionalCompareObjectLess(range.Columns.CountLarge, range.Worksheet.Columns.CountLarge, TextCompare: false))
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
									if (Operators.ConditionalCompareObjectEqual(range.Columns.CountLarge, 1, TextCompare: false))
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
										height = Conversions.ToSingle(range.Height);
										flag = true;
									}
									else if (Operators.ConditionalCompareObjectEqual(range.Rows.CountLarge, 1, TextCompare: false))
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
										width = Conversions.ToSingle(range.Width);
										flag2 = true;
									}
								}
							}
							range = null;
						}
						else if (application.ActiveChart != null)
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
							Range topLeftCell = ((ChartObject)application.ActiveChart.Parent).TopLeftCell;
							num = Conversions.ToSingle(topLeftCell.Top);
							num2 = Conversions.ToSingle(topLeftCell.Left);
							_ = null;
						}
						application.ScreenUpdating = false;
						try
						{
							IEnumerator enumerator = default(IEnumerator);
							try
							{
								ContentItem val;
								for (enumerator = ActiveListView.SelectedItems.GetEnumerator(); enumerator.MoveNext(); A(shape, val, c, d), num += 10f, num2 += 10f)
								{
									val = (ContentItem)enumerator.Current;
									shape = ((Worksheet)application.ActiveSheet).Shapes.AddPicture2(val.ContentPath, MsoTriState.msoFalse, MsoTriState.msoTrue, num2, num, -1f, -1f, MsoPictureCompress.msoPictureCompressDocDefault);
									shape.LockAspectRatio = MsoTriState.msoTrue;
									if (flag && !flag2)
									{
										shape.Height = height;
										continue;
									}
									if (flag2)
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
										if (!flag)
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
											shape.Width = width;
											continue;
										}
									}
									if (!flag)
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
									if (!flag2)
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
									shape.LockAspectRatio = MsoTriState.msoFalse;
									shape.Height = height;
									shape.Width = width;
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
							B(VH.A(89134));
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							C(ex2.Message);
							clsReporting.LogException(ex2);
							ProjectData.ClearProjectError();
						}
						application.ScreenUpdating = true;
					}
					else
					{
						D(VH.A(89177));
					}
				}
			}
			else
			{
				D(VH.A(89272));
			}
		}
		application = null;
		shape = null;
	}

	private void M()
	{
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		Microsoft.Office.Interop.Excel.Application application = MH.A.Application;
		int num = 0;
		Range range;
		try
		{
			if (application.Selection is Range)
			{
				range = (Range)application.Selection;
				bool flag = JH.A(range);
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = ActiveListView.SelectedItems.GetEnumerator();
					while (enumerator.MoveNext())
					{
						string value = File.ReadAllText(((ContentItem)enumerator.Current).ContentPath);
						((Range)range.Cells[1, 1]).get_Offset((object)num, RuntimeHelpers.GetObjectValue(Missing.Value)).Value2 = value;
						num = checked(num + 1);
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							(enumerator as IDisposable).Dispose();
							break;
						}
					}
				}
				if (flag)
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
					JH.A(range, VH.A(89387));
				}
			}
			B(VH.A(89410));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			C(ex2.Message);
			ProjectData.ClearProjectError();
		}
		application = null;
		range = null;
	}

	private void N()
	{
		B(VH.A(89449));
	}

	private void O()
	{
		//IL_008d: Unknown result type (might be due to invalid IL or missing references)
		if (ActiveListView.SelectedItems.Count > 5)
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
			if (!UIFormsExtensions.AskOkCancel((System.Windows.Window)null, string.Format(VH.A(89466), ActiveListView.SelectedItems.Count)))
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
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ActiveListView.SelectedItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Models.OpenFile(((ContentItem)enumerator.Current).ContentPath);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_00a6;
				}
				continue;
				end_IL_00a6:
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
		B(VH.A(89575));
	}

	private Microsoft.Office.Interop.Excel.Workbook A(ContentItem A)
	{
		Microsoft.Office.Interop.Excel.Workbook workbook = null;
		string contentPath = A.ContentPath;
		try
		{
			workbook = MH.A.Application.Workbooks[Path.GetFileName(contentPath)];
			if (Operators.CompareString(workbook.FullName, contentPath, TextCompare: false) != 0)
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
					workbook = null;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (workbook == null)
		{
			try
			{
				workbook = MH.A.Application.Workbooks.Open(contentPath, false, true, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				if (workbook != null)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						HiddenFiles.Add(contentPath);
						break;
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		return workbook;
	}

	private void P()
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
		//IL_004d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0057: Expected O, but got Unknown
		if (DragDropOverlay != null)
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
			try
			{
				DragDropOverlay = new wpfDragDrop((Action)P, (IntPtr)MH.A.Application.ActiveWindow.Hwnd);
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

	private void A(Microsoft.Office.Interop.Excel.Shape A, ContentItem B, string C, string D)
	{
		Tagging.A(A, B, C, D);
	}

	private string A()
	{
		return Core.GetAuthor(MH.A.Application.ActiveWorkbook);
	}

	private bool D(ContentGroup A)
	{
		return Access.UserHasAccess(A.Library, (AccessType)1, IsAdmin);
	}

	private void Q()
	{
		D(VH.A(89614));
	}

	private void C(string A)
	{
		Forms.ErrorMessage(A);
	}

	private void D(string A)
	{
		Forms.WarningMessage(A);
	}

	private ContentGroup A(ContentType A)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Unknown result type (might be due to invalid IL or missing references)
		return ((IEnumerable<ContentGroup>)AllGroups).First([SpecialName] (ContentGroup val) => val.ContentType == A);
	}

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (this.m_K)
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
			this.m_K = true;
			Uri resourceLocator = new Uri(VH.A(89750), UriKind.Relative);
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
		if (connectionId == 13)
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
					((System.Windows.Controls.MenuItem)target).Click += InsertContent;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (2)
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
				switch (2)
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
				switch (4)
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
			chkPersonal = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 19)
		{
			while (true)
			{
				switch (3)
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
				switch (6)
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
				switch (5)
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
				switch (5)
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
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					grdContentFilters = (DockPanel)target;
					return;
				}
			}
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
					chkCharts = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkTables = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
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
					chkShapes = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 29)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkImages = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkText = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
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
					chkPdfs = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 32)
		{
			chkModels = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 33)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkFavorites = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 34)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkImageTypeFilters = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 35)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					grdImageTypeFilters = (StackPanel)target;
					return;
				}
			}
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
					chkImagesSvg = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 37)
		{
			chkImagesPng = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 38)
		{
			chkImagesJpg = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 39)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkImagesEmf = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 40)
		{
			chkImagesWmf = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 41)
		{
			chkImagesGif = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 42)
		{
			chkImagesBmp = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 43)
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
		if (connectionId == 44)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					icFilters = (ItemsControl)target;
					return;
				}
			}
		}
		if (connectionId == 48)
		{
			icContent = (ItemsControl)target;
			return;
		}
		if (connectionId == 52)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkPreview = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 53)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkStars = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 54)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkImageTypeBadge = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 55)
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
		if (connectionId == 56)
		{
			while (true)
			{
				switch (1)
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
		case 57:
			popRemoveGroup = (Popup)target;
			popRemoveGroup.Opened += MorePopupOpened;
			popRemoveGroup.Closed += MorePopupClosed;
			popRemoveGroup.PreviewKeyDown += CloseMorePopup;
			break;
		case 58:
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				btnHideContent = (System.Windows.Controls.Button)target;
				btnHideContent.Click += HideContentGroup;
				return;
			}
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

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
		{
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
			((System.Windows.Controls.TextBox)target).TextChanged += FilterTextChanged;
		}
		if (connectionId == 4)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += FilterListSelectionChanged;
		}
		if (connectionId == 5)
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
			((System.Windows.Controls.CheckBox)target).Checked += ListItemCheckChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += ListItemCheckChanged;
		}
		if (connectionId == 6)
		{
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
				switch (3)
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
				switch (6)
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
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				break;
			}
			((System.Windows.Controls.TextBox)target).TextChanged += FilterMaxValueChanged;
		}
		if (connectionId == 11)
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
		if (connectionId == 12)
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
			((System.Windows.Controls.TextBox)target).TextChanged += FilterMaxValueChanged;
		}
		if (connectionId == 45)
		{
			((System.Windows.Controls.ComboBox)target).SelectionChanged += SortFieldChanged;
		}
		if (connectionId == 46)
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
			((TextBlock)target).MouseUp += SortOrderChanged;
		}
		if (connectionId == 47)
		{
			((TextBlock)target).MouseUp += ResetFilters;
		}
		if (connectionId == 49)
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
			((System.Windows.Controls.CheckBox)target).Checked += ExpandGroup;
		}
		if (connectionId == 50)
		{
			((System.Windows.Controls.ListView)target).PreviewMouseLeftButtonDown += lstView_PreviewMouseLeftButtonDown;
			((System.Windows.Controls.ListView)target).MouseMove += lstView_MouseMove;
			((System.Windows.Controls.ListView)target).DragLeave += lstView_DragLeave;
			((System.Windows.Controls.ListView)target).SelectionChanged += lstView_SelectionChanged;
		}
		if (connectionId != 51)
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
			EventSetter eventSetter = new EventSetter();
			eventSetter.Event = FrameworkElement.ContextMenuOpeningEvent;
			eventSetter.Handler = new ContextMenuEventHandler(MenuOpening);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = FrameworkElement.ContextMenuClosingEvent;
			eventSetter.Handler = new ContextMenuEventHandler(MenuClosing);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = ListBoxItem.SelectedEvent;
			eventSetter.Handler = new RoutedEventHandler(ListViewItemSelected);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseEnterEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(MouseEnterListViewItem);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = UIElement.MouseLeaveEvent;
			eventSetter.Handler = new System.Windows.Input.MouseEventHandler(MouseLeaveListViewItem);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = System.Windows.Controls.Control.MouseDoubleClickEvent;
			eventSetter.Handler = new MouseButtonEventHandler(MouseDblClickListViewItem);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			eventSetter = new EventSetter();
			eventSetter.Event = FrameworkElement.RequestBringIntoViewEvent;
			eventSetter.Handler = new RequestBringIntoViewEventHandler(OnRequestBringIntoView);
			((System.Windows.Style)target).Setters.Add(eventSetter);
			return;
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
}
