using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
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
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Explorer;
using PowerPointAddIn1.Library2.Insert;
using PowerPointAddIn1.Library2.Versioning;
using PowerPointAddIn1.Links;
using PowerPointAddIn1.Shapes;
using PowerPointAddIn1.Shapes.Arrange;
using PowerPointAddIn1.Slides;

namespace PowerPointAddIn1.Library2.UI;

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

		public static Func<Slide, int> A;

		public static Func<Microsoft.Office.Interop.PowerPoint.Shape, string> A;

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
			return ((Collection<ContentItem>)(object)A.AllContentItems)?.Count ?? 0;
		}

		[SpecialName]
		internal int A(Slide A)
		{
			return A.SlideIndex;
		}

		[SpecialName]
		internal string A(Microsoft.Office.Interop.PowerPoint.Shape A)
		{
			return A.Name;
		}
	}

	[CompilerGenerated]
	internal sealed class QD
	{
		public int A;

		[SpecialName]
		internal void A()
		{
			this.A = 1;
		}
	}

	[CompilerGenerated]
	internal sealed class RD
	{
		public Microsoft.Office.Interop.Excel.Chart A;

		public Action A;

		public Action B;

		public RD(RD A)
		{
			if (A == null)
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.ChartArea.Copy();
		}

		[SpecialName]
		internal void B()
		{
			this.A.ChartArea.Copy();
		}
	}

	[CompilerGenerated]
	internal sealed class SD
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public Action A;

		public SD(SD A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.Copy();
		}
	}

	[CompilerGenerated]
	internal sealed class TD
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public Action A;

		public TD(TD A)
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
			this.A.Copy();
		}
	}

	[CompilerGenerated]
	internal sealed class UD
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

	public Microsoft.Office.Interop.Excel.Application xlApp;

	public bool blnQuitXL;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

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

	private System.Windows.Controls.UserControl m_A;

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

	[AccessedThroughProperty("chkPublic")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("scroller")]
	private ScrollViewer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilters")]
	private System.Windows.Controls.CheckBox m_E;

	[AccessedThroughProperty("stkFilters")]
	[CompilerGenerated]
	private StackPanel m_A;

	[AccessedThroughProperty("chkContentFilters")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("grdContentFilters")]
	private DockPanel m_A;

	[AccessedThroughProperty("chkSlides")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_G;

	[AccessedThroughProperty("chkShapes")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_H;

	[AccessedThroughProperty("chkImages")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_I;

	[AccessedThroughProperty("chkVideos")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_J;

	[AccessedThroughProperty("chkCharts")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_K;

	[AccessedThroughProperty("chkText")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_L;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPdfs")]
	private System.Windows.Controls.CheckBox m_M;

	[CompilerGenerated]
	[AccessedThroughProperty("chkDecks")]
	private System.Windows.Controls.CheckBox m_N;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFavorites")]
	private System.Windows.Controls.CheckBox m_O;

	[AccessedThroughProperty("chkImageTypeFilters")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_P;

	[AccessedThroughProperty("grdImageTypeFilters")]
	[CompilerGenerated]
	private StackPanel m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesSvg")]
	private System.Windows.Controls.CheckBox m_Q;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesPng")]
	private System.Windows.Controls.CheckBox m_R;

	[AccessedThroughProperty("chkImagesJpg")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_S;

	[AccessedThroughProperty("chkImagesEmf")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_T;

	[AccessedThroughProperty("chkImagesWmf")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_U;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesGif")]
	private System.Windows.Controls.CheckBox V;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesBmp")]
	private System.Windows.Controls.CheckBox W;

	[CompilerGenerated]
	[AccessedThroughProperty("chkImagesTiff")]
	private System.Windows.Controls.CheckBox X;

	[CompilerGenerated]
	[AccessedThroughProperty("icFilters")]
	private ItemsControl m_A;

	[AccessedThroughProperty("icContent")]
	[CompilerGenerated]
	private ItemsControl m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPreview")]
	private System.Windows.Controls.CheckBox Y;

	[AccessedThroughProperty("chkStars")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox Z;

	[AccessedThroughProperty("chkImageTypeBadge")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox AB;

	[CompilerGenerated]
	[AccessedThroughProperty("chkArrange")]
	private System.Windows.Controls.CheckBox BB;

	[CompilerGenerated]
	[AccessedThroughProperty("btnInsert")]
	private System.Windows.Controls.Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkKeep")]
	private System.Windows.Controls.CheckBox CB;

	[CompilerGenerated]
	[AccessedThroughProperty("popRemoveGroup")]
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
			A(AH.A(10961));
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
			A(AH.A(64956));
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
			A(AH.A(64975));
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
			A(AH.A(65004));
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
				switch (2)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.m_A = value;
				A(AH.A(65033));
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
			if (object.Equals(this.m_B, value))
			{
				this.m_B = value;
				A(AH.A(65073));
			}
		}
	}

	public System.Windows.Controls.UserControl ArrangeView
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(10994));
		}
	}

	private FrameworkElement PreviewParentUIElem => this;

	private bool PreviewSetting
	{
		get
		{
			return PB.Settings.LibraryPaneShowPreview;
		}
		set
		{
			PB.Settings.LibraryPaneShowPreview = value;
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
			A(AH.A(65447));
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
			A(AH.A(65484));
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
				switch (2)
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

	internal virtual System.Windows.Controls.CheckBox chkSlides
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

	internal virtual System.Windows.Controls.CheckBox chkVideos
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

	internal virtual System.Windows.Controls.CheckBox chkCharts
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

	internal virtual System.Windows.Controls.CheckBox chkText
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

	internal virtual System.Windows.Controls.CheckBox chkPdfs
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

	internal virtual System.Windows.Controls.CheckBox chkDecks
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

	internal virtual System.Windows.Controls.CheckBox chkFavorites
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

	internal virtual System.Windows.Controls.CheckBox chkImageTypeFilters
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
			RoutedEventHandler value2 = ImageTypeFiltersChecked;
			RoutedEventHandler value3 = ImageTypeFiltersUnchecked;
			System.Windows.Controls.CheckBox checkBox = this.m_P;
			if (checkBox != null)
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
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_P = value;
			checkBox = this.m_P;
			if (checkBox == null)
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
			return this.m_Q;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_Q = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesPng
	{
		[CompilerGenerated]
		get
		{
			return this.m_R;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_R = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesJpg
	{
		[CompilerGenerated]
		get
		{
			return this.m_S;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_S = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesEmf
	{
		[CompilerGenerated]
		get
		{
			return this.m_T;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_T = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesWmf
	{
		[CompilerGenerated]
		get
		{
			return this.m_U;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_U = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkImagesGif
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

	internal virtual System.Windows.Controls.CheckBox chkImagesBmp
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

	internal virtual System.Windows.Controls.CheckBox chkImagesTiff
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
			return Y;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			Y = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkStars
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

	internal virtual System.Windows.Controls.CheckBox chkImageTypeBadge
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

	internal virtual System.Windows.Controls.CheckBox chkArrange
	{
		[CompilerGenerated]
		get
		{
			return BB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			BB = value;
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
				switch (5)
				{
				case 0:
					continue;
				}
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkKeep
	{
		[CompilerGenerated]
		get
		{
			return CB;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			CB = value;
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
				switch (4)
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
		//IL_00ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b8: Expected O, but got Unknown
		//IL_00e4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fa: Expected O, but got Unknown
		//IL_00f5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ff: Expected O, but got Unknown
		//IL_012d: Unknown result type (might be due to invalid IL or missing references)
		//IL_013e: Expected O, but got Unknown
		//IL_0145: Unknown result type (might be due to invalid IL or missing references)
		//IL_014f: Expected O, but got Unknown
		base.Loaded += wpfLibrary_Loaded;
		base.Unloaded += wpfLibrary_Unloaded;
		blnQuitXL = false;
		this.m_B = false;
		this.m_A = null;
		this.m_A = "";
		this.m_B = AH.A(65425);
		this.m_A = null;
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
			switch (3)
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
			obj = null;
		}
		else
		{
			IList selectedItems = activeListView.SelectedItems;
			if (selectedItems == null)
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
				obj = null;
			}
			else
			{
				obj = selectedItems.Count;
			}
		}
		int? num = obj;
		int valueOrDefault = num.GetValueOrDefault();
		object selectedCountStr;
		if (valueOrDefault >= 2)
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
			selectedCountStr = string.Format(AH.A(65066), valueOrDefault);
		}
		else
		{
			selectedCountStr = "";
		}
		SelectedCountStr = (string)selectedCountStr;
	}

	private void wpfLibrary_Loaded(object sender, RoutedEventArgs e)
	{
		this.m_A.OnLoad();
		System.Windows.Controls.CheckBox checkBox = chkPreview;
		checkBox.IsChecked = PB.Settings.LibraryPaneShowPreview;
		checkBox.Checked += PreviewToggle;
		checkBox.Unchecked += PreviewToggle;
		_ = null;
		System.Windows.Controls.CheckBox checkBox2 = chkStars;
		checkBox2.IsChecked = PB.Settings.LibraryPaneShowStars;
		checkBox2.Checked += StarsToggle;
		checkBox2.Unchecked += StarsToggle;
		_ = null;
		System.Windows.Controls.CheckBox checkBox3 = chkImageTypeBadge;
		checkBox3.IsChecked = PB.Settings.LibraryPaneShowImageTypeBadge;
		checkBox3.Checked += ImageTypeBadgeToggle;
		checkBox3.Unchecked += ImageTypeBadgeToggle;
		_ = null;
		System.Windows.Controls.CheckBox checkBox4 = chkArrange;
		checkBox4.IsChecked = PB.Settings.LibraryPaneOfferArrange;
		checkBox4.Checked += ArrangeToggle;
		checkBox4.Unchecked += ArrangeToggle;
		_ = null;
		System.Windows.Controls.CheckBox checkBox5 = chkKeep;
		checkBox5.IsChecked = PB.Settings.LibraryPaneKeepSourceFormat;
		checkBox5.Checked += KeepSourceFormatToggle;
		checkBox5.Unchecked += KeepSourceFormatToggle;
		_ = null;
	}

	private void wpfLibrary_Unloaded(object sender, RoutedEventArgs e)
	{
		//IL_003d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0047: Expected O, but got Unknown
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Expected O, but got Unknown
		this.m_G = true;
		CleanUp();
		Pane.A(this);
		if (AllGroups != null)
		{
			using IEnumerator<ContentGroup> enumerator = ((Collection<ContentGroup>)(object)AllGroups).GetEnumerator();
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
		if (GroupsListener != null)
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
			GroupsListener.Disconnect();
			GroupsListener = null;
		}
		Search.Release(ref this.m_A);
		this.m_A = null;
		HiddenFiles = null;
		PreviewWinHandler.ClearReferences(ref this.m_A);
		DragDropOverlay = null;
	}

	private void A(ContentGroupsCollection A)
	{
		MySettings settings = PB.Settings;
		A.ShowsPersonalLibrary = settings.ContentInsertShowPersonal;
		A.ShowsSharedLibraries = settings.ContentInsertShowShared;
		A.ShowsPublicLibrary = settings.ContentInsertShowPublic;
		A.Shows3rdPartyLibraries = settings.ContentInsertShow3rdParty;
		A[(ContentType)1] = settings.ContentInsertShowSlides;
		A[(ContentType)3] = settings.ContentInsertShowShapes;
		A[(ContentType)4] = settings.ContentInsertShowImages;
		A[(ContentType)16] = settings.ContentInsertShowVideos;
		A[(ContentType)5] = settings.ContentInsertShowCharts;
		A[(ContentType)7] = settings.ContentInsertShowText;
		int num;
		if (settings.ContentInsertShowPDFs)
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
			num = (Pane.PDFsContentIsEnabled ? 1 : 0);
		}
		else
		{
			num = 0;
		}
		A[(ContentType)15] = (byte)num != 0;
		int num2;
		if (settings.ContentInsertShowDecks)
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
			num2 = (chkDecks.IsEnabled ? 1 : 0);
		}
		else
		{
			num2 = 0;
		}
		A[(ContentType)12] = (byte)num2 != 0;
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
		this.m_A = KG.A.SettingsXml;
		this.m_D = PB.Settings.LibraryPaneShowStars;
		this.m_A = NG.A.Application;
		HiddenFiles = new List<string>();
		IsAdmin = Base.IsUserAdmin();
		ItemsPanelWrap = (ItemsPanelTemplate)FindResource(AH.A(65100));
		ItemsPanelStack = (ItemsPanelTemplate)FindResource(AH.A(65129));
		AllGroups = Load.GetLibraryContent((Func<XmlDocument, string, LibraryItem, bool, ContentGroup>)A, (Action<ContentGroupsCollection>)A);
		B();
		GroupsListener.Groups = AllGroups;
		SourceCollection = CollectionViewSource.GetDefaultView(AllGroups);
		ICollectionView sourceCollection = SourceCollection;
		if (!sourceCollection.GroupDescriptions.Any())
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
			sourceCollection.GroupDescriptions.Add(new PropertyGroupDescription(AH.A(52438)));
		}
		sourceCollection.Filter = A;
		sourceCollection = null;
		chkContentFilters.IsChecked = true;
		chkImageTypeFilters.IsChecked = true;
		this.m_A = null;
	}

	private ContentGroup A(XmlDocument A, string B, LibraryItem C, bool D)
	{
		//IL_0252: Unknown result type (might be due to invalid IL or missing references)
		//IL_0255: Invalid comparison between Unknown and I4
		//IL_004c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Unknown result type (might be due to invalid IL or missing references)
		//IL_0053: Unknown result type (might be due to invalid IL or missing references)
		//IL_0058: Unknown result type (might be due to invalid IL or missing references)
		//IL_005f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0062: Invalid comparison between Unknown and I4
		//IL_0264: Unknown result type (might be due to invalid IL or missing references)
		//IL_0267: Invalid comparison between Unknown and I4
		//IL_0064: Unknown result type (might be due to invalid IL or missing references)
		//IL_0067: Invalid comparison between Unknown and I4
		//IL_0269: Unknown result type (might be due to invalid IL or missing references)
		//IL_026d: Invalid comparison between Unknown and I4
		//IL_00b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fb: Expected I4, but got Unknown
		//IL_0073: Unknown result type (might be due to invalid IL or missing references)
		//IL_0077: Invalid comparison between Unknown and I4
		//IL_01a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_01af: Invalid comparison between Unknown and I4
		//IL_01c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ef: Expected O, but got Unknown
		//IL_01f8: Unknown result type (might be due to invalid IL or missing references)
		//IL_0202: Expected O, but got Unknown
		//IL_020b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0215: Expected O, but got Unknown
		ContentGroup val = null;
		XmlNodeList childNodes = A.DocumentElement.ChildNodes;
		string directoryName;
		ContentType contentType;
		if (childNodes != null)
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
			if (childNodes.Count > 0)
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
				directoryName = Path.GetDirectoryName(B);
				contentType = Manifests.GetContentType(A);
				Base.ProcessManifestNodes(ref childNodes, directoryName, contentType);
				if ((int)contentType != 2 && (int)contentType != 8)
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
					if ((int)contentType != 12)
					{
						goto IL_00b2;
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
				}
				if (!Access.IsEnterprisePlanOrTrialMode())
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
				goto IL_00b2;
			}
		}
		goto IL_035c;
		IL_035c:
		childNodes = null;
		return val;
		IL_02ca:
		XmlNode xmlNode = null;
		goto IL_035c;
		IL_00b2:
		string a;
		switch (contentType - 1)
		{
		case 0:
			a = AH.A(65160);
			break;
		case 1:
			a = AH.A(65179);
			break;
		case 2:
			a = AH.A(65208);
			break;
		case 7:
			a = AH.A(65227);
			break;
		case 3:
			a = AH.A(65254);
			break;
		case 4:
			a = AH.A(65273);
			break;
		case 6:
			a = AH.A(65292);
			break;
		case 14:
			a = AH.A(65309);
			break;
		case 11:
			a = AH.A(65324);
			break;
		case 15:
			a = AH.A(65341);
			break;
		default:
			childNodes = null;
			return null;
		}
		bool flag = Favorites.ManifestContainsFavorite(A, contentType);
		ItemsPanelTemplate itemsPanelTemplate = (((int)contentType == 7) ? ItemsPanelStack : ItemsPanelWrap);
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
			this.C(AH.A(65360) + directoryName);
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
			if ((int)contentType != 2)
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
				if ((int)contentType == 8)
				{
					xmlNode = CustomFields.GetSchemaNodeById(this.m_A, val.MetadataSchemaId);
					if (xmlNode == null || xmlNode.ChildNodes.Count <= 0)
					{
						if (val.PitchlyTableId == null)
						{
							goto IL_02ca;
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
					}
					this.A(val, xmlNode);
					goto IL_02ca;
				}
				if ((int)contentType == 12)
				{
					xmlNode = CustomFields.GetSchemaNodeById(this.m_A, val.MetadataSchemaId);
					if (xmlNode != null)
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
						if (xmlNode.ChildNodes.Count > 0)
						{
							this.A(val, xmlNode);
						}
					}
					xmlNode = null;
				}
			}
			else
			{
				xmlNode = CustomFields.GetSchemaNodeById(this.m_A, val.MetadataSchemaId);
				if (xmlNode != null && xmlNode.ChildNodes.Count > 0)
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
					this.A(val, xmlNode);
				}
				xmlNode = null;
			}
		}
		goto IL_035c;
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
					icFilters.SetBinding(ItemsControl.ItemsSourceProperty, AH.A(64975));
					ProjectData.ClearProjectError();
				}
			}
			IEnumerable<FiltersGroup> source = ((IEnumerable<ContentGroup>)AllGroups).Select([SpecialName] (ContentGroup A) => A.FiltersGroup);
			Func<FiltersGroup, bool> predicate;
			if (_Closure_0024__.A == null)
			{
				predicate = (_Closure_0024__.A = [SpecialName] (FiltersGroup A) => A != null);
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
				predicate = _Closure_0024__.A;
			}
			IEnumerable<FiltersGroup> enumerable = source.Where(predicate);
			if (ContentFilters.SequenceEqual(enumerable))
			{
				return;
			}
			try
			{
				ContentFilters.Clear();
			}
			catch (Exception projectError2)
			{
				ProjectData.SetProjectError(projectError2);
				BindingOperations.ClearBinding(icFilters, ItemsControl.ItemsSourceProperty);
				icFilters.SetBinding(ItemsControl.ItemsSourceProperty, AH.A(64975));
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
			((System.Windows.Window)(object)DragDropOverlay).Close();
			DragDropOverlay = null;
		}
		List<string> list = HiddenFiles.Distinct().ToList();
		checked
		{
			for (int i = list.Count - 1; i >= 0; i += -1)
			{
				try
				{
					this.m_A.Presentations[Path.GetFileName(list[i])].Close();
					list.RemoveAt(i);
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
				switch (3)
				{
				case 0:
					continue;
				}
				ExcelApp.CleanUp(xlApp, list, blnQuitXL);
				JG.A(xlApp);
				xlApp = null;
				list = null;
				HiddenFiles.Clear();
				ArrangeView = null;
				return;
			}
		}
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
		//IL_00a2: Expected I4, but got Unknown
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
						switch (contentType - 1)
						{
						case 0:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)A);
							break;
						case 2:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)C);
							break;
						case 3:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)D);
							break;
						case 4:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)F);
							break;
						case 6:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)G);
							break;
						case 14:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)I);
							break;
						case 15:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)E);
							break;
						case 7:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)H);
							flag = true;
							break;
						case 1:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)B);
							flag = true;
							break;
						case 11:
							val2.CreateContentItems((Func<ContentGroup, XmlNode, ContentItem>)J);
							flag = true;
							break;
						case 5:
						case 8:
						case 9:
						case 10:
						case 12:
						case 13:
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
							switch (4)
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
					switch (2)
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
			switch (5)
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
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
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
		return (ContentItem)(object)new SlideItem(A, B);
	}

	private ContentItem B(ContentGroup A, XmlNode B)
	{
		return (ContentItem)(object)new MetaSlidesItem(A, B);
	}

	private ContentItem C(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new ShapeItem(A, B);
	}

	private ContentItem D(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new ImageItem(A, B);
	}

	private ContentItem E(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new VideoItem(A, B);
	}

	private ContentItem F(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new ChartItem(A, B);
	}

	private ContentItem G(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new TextItem(A, B);
	}

	private ContentItem H(ContentGroup A, XmlNode B)
	{
		if (string.IsNullOrEmpty(A.PitchlyTableId))
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
					return (ContentItem)(object)new MetaShapeItem(A, B);
				}
			}
		}
		return (ContentItem)(object)new PitchlyItem(A, B);
	}

	private ContentItem I(ContentGroup A, XmlNode B)
	{
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0008: Expected O, but got Unknown
		return (ContentItem)new PdfItem(A, B);
	}

	private ContentItem J(ContentGroup A, XmlNode B)
	{
		return (ContentItem)(object)new DeckItem(A, B);
	}

	private bool PreviewIsEnabled()
	{
		return chkPreview.IsChecked.Value;
	}

	private void ListViewItemSelected(object sender, RoutedEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		//IL_009d: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a6: Invalid comparison between Unknown and I4
		//IL_00b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b7: Invalid comparison between Unknown and I4
		System.Windows.Controls.ListViewItem listViewItem = (System.Windows.Controls.ListViewItem)sender;
		ContentItem val = (ContentItem)listViewItem.DataContext;
		System.Windows.Controls.ListView listView = (System.Windows.Controls.ListView)ItemsControl.ItemsControlFromItemContainer(listViewItem);
		Base.ScrollToItem(scroller, listViewItem);
		if (ActiveListView != null && listView != ActiveListView)
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
			if (ActiveListView.SelectedItems.Count > 0)
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
				ActiveListView.SelectedItems.Clear();
			}
		}
		ActiveListView = listView;
		A();
		if ((int)val.Group.ContentType != 12 && (int)val.Group.ContentType != 15)
		{
			SelActionText = AH.A(65425);
		}
		else
		{
			SelActionText = AH.A(65438);
		}
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
		bool value = ((System.Windows.Controls.CheckBox)sender).IsChecked.Value;
		PB.Settings.LibraryPaneShowStars = value;
		IEnumerator<ContentGroup> enumerator = default(IEnumerator<ContentGroup>);
		try
		{
			enumerator = ((Collection<ContentGroup>)(object)AllGroups).GetEnumerator();
			while (enumerator.MoveNext())
			{
				enumerator.Current.ShowFavorites = value;
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

	private void ImageTypeBadgeToggle(object sender, RoutedEventArgs e)
	{
		PB.Settings.LibraryPaneShowImageTypeBadge = ((System.Windows.Controls.CheckBox)sender).IsChecked.Value;
	}

	private void ArrangeToggle(object sender, RoutedEventArgs e)
	{
		PB.Settings.LibraryPaneOfferArrange = ((System.Windows.Controls.CheckBox)sender).IsChecked.Value;
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
						switch (7)
						{
						case 0:
							continue;
						}
						break;
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
			return;
		}
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

	public void DeleteTag(object sender, RoutedEventArgs e)
	{
		this.m_A.DeleteTag(sender as System.Windows.Controls.Button);
	}

	private void E(bool? A, bool? B, bool? C, bool? D)
	{
		if (A.HasValue)
		{
			PB.Settings.ContentInsertShowPersonal = A.Value;
		}
		if (B.HasValue)
		{
			PB.Settings.ContentInsertShowShared = B.Value;
		}
		if (C.HasValue)
		{
			PB.Settings.ContentInsertShowPublic = C.Value;
		}
		if (D.HasValue)
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
			PB.Settings.ContentInsertShow3rdParty = D.Value;
		}
		this.A(A: false);
	}

	private void A(ContentType A, bool B)
	{
		//IL_0000: Unknown result type (might be due to invalid IL or missing references)
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		//IL_0004: Unknown result type (might be due to invalid IL or missing references)
		//IL_0026: Expected I4, but got Unknown
		//IL_0026: Unknown result type (might be due to invalid IL or missing references)
		//IL_0029: Unknown result type (might be due to invalid IL or missing references)
		//IL_0043: Expected I4, but got Unknown
		switch (A - 1)
		{
		default:
			switch (A - 12)
			{
			case 3:
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				PB.Settings.ContentInsertShowPDFs = B;
				break;
			case 0:
				PB.Settings.ContentInsertShowDecks = B;
				break;
			case 4:
				PB.Settings.ContentInsertShowVideos = B;
				break;
			}
			break;
		case 2:
			PB.Settings.ContentInsertShowShapes = B;
			break;
		case 3:
			PB.Settings.ContentInsertShowImages = B;
			break;
		case 0:
			PB.Settings.ContentInsertShowSlides = B;
			break;
		case 4:
			PB.Settings.ContentInsertShowCharts = B;
			break;
		case 6:
			PB.Settings.ContentInsertShowText = B;
			break;
		case 1:
		case 5:
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
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			PB.Settings.ContentInsertExcludeImageTypes = AllGroups.ImgTypesFilter.ToListString();
		}
		this.A(A: false);
	}

	internal void A(bool A = false)
	{
		if (SourceCollection == null || GroupsListener.IsChangingGroups)
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
					selector = (_Closure_0024__.A = [SpecialName] (ContentGroup val) => ((Collection<ContentItem>)(object)val.AllContentItems)?.Count ?? 0);
				}
				else
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
					selector = _Closure_0024__.A;
				}
				if (((IEnumerable<ContentGroup>)allGroups).Sum(selector) > 100)
				{
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
					return;
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (5)
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
				switch (3)
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
			switch (2)
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
					switch (4)
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
			Visibility visibility = ((!num) ? Visibility.Collapsed : Visibility.Visible);
			if (A.FiltersGroup.GroupVisibility == visibility)
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
			A.ApplyItemsCriteria();
		}
		this.A(A, (bool?)flag);
		return flag;
	}

	private bool B(ContentGroup A)
	{
		if (AllGroups == null)
		{
			return false;
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
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
		return Filter.ApplyLibraryFilter(A, PB.Settings.ContentInsertShowShared, PB.Settings.ContentInsertShowPersonal, PB.Settings.ContentInsertShowPublic, PB.Settings.ContentInsertShow3rdParty, (Func<ContentGroup, bool>)C);
	}

	private bool C(ContentGroup A)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		//IL_0009: Invalid comparison between Unknown and I4
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_004e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Invalid comparison between Unknown and I4
		//IL_001f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		//IL_0027: Invalid comparison between Unknown and I4
		//IL_0080: Unknown result type (might be due to invalid IL or missing references)
		//IL_0086: Invalid comparison between Unknown and I4
		//IL_005e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0063: Unknown result type (might be due to invalid IL or missing references)
		//IL_0066: Invalid comparison between Unknown and I4
		//IL_00a8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b0: Invalid comparison between Unknown and I4
		//IL_00dc: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e2: Invalid comparison between Unknown and I4
		//IL_010e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0115: Invalid comparison between Unknown and I4
		//IL_0151: Unknown result type (might be due to invalid IL or missing references)
		//IL_0158: Invalid comparison between Unknown and I4
		//IL_0177: Unknown result type (might be due to invalid IL or missing references)
		//IL_017c: Unknown result type (might be due to invalid IL or missing references)
		//IL_0180: Invalid comparison between Unknown and I4
		if ((int)A.ContentType != 1)
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
			if ((int)A.ContentType != 2)
			{
				goto IL_0048;
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
		if (PB.Settings.ContentInsertShowSlides)
		{
			return true;
		}
		goto IL_0048;
		IL_0048:
		if ((int)A.ContentType != 3)
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
			if ((int)A.ContentType != 8)
			{
				goto IL_007f;
			}
		}
		if (PB.Settings.ContentInsertShowShapes)
		{
			return true;
		}
		goto IL_007f;
		IL_007f:
		if ((int)A.ContentType == 4)
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
			if (PB.Settings.ContentInsertShowImages)
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
			if (PB.Settings.ContentInsertShowCharts)
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
		if ((int)A.ContentType == 7)
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
			if (PB.Settings.ContentInsertShowText)
			{
				while (true)
				{
					switch (4)
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
				switch (3)
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
					switch (6)
					{
					case 0:
						continue;
					}
					break;
				}
				if (PB.Settings.ContentInsertShowPDFs)
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
		}
		if ((int)A.ContentType == 12 && PB.Settings.ContentInsertShowDecks)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					return true;
				}
			}
		}
		if ((int)A.ContentType == 16)
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
			if (PB.Settings.ContentInsertShowVideos)
			{
				return true;
			}
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
				Monitor.Exit(a);
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
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		ListField field = ((ListOption)((System.Windows.Controls.CheckBox)sender).DataContext).Field;
		field.ComboBoxText = Core.ListOptionText(field.Options);
		A((BaseField)(object)field);
		field = null;
	}

	private void FilterBooleanSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0034: Expected O, but got Unknown
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
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0018: Expected O, but got Unknown
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
				switch (1)
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

	private void ApplySearch(object sender, RoutedEventArgs e)
	{
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Expected O, but got Unknown
		if (!Access.AllowPowerPointOperation((PlanType)6, (Restriction)2, false))
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
			SavedSearch val = (SavedSearch)((System.Windows.Controls.Button)sender).DataContext;
			try
			{
				val.FiltersGroup.ApplySearch(val, this.m_C);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(ex2.Message);
				ProjectData.ClearProjectError();
			}
			val = null;
			return;
		}
	}

	private void SaveSearch(object sender, RoutedEventArgs e)
	{
		//IL_002a: Unknown result type (might be due to invalid IL or missing references)
		//IL_0030: Expected O, but got Unknown
		if (!Access.AllowPowerPointOperation((PlanType)6, (Restriction)1, false))
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
			FiltersGroup val = (FiltersGroup)((System.Windows.Controls.Button)sender).DataContext;
			try
			{
				val.SaveSearch();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(ex2.Message);
				ProjectData.ClearProjectError();
			}
			val = null;
			return;
		}
	}

	private void DeleteSearch(object sender, RoutedEventArgs e)
	{
		//IL_000d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Expected O, but got Unknown
		SavedSearch val = (SavedSearch)((System.Windows.Controls.Button)sender).DataContext;
		try
		{
			FiltersGroup filtersGroup = val.FiltersGroup;
			if (filtersGroup.DeleteSearch(val))
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
				filtersGroup.ResetFilters(ref this.m_C);
			}
			filtersGroup = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			C(ex2.Message);
			ProjectData.ClearProjectError();
		}
		val = null;
	}

	private void SortFieldChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_000b: Unknown result type (might be due to invalid IL or missing references)
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
		bool flag = D(val.Group);
		this.m_B = true;
		Base.ShowHideContextMenuOptionForContentTypes(contextMenu, val, AH.A(65521), flag, (ContentType[])(object)new ContentType[1] { (ContentType)1 });
		Base.ShowHideContextMenuOptionForContentTypes(contextMenu, val, AH.A(65546), true, (ContentType[])(object)new ContentType[2]
		{
			(ContentType)4,
			(ContentType)16
		});
		val = null;
		contextMenu = null;
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
		PB.Settings.LibraryPaneKeepSourceFormat = chkKeep.IsChecked.Value;
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
		//IL_00f4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fe: Unknown result type (might be due to invalid IL or missing references)
		//IL_0100: Unknown result type (might be due to invalid IL or missing references)
		//IL_0102: Unknown result type (might be due to invalid IL or missing references)
		//IL_0105: Unknown result type (might be due to invalid IL or missing references)
		//IL_014b: Expected I4, but got Unknown
		SlideRange slideRange = null;
		string text = string.Empty;
		bool flag = false;
		Microsoft.Office.Interop.PowerPoint.Shape shape = null;
		RectangleF? rect = null;
		if (ActiveListView != null)
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
			if (ActiveListView.SelectedItems.Count > 0)
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
				Selection selection = this.m_A.ActiveWindow.Selection;
				try
				{
					slideRange = selection.SlideRange;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				try
				{
					shape = selection.ShapeRange[1];
					Microsoft.Office.Interop.PowerPoint.Shape shape2 = shape;
					rect = new RectangleF(shape2.Left, shape2.Top, shape2.Width, shape2.Height);
					shape2 = null;
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				selection = null;
				ContentType contentType = ((ContentGroup)ActiveListView.DataContext).ContentType;
				switch (contentType - 1)
				{
				case 0:
					this.m_A.StartNewUndoEntry();
					I();
					break;
				case 1:
					this.m_A.StartNewUndoEntry();
					J();
					break;
				case 2:
					if (slideRange == null)
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
						text = AH.A(65569);
					}
					else
					{
						this.m_A.StartNewUndoEntry();
						B(shape);
						flag = true;
					}
					break;
				case 7:
					if (slideRange == null)
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
						text = AH.A(65569);
					}
					else
					{
						this.m_A.StartNewUndoEntry();
						if (ActiveListView.SelectedItems[0] is PitchlyItem)
						{
							D(shape);
						}
						else
						{
							C(shape);
						}
						flag = true;
					}
					break;
				case 3:
					if (slideRange == null)
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
						text = AH.A(65634);
					}
					else
					{
						this.m_A.StartNewUndoEntry();
						L();
						flag = true;
					}
					break;
				case 4:
					if (slideRange == null)
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
						text = AH.A(65699);
					}
					else
					{
						this.m_A.StartNewUndoEntry();
						N();
					}
					break;
				case 6:
					if (slideRange == null)
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
						text = AH.A(65764);
					}
					else
					{
						this.m_A.StartNewUndoEntry();
						O();
					}
					break;
				case 15:
					if (slideRange == null)
					{
						text = AH.A(65825);
						break;
					}
					this.m_A.StartNewUndoEntry();
					M();
					break;
				case 14:
					P();
					break;
				case 11:
					R();
					break;
				}
				if (Operators.CompareString(text, string.Empty, TextCompare: false) != 0)
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
					D(text);
				}
				else
				{
					bool? flag2;
					if (!flag)
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
						flag2 = false;
					}
					else
					{
						flag2 = chkArrange.IsChecked;
					}
					bool? flag3 = flag2;
					if (flag3.HasValue)
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
						if (flag3 != true)
						{
							goto IL_03b9;
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
					if (ActiveListView.SelectedItems.Count >= 4)
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
						if (flag3.HasValue)
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
							if (Access.AllowPowerPointOperation((PlanType)6, (Restriction)1, false))
							{
								ArrangeView = new wpfArrange(rect, T);
							}
							else
							{
								Forms.InfoMessage(AH.A(65890));
							}
						}
					}
				}
				goto IL_03b9;
			}
		}
		goto IL_03d3;
		IL_03b9:
		UsageLogger.LogInsertion(ActiveListView.SelectedItems.Cast<ContentItem>(), (OfficeApp)2);
		goto IL_03d3;
		IL_03d3:
		slideRange = null;
		shape = null;
	}

	private void B(string A)
	{
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)6, A);
	}

	private void I()
	{
		K();
		B(AH.A(66068));
	}

	private void J()
	{
		K();
		B(AH.A(66111));
	}

	private void K()
	{
		//IL_00a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ae: Expected O, but got Unknown
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
		{
			return;
		}
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			Microsoft.Office.Interop.PowerPoint.Presentation presentation = default(Microsoft.Office.Interop.PowerPoint.Presentation);
			Slide slide = default(Slide);
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
				List<Slide> list = new List<Slide>();
				bool value = chkKeep.IsChecked.Value;
				if (A())
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
					PowerPointAddIn1.Library2.Insert.Slides.A(this.m_A);
					Microsoft.Office.Interop.PowerPoint.Presentation activePresentation;
					try
					{
						activePresentation = this.m_A.ActivePresentation;
						string c = Content.RightNow();
						string d = A();
						try
						{
							enumerator = ActiveListView.SelectedItems.GetEnumerator();
							while (true)
							{
								if (enumerator.MoveNext())
								{
									ContentItem val = (ContentItem)enumerator.Current;
									presentation = A(val);
									PageSetup pageSetup = activePresentation.PageSetup;
									bool flag = presentation.PageSetup.SlideWidth != pageSetup.SlideWidth;
									bool flag2 = presentation.PageSetup.SlideHeight != pageSetup.SlideHeight;
									pageSetup = null;
									int count = activePresentation.Designs.Count;
									_ = null;
									int num2;
									int num;
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
										try
										{
											num = PowerPointAddIn1.Slides.Helpers.GetSlideIndex() + 1;
										}
										catch (Exception ex)
										{
											ProjectData.SetProjectError(ex);
											Exception ex2 = ex;
											num = 1;
											ProjectData.ClearProjectError();
										}
										try
										{
											num2 = Slides.A(presentation, this.m_A);
										}
										catch (Exception ex3)
										{
											ProjectData.SetProjectError(ex3);
											Exception ex4 = ex3;
											C(ex4.Message);
											clsReporting.LogException(ex4);
											ProjectData.ClearProjectError();
											break;
										}
									}
									else
									{
										try
										{
											num = PowerPointAddIn1.Slides.Helpers.GetSlideIndex();
										}
										catch (Exception ex5)
										{
											ProjectData.SetProjectError(ex5);
											Exception ex6 = ex5;
											num = 0;
											ProjectData.ClearProjectError();
										}
										num2 = activePresentation.Slides.InsertFromFile(val.ContentPath, num);
										num++;
									}
									int num3 = num2 - 1;
									for (int i = 0; i <= num3; list.Add(slide), i++)
									{
										slide = activePresentation.Slides[num + i];
										string b = Base.ConvertCachedToRemotePath(val.ContentPath, B(val));
										PowerPointAddIn1.Links.Slides.A(slide, b, A(val), value);
										A(slide, val, c, d, value);
										if (value && AirplaneMode.IsOn())
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
											AirplaneMode.HideSlideImages(slide);
										}
										if (value)
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
											if (PowerPointAddIn1.Explorer.Pane.IsOpen)
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
												try
												{
													Events.PresentationNewSlide(slide);
												}
												catch (Exception ex7)
												{
													ProjectData.SetProjectError(ex7);
													Exception ex8 = ex7;
													clsReporting.LogException(ex8);
													ProjectData.ClearProjectError();
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
													continue;
												}
												break;
											}
											if (!flag2)
											{
												continue;
											}
										}
										try
										{
											PowerPointAddIn1.Shapes.Images.FixDistortion(slide, flag, flag2);
										}
										catch (Exception ex9)
										{
											ProjectData.SetProjectError(ex9);
											Exception ex10 = ex9;
											ProjectData.ClearProjectError();
										}
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
									if (!flag)
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
									}
									try
									{
										Designs designs = activePresentation.Designs;
										if (designs.Count > count)
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
											PowerPointAddIn1.Shapes.Images.FixDistortion(designs[designs.Count], flag, flag2);
										}
										designs = null;
									}
									catch (Exception ex11)
									{
										ProjectData.SetProjectError(ex11);
										Exception ex12 = ex11;
										ProjectData.ClearProjectError();
									}
									continue;
								}
								while (true)
								{
									switch (5)
									{
									case 0:
										break;
									default:
										goto end_IL_033c;
									}
									continue;
									end_IL_033c:
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
					}
					finally
					{
						PowerPointAddIn1.Library2.Insert.Slides.B(this.m_A);
					}
					if (ActiveListView.SelectedItems.Count <= 1)
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
						if (value)
						{
							goto IL_03f3;
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
					try
					{
						List<Slide> source = list;
						Func<Slide, int> selector;
						if (_Closure_0024__.A == null)
						{
							selector = (_Closure_0024__.A = [SpecialName] (Slide A) => A.SlideIndex);
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
						PowerPointAddIn1.Slides.Helpers.SelectMultipleSlides(activePresentation, source.Select(selector));
					}
					catch (Exception ex13)
					{
						ProjectData.SetProjectError(ex13);
						Exception ex14 = ex13;
						ProjectData.ClearProjectError();
					}
					goto IL_03f3;
					IL_03f3:
					Focus();
					JG.A(presentation);
					JG.A(activePresentation);
					JG.A(slide);
					list = null;
					return;
				}
			}
		}
	}

	private string A(ContentItem A)
	{
		if (A is SlideItem)
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
					return ((SlideItem)(object)A).LinkId;
				}
			}
		}
		return ((MetaSlidesItem)(object)A).LinkId;
	}

	private string B(ContentItem A)
	{
		return A.Group.Library.Id;
	}

	private bool A()
	{
		switch (this.m_A.ActiveWindow.ActivePane.ViewType)
		{
		case PpViewType.ppViewSlideMaster:
		case PpViewType.ppViewHandoutMaster:
		case PpViewType.ppViewNotesMaster:
		case PpViewType.ppViewMasterThumbnails:
			D(AH.A(66164));
			return true;
		case PpViewType.ppViewNotesPage:
			D(AH.A(66257));
			return true;
		default:
			_ = null;
			return false;
		}
	}

	private void L()
	{
		//IL_029f: Unknown result type (might be due to invalid IL or missing references)
		//IL_02a6: Expected O, but got Unknown
		//IL_0167: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f1: Expected O, but got Unknown
		Selection selection = null;
		Microsoft.Office.Interop.PowerPoint.Presentation activePresentation = this.m_A.ActivePresentation;
		bool flag = false;
		Microsoft.Office.Interop.PowerPoint.Shape shape;
		List<string> list;
		try
		{
			selection = this.m_A.ActiveWindow.Selection;
			try
			{
				if (selection.ShapeRange.Count == 1)
				{
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
						if (ActiveListView.SelectedItems.Count != 1)
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
							flag = PowerPointAddIn1.Library2.Insert.Images.A(selection.ShapeRange[1]);
							break;
						}
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
			float slideWidth = default(float);
			float slideHeight = default(float);
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
				PageSetup pageSetup = activePresentation.PageSetup;
				slideWidth = pageSetup.SlideWidth;
				slideHeight = pageSetup.SlideHeight;
				_ = null;
			}
			string c = Content.RightNow();
			string d = A();
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
				Slide slide = selection.SlideRange[1];
				int B = 0;
				Dictionary<Microsoft.Office.Interop.PowerPoint.Shape, PowerPointAddIn1.Library2.Insert.Images.FD> dictionary = new Dictionary<Microsoft.Office.Interop.PowerPoint.Shape, PowerPointAddIn1.Library2.Insert.Images.FD>();
				Dictionary<int, PowerPointAddIn1.Library2.Insert.Images.FD> C = new Dictionary<int, PowerPointAddIn1.Library2.Insert.Images.FD>();
				Microsoft.Office.Interop.PowerPoint.Shape a = selection.ShapeRange[1];
				a = PowerPointAddIn1.Library2.Insert.Images.A(a, slide);
				PowerPointAddIn1.Library2.Insert.Images.A(slide, ref B, ref C);
				if (B > 0)
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
					int zOrderPosition = a.ZOrderPosition;
					string contentPath = ((ContentItem)ActiveListView.SelectedItems[0]).ContentPath;
					int num = B;
					int num2 = 1;
					while (true)
					{
						if (num2 <= num)
						{
							shape = PowerPointAddIn1.Library2.Insert.Images.A(slide, contentPath);
							if (shape.ZOrderPosition != zOrderPosition)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_0199;
									}
									continue;
									end_IL_0199:
									break;
								}
								dictionary.Add(shape, C[shape.ZOrderPosition]);
								num2 = checked(num2 + 1);
								continue;
							}
							PowerPointAddIn1.Library2.Insert.Images.A(shape, this.m_A);
							A(shape, (ContentItem)ActiveListView.SelectedItems[0], c, d);
							break;
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
						break;
					}
					PowerPointAddIn1.Library2.Insert.Images.A(slide, dictionary);
				}
				dictionary = null;
				C = null;
				slide = null;
				a = null;
			}
			else
			{
				if (ActiveListView.SelectedItems.Count > 1)
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
						if (selection.ShapeRange.Count > 0)
						{
							while (true)
							{
								switch (1)
								{
								case 0:
									continue;
								}
								selection.Unselect();
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
				list = new List<string>();
				{
					IEnumerator enumerator = ActiveListView.SelectedItems.GetEnumerator();
					try
					{
						while (enumerator.MoveNext())
						{
							ContentItem val = (ContentItem)enumerator.Current;
							shape = PowerPointAddIn1.Library2.Insert.Images.A(this.m_A, selection.SlideRange[1], val.ContentPath);
							A(shape, slideWidth, slideHeight);
							list.Add(shape.Name);
							A(shape, val, c, d);
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_0300;
							}
							continue;
							end_IL_0300:
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
				}
				this.m_A.ActiveWindow.Selection.SlideRange[1].Shapes.Range(list.ToArray()).Select();
			}
			this.B(AH.A(66356));
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			this.C(ex6.Message);
			ProjectData.ClearProjectError();
		}
		selection = null;
		shape = null;
		activePresentation = null;
		list = null;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, float B, float C)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		shape.LockAspectRatio = MsoTriState.msoTrue;
		if (shape.Width / shape.Height > B / C)
		{
			if (shape.Width > B)
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
				shape.Width = B;
				shape.Left = 0f;
			}
		}
		else if (shape.Height > C)
		{
			shape.Height = C;
			shape.Top = 0f;
		}
		shape = null;
	}

	private void M()
	{
		throw new NotImplementedException();
	}

	private void N()
	{
		//IL_0114: Unknown result type (might be due to invalid IL or missing references)
		//IL_011b: Expected O, but got Unknown
		//IL_0124: Unknown result type (might be due to invalid IL or missing references)
		//IL_0130: Expected O, but got Unknown
		//IL_00ae: Unknown result type (might be due to invalid IL or missing references)
		RD a = default(RD);
		RD CS_0024_003C_003E8__locals10 = new RD(a);
		if (!Access.AllowPowerPointOperation((PlanType)4, (Restriction)1, false))
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
			Microsoft.Office.Interop.PowerPoint.Shape shape = null;
			bool value = chkKeep.IsChecked.Value;
			PpViewType viewType = this.m_A.ActiveWindow.ViewType;
			if (viewType != PpViewType.ppViewSlide && viewType != PpViewType.ppViewNormal)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						D(AH.A(66399));
						return;
					}
				}
			}
			workbook = ExcelApp.GetWorkbook(ref xlApp, ref HiddenFiles, ref blnQuitXL, ((ContentItem)ActiveListView.SelectedItems[0]).ContentPath);
			if (workbook != null)
			{
				bool copyObjectsWithCells = xlApp.CopyObjectsWithCells;
				xlApp.CopyObjectsWithCells = true;
				Workbook workbook2;
				try
				{
					string c = Content.RightNow();
					string d = A();
					try
					{
						enumerator = ActiveListView.SelectedItems.GetEnumerator();
						while (enumerator.MoveNext())
						{
							ContentItem val = (ContentItem)enumerator.Current;
							workbook2 = ExcelApp.CreateTargetWorkbook(xlApp, workbook, (ChartItem)val, value);
							CS_0024_003C_003E8__locals10.A = ((ChartObject)((Worksheet)workbook2.Worksheets[1]).ChartObjects(1)).Chart;
							Action action;
							if (CS_0024_003C_003E8__locals10.A != null)
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
								action = CS_0024_003C_003E8__locals10.A;
							}
							else
							{
								action = (CS_0024_003C_003E8__locals10.A = [SpecialName] () =>
								{
									CS_0024_003C_003E8__locals10.A.ChartArea.Copy();
								});
							}
							clsClipboard.CopyWithWait(action, 4000);
							try
							{
								if (value)
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
									this.m_A.CommandBars.ExecuteMso(AH.A(66466));
								}
								else
								{
									this.m_A.CommandBars.ExecuteMso(AH.A(66529));
								}
								System.Windows.Forms.Application.DoEvents();
								shape = this.m_A.ActiveWindow.Selection.ShapeRange[1];
								shape.Chart.ChartData.ActivateChartDataWindow();
								if (value)
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
										string text = Path.Combine(Interaction.Environ(AH.A(66592)), AH.A(66601));
										CS_0024_003C_003E8__locals10.A.SaveChartTemplate(text);
										shape.Chart.ApplyChartTemplate(text);
										File.Delete(text);
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										ProjectData.ClearProjectError();
									}
								}
								A(shape, val, c, d);
							}
							catch (Exception ex3)
							{
								ProjectData.SetProjectError(ex3);
								Exception ex4 = ex3;
								try
								{
									CS_0024_003C_003E8__locals10.A = CS_0024_003C_003E8__locals10.A.Location(XlChartLocation.xlLocationAsNewSheet, RuntimeHelpers.GetObjectValue(Missing.Value));
									clsClipboard.CopyWithWait((Action)([SpecialName] () =>
									{
										CS_0024_003C_003E8__locals10.A.ChartArea.Copy();
									}), 4000);
									this.m_A.ActiveWindow.View.PasteSpecial(PpPasteDataType.ppPasteOLEObject);
									clsReporting.LogException(ex4);
								}
								catch (ArgumentException ex5)
								{
									ProjectData.SetProjectError(ex5);
									ArgumentException ex6 = ex5;
									if (!((Worksheet)workbook2.Worksheets[1]).ProtectContents)
									{
										throw ex6;
									}
									C(AH.A(66628));
									ProjectData.ClearProjectError();
								}
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
							_ = ActiveListView.SelectedItems.Count;
							_ = 1;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_03e2;
							}
							continue;
							end_IL_03e2:
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
					B(AH.A(66904));
				}
				catch (Exception ex9)
				{
					ProjectData.SetProjectError(ex9);
					Exception ex10 = ex9;
					C(AH.A(66947) + ex10.Message);
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
				workbook2 = null;
				workbook = null;
				CS_0024_003C_003E8__locals10.A = null;
				shape = null;
			}
			if (xlApp == null)
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
				xlApp.DisplayAlerts = true;
				xlApp.EnableEvents = true;
				xlApp.ScreenUpdating = true;
				return;
			}
		}
	}

	private void O()
	{
		//IL_0024: Unknown result type (might be due to invalid IL or missing references)
		Selection selection;
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = ActiveListView.SelectedItems.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					string text = File.ReadAllText(((ContentItem)enumerator.Current).ContentPath);
					selection = this.m_A.ActiveWindow.Selection;
					switch (selection.Type)
					{
					case PpSelectionType.ppSelectionShapes:
						try
						{
							enumerator2 = selection.ShapeRange.GetEnumerator();
							while (enumerator2.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
								if (shape.HasTextFrame != MsoTriState.msoTrue)
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
								shape.TextFrame2.TextRange.Text = text;
							}
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									goto end_IL_00c4;
								}
								continue;
								end_IL_00c4:
								break;
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
						break;
					case PpSelectionType.ppSelectionText:
						selection.TextRange2.Text = text;
						break;
					default:
						selection.SlideRange[1].Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0f, 0f, 400f, 400f).TextFrame2.TextRange.Text = text;
						_ = null;
						break;
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
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			B(AH.A(67042));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			C(ex2.Message);
			ProjectData.ClearProjectError();
		}
		selection = null;
	}

	private void P()
	{
		B(AH.A(67081));
	}

	private void B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		//IL_0074: Unknown result type (might be due to invalid IL or missing references)
		//IL_007e: Expected O, but got Unknown
		//IL_00a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ad: Expected O, but got Unknown
		//IL_00dd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e8: Expected O, but got Unknown
		//IL_00bd: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d1: Expected O, but got Unknown
		SD a = default(SD);
		SD CS_0024_003C_003E8__locals4 = new SD(a);
		DialogResult dialogResult = this.A(A);
		if (dialogResult == DialogResult.Cancel)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
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
			Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
			List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
			string c = Content.RightNow();
			string d = this.A();
			if (Insert.AreShapesLegacyPublished(ActiveListView))
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
				presentation = this.A((ContentItem)ActiveListView.SelectedItems[0]);
			}
			try
			{
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
							CS_0024_003C_003E8__locals4.A = PowerPointApp.GetSourceShape((ShapeItem)val, this.A(val));
						}
						else
						{
							CS_0024_003C_003E8__locals4.A = PowerPointApp.GetSourceShape((ShapeItem)val, presentation);
						}
						clsClipboard.CopyWithWait((Action)([SpecialName] () =>
						{
							CS_0024_003C_003E8__locals4.A.Copy();
						}), 4000);
						shape = PowerPointAddIn1.Library2.Insert.Shapes.A(this.m_A);
						clsClipboard.ClearClipboard();
						list.Add(shape);
						Q(shape, A);
						this.A(shape, val, c, d);
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_015f;
						}
						continue;
						end_IL_015f:
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
				this.A(A, dialogResult, list, AH.A(67098));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(ex2.Message);
				ProjectData.ClearProjectError();
			}
			JG.A(presentation);
			JG.A(CS_0024_003C_003E8__locals4.A);
			JG.A(shape);
			list = null;
			return;
		}
	}

	private void C(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		//IL_0064: Unknown result type (might be due to invalid IL or missing references)
		//IL_006b: Expected O, but got Unknown
		TD a = default(TD);
		TD CS_0024_003C_003E8__locals6 = new TD(a);
		DialogResult dialogResult = this.A(A);
		if (dialogResult == DialogResult.Cancel)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		Microsoft.Office.Interop.PowerPoint.Shape shape = default(Microsoft.Office.Interop.PowerPoint.Shape);
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
			List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
			string c = Content.RightNow();
			string d = this.A();
			try
			{
				try
				{
					enumerator = ActiveListView.SelectedItems.GetEnumerator();
					while (enumerator.MoveNext())
					{
						ContentItem val = (ContentItem)enumerator.Current;
						if (Operators.CompareString(Path.GetExtension(((ContentItem)(MetaShapeItem)(object)val).ContentPath), AH.A(67141), TextCompare: false) == 0)
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
							shape = this.m_A.ActiveWindow.Selection.SlideRange[1].Shapes.AddPicture(((ContentItem)(MetaShapeItem)(object)val).ContentPath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0f, 0f);
							Q(shape, A);
							list.Add(shape);
							this.A(shape, val, c, d);
							continue;
						}
						CS_0024_003C_003E8__locals6.A = this.A(val).Slides[1].Shapes[1];
						Action action;
						if (CS_0024_003C_003E8__locals6.A != null)
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
							action = CS_0024_003C_003E8__locals6.A;
						}
						else
						{
							action = (CS_0024_003C_003E8__locals6.A = [SpecialName] () =>
							{
								CS_0024_003C_003E8__locals6.A.Copy();
							});
						}
						clsClipboard.CopyWithWait(action, 4000);
						shape = PowerPointAddIn1.Library2.Insert.Shapes.A(this.m_A);
						list.Add(shape);
						Q(shape, A);
						this.A(shape, val, c, d);
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
				this.A(A, dialogResult, list, AH.A(67150));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(ex2.Message);
				ProjectData.ClearProjectError();
			}
			JG.A(CS_0024_003C_003E8__locals6.A);
			JG.A(shape);
			list = null;
			return;
		}
	}

	private void D(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		//IL_0059: Unknown result type (might be due to invalid IL or missing references)
		//IL_0060: Expected O, but got Unknown
		DialogResult dialogResult = this.A(A);
		if (dialogResult == DialogResult.Cancel)
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
			List<Microsoft.Office.Interop.PowerPoint.Shape> list = new List<Microsoft.Office.Interop.PowerPoint.Shape>();
			string c = Content.RightNow();
			string d = this.A();
			Microsoft.Office.Interop.PowerPoint.Shape shape;
			try
			{
				try
				{
					enumerator = ActiveListView.SelectedItems.GetEnumerator();
					while (enumerator.MoveNext())
					{
						ContentItem val = (ContentItem)enumerator.Current;
						this.A((PitchlyItem)(object)val);
						shape = PowerPointAddIn1.Library2.Insert.Shapes.A(this.m_A);
						list.Add(shape);
						Q(shape, A);
						this.A(shape, val, c, d);
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
				this.A(A, dialogResult, list, AH.A(67203));
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				C(ex2.Message);
				ProjectData.ClearProjectError();
			}
			list = null;
			shape = null;
			return;
		}
	}

	private void A(PitchlyItem A)
	{
		HD.A(this.A((ContentItem)(object)A));
	}

	private DialogResult A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		DialogResult dialogResult = DialogResult.No;
		if (ActiveListView.SelectedItems.Count == 1)
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
			if (A != null)
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
				if (A.Type != MsoShapeType.msoPlaceholder)
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
					dialogResult = System.Windows.Forms.MessageBox.Show(AH.A(67262), AH.A(5874), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
					if (dialogResult == DialogResult.Cancel)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								return dialogResult;
							}
						}
					}
				}
				if (this.m_A.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionText)
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
					A.Select();
				}
			}
		}
		return dialogResult;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, DialogResult B, List<Microsoft.Office.Interop.PowerPoint.Shape> C, string D)
	{
		try
		{
			if (A != null)
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
					if (B != DialogResult.Yes)
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
						A.Delete();
						break;
					}
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
		this.A(this.m_A, C);
		this.B(D);
	}

	private void Q(Microsoft.Office.Interop.PowerPoint.Shape A, Microsoft.Office.Interop.PowerPoint.Shape B)
	{
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		if (B != null)
		{
			shape.Top = B.Top;
			shape.Left = B.Left;
		}
		else
		{
			shape.Top = this.m_A.ActivePresentation.PageSetup.SlideHeight / 2f - shape.Height / 2f;
			shape.Left = this.m_A.ActivePresentation.PageSetup.SlideWidth / 2f - shape.Width / 2f;
		}
		shape = null;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Application A, List<Microsoft.Office.Interop.PowerPoint.Shape> B)
	{
		checked
		{
			try
			{
				Func<Microsoft.Office.Interop.PowerPoint.Shape, string> selector;
				if (_Closure_0024__.A == null)
				{
					selector = (_Closure_0024__.A = [SpecialName] (Microsoft.Office.Interop.PowerPoint.Shape shape) => shape.Name);
				}
				else
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
					selector = _Closure_0024__.A;
				}
				List<string> list = B.Select(selector).ToList();
				if (list.Distinct().Count() == B.Count)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							A.ActiveWindow.Selection.SlideRange[1].Shapes.Range(list.ToArray()).Select();
							return;
						}
					}
				}
				int num = B.Count - 1;
				for (int num2 = 0; num2 <= num; num2++)
				{
					if (num2 > 0)
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
						B[num2].Select(MsoTriState.msoFalse);
					}
					else
					{
						B[num2].Select();
					}
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
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			finally
			{
				List<string> list = null;
			}
		}
	}

	private void R()
	{
		//IL_00a3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a9: Expected O, but got Unknown
		if (ActiveListView.SelectedItems.Count > 5)
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
			if (System.Windows.Forms.MessageBox.Show(AH.A(67416) + ActiveListView.SelectedItems.Count + AH.A(67485), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
			{
				return;
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = ActiveListView.SelectedItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				ContentItem val = (ContentItem)enumerator.Current;
				this.m_A.Presentations.Open(val.ContentPath, MsoTriState.msoFalse, MsoTriState.msoTrue);
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
		B(AH.A(67534));
	}

	private Microsoft.Office.Interop.PowerPoint.Presentation A(ContentItem A)
	{
		return PowerPointAddIn1.Library2.Insert.Common.A(A.ContentPath, this.m_A, ref HiddenFiles);
	}

	private void S()
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
		//IL_0047: Unknown result type (might be due to invalid IL or missing references)
		//IL_0051: Expected O, but got Unknown
		if (DragDropOverlay != null)
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
			try
			{
				DragDropOverlay = new wpfDragDrop((Action)S, (IntPtr)this.m_A.ActiveWindow.HWND);
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

	private void EditSlide(object sender, RoutedEventArgs e)
	{
		ContentItem val = A(RuntimeHelpers.GetObjectValue(sender));
		if (D(val.Group))
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
			this.m_A.Presentations.Open(val.ContentPath, MsoTriState.msoFalse, MsoTriState.msoTrue);
			Pane.Toggle(blnShow: true);
		}
		else
		{
			U();
		}
		val = null;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, ContentItem B, string C, string D)
	{
		Tagging.A(A, B, C, D);
	}

	private void A(Slide A, ContentItem B, string C, string D, bool E)
	{
		Tagging.A(A, B, C, D, E);
	}

	private string A()
	{
		return Core.GetAuthor(this.m_A.ActivePresentation);
	}

	private void T()
	{
		ArrangeView = null;
	}

	private bool D(ContentGroup A)
	{
		return Access.UserHasAccess(A.Library, (AccessType)1, IsAdmin);
	}

	private void U()
	{
		D(AH.A(67571));
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
		if (!this.m_K)
		{
			this.m_K = true;
			Uri resourceLocator = new Uri(AH.A(67707), UriKind.Relative);
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
		if (connectionId == 13)
		{
			((System.Windows.Controls.MenuItem)target).Click += InsertContent;
			return;
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
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
					((System.Windows.Controls.MenuItem)target).Click += EditSlide;
					return;
				}
			}
		}
		if (connectionId == 16)
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
		if (connectionId == 17)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					rtbSearch = (System.Windows.Controls.RichTextBox)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkShared = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			chkPersonal = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chk3rdParty = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkPublic = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
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
					scroller = (ScrollViewer)target;
					scroller.PreviewMouseWheel += ScrollViewer_PreviewMouseWheel;
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
					chkFilters = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 24)
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
		if (connectionId == 25)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkContentFilters = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					grdContentFilters = (DockPanel)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkSlides = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 28)
		{
			while (true)
			{
				switch (4)
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
				switch (2)
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
				switch (7)
				{
				case 0:
					break;
				default:
					chkVideos = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 31)
		{
			chkCharts = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 32)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkText = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 33)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkPdfs = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 34)
		{
			chkDecks = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 35)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkFavorites = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 36)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkImageTypeFilters = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 37)
		{
			grdImageTypeFilters = (StackPanel)target;
			return;
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
					chkImagesSvg = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 39)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkImagesPng = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 40)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkImagesJpg = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 41)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkImagesEmf = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
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
					chkImagesWmf = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 43)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkImagesGif = (System.Windows.Controls.CheckBox)target;
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
					chkImagesBmp = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 45)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkImagesTiff = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 46)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					icFilters = (ItemsControl)target;
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
					icContent = (ItemsControl)target;
					return;
				}
			}
		}
		if (connectionId == 57)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkPreview = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 58)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkStars = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 59)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkImageTypeBadge = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 60)
		{
			chkArrange = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 61)
		{
			btnInsert = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 62)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkKeep = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 63)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					popRemoveGroup = (Popup)target;
					popRemoveGroup.Opened += MorePopupOpened;
					popRemoveGroup.Closed += MorePopupClosed;
					popRemoveGroup.PreviewKeyDown += CloseMorePopup;
					return;
				}
			}
		}
		if (connectionId == 64)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnHideContent = (System.Windows.Controls.Button)target;
					btnHideContent.Click += HideContentGroup;
					return;
				}
			}
		}
		this.m_K = true;
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
			((System.Windows.Controls.Button)target).Click += DeleteTag;
		}
		if (connectionId == 2)
		{
			((System.Windows.Controls.Button)target).Click += ShowMoreMenu;
		}
		if (connectionId == 3)
		{
			((System.Windows.Controls.TextBox)target).TextChanged += FilterTextChanged;
		}
		if (connectionId == 4)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += FilterListSelectionChanged;
		}
		if (connectionId == 5)
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
			((System.Windows.Controls.CheckBox)target).Checked += ListItemCheckChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += ListItemCheckChanged;
		}
		if (connectionId == 6)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += FilterBooleanSelectionChanged;
		}
		if (connectionId == 7)
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
			((DatePicker)target).SelectedDateChanged += FilterMinDateChanged;
		}
		if (connectionId == 8)
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
			((DatePicker)target).SelectedDateChanged += FilterMaxDateChanged;
		}
		if (connectionId == 9)
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
			((System.Windows.Controls.TextBox)target).TextChanged += FilterMinValueChanged;
		}
		if (connectionId == 10)
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
		if (connectionId == 11)
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
		if (connectionId == 12)
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
			((System.Windows.Controls.TextBox)target).TextChanged += FilterMaxValueChanged;
		}
		if (connectionId == 47)
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
			((System.Windows.Controls.Button)target).Click += SaveSearch;
		}
		if (connectionId == 48)
		{
			((System.Windows.Controls.Button)target).Click += ApplySearch;
		}
		if (connectionId == 49)
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
			((System.Windows.Controls.Button)target).Click += DeleteSearch;
		}
		if (connectionId == 50)
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
			((System.Windows.Controls.ComboBox)target).SelectionChanged += SortFieldChanged;
		}
		if (connectionId == 51)
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
			((TextBlock)target).MouseUp += SortOrderChanged;
		}
		if (connectionId == 52)
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
			((TextBlock)target).MouseUp += ResetFilters;
		}
		if (connectionId == 54)
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
			((System.Windows.Controls.CheckBox)target).Checked += ExpandGroup;
		}
		if (connectionId == 55)
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
			((System.Windows.Controls.ListView)target).PreviewMouseLeftButtonDown += lstView_PreviewMouseLeftButtonDown;
			((System.Windows.Controls.ListView)target).MouseMove += lstView_MouseMove;
			((System.Windows.Controls.ListView)target).DragLeave += lstView_DragLeave;
			((System.Windows.Controls.ListView)target).SelectionChanged += lstView_SelectionChanged;
		}
		if (connectionId != 56)
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
