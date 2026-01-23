using System;
using System.CodeDom.Compiler;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Markup;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.ImportExport;
using MacabacusMacros.Links;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Links;

[DesignerGenerated]
public sealed class wpfManageLinks : System.Windows.Window, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	private struct CF
	{
		public string A;

		public string B;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<ExcelLinkItem, bool> A;

		public static Func<ExcelLinkItem, bool> B;

		public static Func<ExcelLinkItem, bool> C;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(ExcelLinkItem A)
		{
			return ((LinkItem)A).IsChecked;
		}

		[SpecialName]
		internal bool B(ExcelLinkItem A)
		{
			return ((LinkItem)A).IsChecked;
		}

		[SpecialName]
		internal bool C(ExcelLinkItem A)
		{
			return ((LinkItem)A).IsChecked;
		}
	}

	[CompilerGenerated]
	internal sealed class DF
	{
		public ExcelLinkItem A;

		public EF A;

		public DF(DF A)
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
			//IL_00d4: Unknown result type (might be due to invalid IL or missing references)
			//IL_00fa: Expected O, but got Unknown
			//IL_028b: Unknown result type (might be due to invalid IL or missing references)
			//IL_0290: Unknown result type (might be due to invalid IL or missing references)
			//IL_0175: Unknown result type (might be due to invalid IL or missing references)
			//IL_017a: Unknown result type (might be due to invalid IL or missing references)
			//IL_017c: Unknown result type (might be due to invalid IL or missing references)
			//IL_0326: Unknown result type (might be due to invalid IL or missing references)
			//IL_0239: Unknown result type (might be due to invalid IL or missing references)
			//IL_0230: Unknown result type (might be due to invalid IL or missing references)
			//IL_0236: Unknown result type (might be due to invalid IL or missing references)
			//IL_023e: Unknown result type (might be due to invalid IL or missing references)
			//IL_0256: Unknown result type (might be due to invalid IL or missing references)
			this.A.A.m_A.ActiveWindow.View.GotoSlide(this.A.Slide.SlideIndex);
			if (this.A.A is Microsoft.Office.Interop.PowerPoint.Shape)
			{
				this.A.LinkedObject = Shapes.Refresh((Microsoft.Office.Interop.PowerPoint.Shape)this.A.A, blnAll: false, ref this.A.A, ref this.A.A.m_A, this.A.A, this.A.A);
			}
			else if (this.A.A is TextLink)
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
				this.A.LinkedObject = Text.Refresh((TextLink)this.A.A, blnAll: false, ref this.A.A, ref this.A.A.m_A);
			}
			else
			{
				this.A.LinkedObject = Hyperlinks.Refresh((Microsoft.Office.Interop.PowerPoint.Hyperlink)this.A.A, blnAll: false, ref this.A.A, ref this.A.A.m_A);
			}
			if (this.A.LinkedObject is Microsoft.Office.Interop.PowerPoint.Shape)
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
				this.A.A = Shapes.LinkDetails((Microsoft.Office.Interop.PowerPoint.Shape)this.A.LinkedObject);
				this.A.A = (Microsoft.Office.Interop.PowerPoint.Shape)this.A.LinkedObject;
			}
			else if (this.A.LinkedObject is TextLink)
			{
				EF eF = this.A;
				Type typeFromHandle = typeof(Text);
				string memberName = AH.A(93278);
				ExcelLinkItem excelLinkItem;
				object[] obj = new object[1] { (excelLinkItem = this.A).LinkedObject };
				object[] array = obj;
				bool[] obj2 = new bool[1] { true };
				bool[] array2 = obj2;
				object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
				if (array2[0])
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
					excelLinkItem.LinkedObject = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(array[0]));
				}
				eF.A = ((obj3 != null) ? ((Link)obj3) : default(Link));
				this.A.A = Text.TextRangeParentShape(((TextLink)this.A.LinkedObject).TextRange);
			}
			else
			{
				this.A.A = Hyperlinks.LinkDetails((Microsoft.Office.Interop.PowerPoint.Hyperlink)this.A.LinkedObject);
				EF eF2 = this.A;
				Type typeFromHandle2 = typeof(Hyperlinks);
				string memberName2 = AH.A(98485);
				ExcelLinkItem excelLinkItem;
				object[] obj4 = new object[2]
				{
					(excelLinkItem = this.A).LinkedObject,
					false
				};
				object[] array = obj4;
				bool[] obj5 = new bool[2] { true, false };
				bool[] array2 = obj5;
				object obj6 = NewLateBinding.LateGet(null, typeFromHandle2, memberName2, obj4, null, null, obj5);
				if (array2[0])
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
					excelLinkItem.LinkedObject = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(array[0]));
				}
				eF2.A = (Microsoft.Office.Interop.PowerPoint.Shape)obj6;
			}
			this.A.Link = this.A.A;
			this.A.LinkedShape = this.A.A;
		}
	}

	[CompilerGenerated]
	internal sealed class EF
	{
		public object A;

		public List<string> A;

		public CopierAsPicture A;

		public TimelineRestorer A;

		public Link A;

		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public wpfManageLinks A;

		public EF(EF A)
		{
			//IL_004e: Unknown result type (might be due to invalid IL or missing references)
			//IL_0053: Unknown result type (might be due to invalid IL or missing references)
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
				this.A = A.A;
				this.A = A.A;
				this.A = A.A;
				this.A = A.A;
				this.A = A.A;
				return;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class FF
	{
		public UpdateLinkException A;

		public DF A;

		public FF(FF A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal void A()
		{
			((LinkItem)this.A.A).MarkBroken(((Exception)(object)this.A).Message);
		}
	}

	[CompilerGenerated]
	internal sealed class GF
	{
		public Exception A;

		public DF A;

		public GF(GF A)
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
			((LinkItem)this.A.A).MarkBroken(this.A.Message);
		}
	}

	[CompilerGenerated]
	internal sealed class HF
	{
		public Exception A;

		public wpfManageLinks A;

		public HF(HF A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A.D(this.A.Message);
		}
	}

	[CompilerGenerated]
	internal sealed class IF
	{
		public string A;

		public string B;

		public bool A;

		public wpfManageLinks A;

		public IF(IF A)
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
				B = A.B;
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal void A()
		{
			this.A = this.A.txtFind.Text;
			B = this.A.txtReplace.Text;
			this.A = this.A.chkRegex.IsChecked.Value;
		}
	}

	[CompilerGenerated]
	internal sealed class JF
	{
		public TextBlock A;

		public string A;

		[SpecialName]
		internal void A()
		{
			this.A.Text = this.A;
		}
	}

	[CompilerGenerated]
	internal sealed class KF
	{
		public LinkItem A;

		public string A;

		[SpecialName]
		internal void A()
		{
			((LinkItem)this.A).MarkBroken(this.A);
		}
	}

	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	private GridViewColumnHeader m_A;

	private SortAdorner m_A;

	private ICollectionView m_A;

	private Thickness m_A;

	private Microsoft.Office.Interop.PowerPoint.Application m_A;

	private Microsoft.Office.Interop.PowerPoint.Presentation m_A;

	private ObservableCollection<ExcelLinkItem> m_A;

	private ObservableCollection<ExcelLinkItem> m_B;

	private BackgroundWorker m_A;

	private bool m_A;

	private RefreshInstance m_A;

	private IEnumerable<Microsoft.Office.Interop.Excel.Application> m_A;

	private List<object> m_A;

	private int m_A;

	private int m_B;

	private bool m_B;

	private FileInfo[] m_A;

	private int m_C;

	private int m_D;

	private List<CF> m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("TabControl1")]
	private System.Windows.Controls.TabControl m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("tabShapes")]
	private TabItem m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lvShapes")]
	private System.Windows.Controls.ListView m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("gvShapes")]
	private GridView m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkShapes")]
	private System.Windows.Controls.CheckBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterShapesToggle")]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("btnUpdateShape")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnVerifyExcel")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[AccessedThroughProperty("btnEditShape")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_C;

	[AccessedThroughProperty("btnUnlinkShape")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_D;

	[AccessedThroughProperty("btnViewShape")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_E;

	[AccessedThroughProperty("btnExportLinks")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_F;

	[AccessedThroughProperty("pbShapes")]
	[CompilerGenerated]
	private System.Windows.Controls.ProgressBar m_A;

	[AccessedThroughProperty("tbShapeCount")]
	[CompilerGenerated]
	private TextBlock m_A;

	[AccessedThroughProperty("grdShapeFilters")]
	[CompilerGenerated]
	private Grid m_A;

	[AccessedThroughProperty("cbxFilterShapeSource")]
	[CompilerGenerated]
	private System.Windows.Controls.ComboBox m_A;

	[AccessedThroughProperty("cbxFilterModifiedBy")]
	[CompilerGenerated]
	private System.Windows.Controls.ComboBox m_B;

	[AccessedThroughProperty("chkFilterRanges")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_C;

	[AccessedThroughProperty("chkFilterCharts")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterTypeGraphic")]
	private System.Windows.Controls.CheckBox m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterTypePicture")]
	private System.Windows.Controls.CheckBox m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterTypeTable")]
	private System.Windows.Controls.CheckBox m_G;

	[AccessedThroughProperty("chkFilterTypeWorkbook")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox H;

	[AccessedThroughProperty("chkFilterTypeText")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox I;

	[AccessedThroughProperty("chkFilterTypeChart")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox J;

	[CompilerGenerated]
	[AccessedThroughProperty("btnReset")]
	private System.Windows.Controls.Button m_G;

	[CompilerGenerated]
	[AccessedThroughProperty("grpFindReplace")]
	private System.Windows.Controls.GroupBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtFind")]
	private System.Windows.Controls.TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtReplace")]
	private System.Windows.Controls.TextBox m_B;

	[AccessedThroughProperty("grpScope")]
	[CompilerGenerated]
	private System.Windows.Controls.GroupBox m_B;

	[AccessedThroughProperty("radThisPresentation")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[AccessedThroughProperty("radAllPresentations")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("txtFolder")]
	private System.Windows.Controls.TextBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnBrowse")]
	private System.Windows.Controls.Button H;

	[CompilerGenerated]
	[AccessedThroughProperty("chkSubfolders")]
	private System.Windows.Controls.CheckBox K;

	[AccessedThroughProperty("btnReplace")]
	[CompilerGenerated]
	private System.Windows.Controls.Button I;

	[AccessedThroughProperty("btnStop")]
	[CompilerGenerated]
	private System.Windows.Controls.Button J;

	[AccessedThroughProperty("chkRegex")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox L;

	[CompilerGenerated]
	[AccessedThroughProperty("stkFindReplace")]
	private StackPanel m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("pbFindReplace")]
	private System.Windows.Controls.ProgressBar m_B;

	[AccessedThroughProperty("lblReplacing")]
	[CompilerGenerated]
	private TextBlock m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private System.Windows.Controls.Button K;

	private bool m_C;

	public ICollectionView ShapesCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(96024));
		}
	}

	public Thickness AdornerPadding
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(96057));
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
			this.m_A = value;
		}
	}

	internal virtual TabItem tabShapes
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

	internal virtual System.Windows.Controls.ListView lvShapes
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
			System.Windows.Input.KeyEventHandler value2 = SpacebarToggleShapes;
			System.Windows.Controls.ListView listView = this.m_A;
			if (listView != null)
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
				listView.PreviewKeyDown -= value2;
			}
			this.m_A = value;
			listView = this.m_A;
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
				listView.PreviewKeyDown += value2;
				return;
			}
		}
	}

	internal virtual GridView gvShapes
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

	internal virtual System.Windows.Controls.CheckBox chkShapes
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

	internal virtual System.Windows.Controls.CheckBox chkFilterShapesToggle
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
			RoutedEventHandler value2 = ToggleShapeFilters;
			RoutedEventHandler value3 = ToggleShapeFilters;
			System.Windows.Controls.CheckBox checkBox = this.m_B;
			if (checkBox != null)
			{
				checkBox.Checked -= value2;
				checkBox.Unchecked -= value3;
			}
			this.m_B = value;
			checkBox = this.m_B;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				checkBox.Checked += value2;
				checkBox.Unchecked += value3;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnUpdateShape
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
			RoutedEventHandler value2 = btnUpdateShape_Click;
			System.Windows.Controls.Button button = this.m_A;
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

	internal virtual System.Windows.Controls.Button btnVerifyExcel
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
			RoutedEventHandler value2 = btnVerifyExcel_Click;
			System.Windows.Controls.Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
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

	internal virtual System.Windows.Controls.Button btnEditShape
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
			RoutedEventHandler value2 = btnEditShape_Click;
			System.Windows.Controls.Button button = this.m_C;
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
			this.m_C = value;
			button = this.m_C;
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.Button btnUnlinkShape
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
			RoutedEventHandler value2 = btnUnlinkShape_Click;
			System.Windows.Controls.Button button = this.m_D;
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
			this.m_D = value;
			button = this.m_D;
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

	internal virtual System.Windows.Controls.Button btnViewShape
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
			RoutedEventHandler value2 = btnViewShape_Click;
			System.Windows.Controls.Button button = this.m_E;
			if (button != null)
			{
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

	internal virtual System.Windows.Controls.Button btnExportLinks
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
			RoutedEventHandler value2 = btnExportLinks_Click;
			System.Windows.Controls.Button button = this.m_F;
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

	internal virtual System.Windows.Controls.ProgressBar pbShapes
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

	internal virtual TextBlock tbShapeCount
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

	internal virtual Grid grdShapeFilters
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

	internal virtual System.Windows.Controls.ComboBox cbxFilterShapeSource
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

	internal virtual System.Windows.Controls.ComboBox cbxFilterModifiedBy
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

	internal virtual System.Windows.Controls.CheckBox chkFilterRanges
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

	internal virtual System.Windows.Controls.CheckBox chkFilterCharts
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

	internal virtual System.Windows.Controls.CheckBox chkFilterTypeGraphic
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

	internal virtual System.Windows.Controls.CheckBox chkFilterTypePicture
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

	internal virtual System.Windows.Controls.CheckBox chkFilterTypeTable
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

	internal virtual System.Windows.Controls.CheckBox chkFilterTypeWorkbook
	{
		[CompilerGenerated]
		get
		{
			return this.H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.H = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterTypeText
	{
		[CompilerGenerated]
		get
		{
			return this.I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.I = value;
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkFilterTypeChart
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
			this.J = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnReset
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
			RoutedEventHandler value2 = btnReset_Click;
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

	internal virtual System.Windows.Controls.GroupBox grpFindReplace
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

	internal virtual System.Windows.Controls.TextBox txtReplace
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

	internal virtual System.Windows.Controls.GroupBox grpScope
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

	internal virtual System.Windows.Controls.RadioButton radThisPresentation
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

	internal virtual System.Windows.Controls.RadioButton radAllPresentations
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

	internal virtual System.Windows.Controls.TextBox txtFolder
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

	internal virtual System.Windows.Controls.Button btnBrowse
	{
		[CompilerGenerated]
		get
		{
			return H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnBrowse_Click;
			System.Windows.Controls.Button button = H;
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
			H = value;
			button = H;
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
				button.Click += value2;
				return;
			}
		}
	}

	internal virtual System.Windows.Controls.CheckBox chkSubfolders
	{
		[CompilerGenerated]
		get
		{
			return this.K;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.K = value;
		}
	}

	internal virtual System.Windows.Controls.Button btnReplace
	{
		[CompilerGenerated]
		get
		{
			return I;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnReplace_Click;
			System.Windows.Controls.Button button = I;
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
			I = value;
			button = I;
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

	internal virtual System.Windows.Controls.Button btnStop
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
			RoutedEventHandler value2 = btnStop_Click;
			System.Windows.Controls.Button button = J;
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
			J = value;
			button = J;
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

	internal virtual System.Windows.Controls.CheckBox chkRegex
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
			L = value;
		}
	}

	internal virtual StackPanel stkFindReplace
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

	internal virtual System.Windows.Controls.ProgressBar pbFindReplace
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

	internal virtual TextBlock lblReplacing
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

	internal virtual System.Windows.Controls.Button btnClose
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
			RoutedEventHandler value2 = btnClose_Click;
			System.Windows.Controls.Button button = K;
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
			K = value;
			button = K;
			if (button != null)
			{
				button.Click += value2;
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

	public wpfManageLinks()
	{
		base.Loaded += wpfManageLinks_Loaded;
		base.Closing += wpfManageLinks_Closing;
		this.m_A = null;
		this.m_A = null;
		this.m_A = false;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
	}

	private void A(string A)
	{
		this.m_A?.Invoke(this, new PropertyChangedEventArgs(A));
	}

	private void wpfManageLinks_Loaded(object sender, RoutedEventArgs e)
	{
		int num = 1;
		List<string> D = new List<string>();
		List<string> E = new List<string>();
		this.m_A = NG.A.Application;
		this.m_A = this.m_A.ActivePresentation;
		this.m_A.ActiveWindow.ViewType = PpViewType.ppViewNormal;
		this.m_A = new ObservableCollection<ExcelLinkItem>();
		bool flag = this.m_A.SectionProperties.Count > 0;
		string text = "";
		D.Add(AH.A(96086));
		E.Add(AH.A(96086));
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			ExcelLinkItem excelLinkItem;
			try
			{
				enumerator = this.m_A.Slides.GetEnumerator();
				IEnumerator enumerator2 = default(IEnumerator);
				IEnumerator enumerator4 = default(IEnumerator);
				while (enumerator.MoveNext())
				{
					Slide slide = (Slide)enumerator.Current;
					if (flag)
					{
						SectionProperties sectionProperties = this.m_A.SectionProperties;
						int num2 = num;
						int count = sectionProperties.Count;
						int num3 = num2;
						while (true)
						{
							if (num3 <= count)
							{
								if (slide.SlideIndex >= sectionProperties.FirstSlide(num3))
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
									if (slide.SlideIndex < sectionProperties.FirstSlide(num3) + sectionProperties.SlidesCount(num3))
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
										text = sectionProperties.Name(num3);
										num++;
										break;
									}
								}
								num3++;
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
							break;
						}
						sectionProperties = null;
						if (text.Length == 0)
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
							text = AH.A(7090);
						}
					}
					try
					{
						enumerator2 = slide.Shapes.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
							if (shape.Visible != MsoTriState.msoTrue)
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
							A(shape, slide, text, ref D, ref E);
							if (!Text.ContainsLinks(shape))
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
							using List<TextLink>.Enumerator enumerator3 = Text.SelectedLinks(shape).GetEnumerator();
							while (enumerator3.MoveNext())
							{
								TextLink current = enumerator3.Current;
								excelLinkItem = A(slide, current, text);
								if (excelLinkItem == null)
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
								this.m_A.Add(excelLinkItem);
								D.Add(((LinkItem)excelLinkItem).SourcePath);
								E.Add(((LinkItem)excelLinkItem).ModifiedBy);
							}
							while (true)
							{
								switch (1)
								{
								case 0:
									break;
								default:
									goto end_IL_024f;
								}
								continue;
								end_IL_024f:
								break;
							}
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_0275;
							}
							continue;
							end_IL_0275:
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
					try
					{
						enumerator4 = slide.Hyperlinks.GetEnumerator();
						while (enumerator4.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Hyperlink hyperlink = (Microsoft.Office.Interop.PowerPoint.Hyperlink)enumerator4.Current;
							if (!Hyperlinks.IsLinked(hyperlink))
							{
								continue;
							}
							excelLinkItem = A(slide, hyperlink, text);
							if (excelLinkItem == null)
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
							this.m_A.Add(excelLinkItem);
							D.Add(((LinkItem)excelLinkItem).SourcePath);
							E.Add(((LinkItem)excelLinkItem).ModifiedBy);
						}
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								goto end_IL_031e;
							}
							continue;
							end_IL_031e:
							break;
						}
					}
					finally
					{
						if (enumerator4 is IDisposable)
						{
							while (true)
							{
								switch (7)
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
				while (true)
				{
					switch (5)
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
			excelLinkItem = null;
			this.m_B = new ObservableCollection<ExcelLinkItem>(this.m_A);
			ShapesCollection = CollectionViewSource.GetDefaultView(this.m_B);
			Manage2.UpdateLinkCount(lvShapes, tbShapeCount);
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
				CollectionView obj = (CollectionView)CollectionViewSource.GetDefaultView(lvShapes.ItemsSource);
				PropertyGroupDescription item = new PropertyGroupDescription(AH.A(96093));
				obj.GroupDescriptions.Add(item);
			}
			System.Windows.Controls.ComboBox comboBox = cbxFilterShapeSource;
			comboBox.ItemsSource = D.Distinct();
			comboBox.SelectedIndex = 0;
			_ = null;
			System.Windows.Controls.ComboBox comboBox2 = cbxFilterModifiedBy;
			comboBox2.ItemsSource = E.Distinct();
			comboBox2.SelectedIndex = 0;
			_ = null;
			D = null;
			E = null;
			B();
			try
			{
				if (this.m_A.ActiveWindow.Selection.HasChildShapeRange)
				{
					IEnumerator enumerator5 = default(IEnumerator);
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						try
						{
							enumerator5 = this.m_A.ActiveWindow.Selection.ChildShapeRange.GetEnumerator();
							while (enumerator5.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator5.Current;
								A(a);
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_04b7;
								}
								continue;
								end_IL_04b7:
								break;
							}
						}
						finally
						{
							if (enumerator5 is IDisposable)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										continue;
									}
									(enumerator5 as IDisposable).Dispose();
									break;
								}
							}
						}
						break;
					}
				}
				else
				{
					IEnumerator enumerator6 = default(IEnumerator);
					try
					{
						enumerator6 = this.m_A.ActiveWindow.Selection.ShapeRange.GetEnumerator();
						while (enumerator6.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape a2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator6.Current;
							A(a2);
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_0530;
							}
							continue;
							end_IL_0530:
							break;
						}
					}
					finally
					{
						if (enumerator6 is IDisposable)
						{
							while (true)
							{
								switch (4)
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
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			lvShapes.Focus();
			chkShapes.IsEnabled = lvShapes.Items.Count > 0;
			A(this.m_A);
			radAllPresentations.IsEnabled = Base.IsUserAdmin();
			base.Activated += wpfManageLinks_Activated;
			base.Deactivated += wpfManageLinks_Deactivated;
			chkShapes.Checked += chkShapes_CheckedChanged;
			chkShapes.Unchecked += chkShapes_CheckedChanged;
			lvShapes.SelectionChanged += lvShapes_SelectionChanged;
			radThisPresentation.Checked += FindReplaceScopeChanged;
			radAllPresentations.Checked += FindReplaceScopeChanged;
			new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56688)).AddEventHandler(this.m_A, new EApplication_PresentationBeforeCloseEventHandler(A));
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (A.Type != MsoShapeType.msoGroup)
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
					{
						foreach (ExcelLinkItem item in this.m_A)
						{
							if (item.LinkedShape == A)
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
								((LinkItem)item).IsChecked = true;
								B(A: true);
							}
						}
						return;
					}
				}
			}
		}
		IEnumerator enumerator2 = A.GroupItems.GetEnumerator();
		try
		{
			while (enumerator2.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
				this.A(a);
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
			IDisposable disposable = enumerator2 as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, Slide B, string C, ref List<string> D, ref List<string> E)
	{
		if (Shapes.IsLinked(A))
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					ExcelLinkItem excelLinkItem = this.A(B, A, C);
					if (excelLinkItem != null)
					{
						this.m_A.Add(excelLinkItem);
						D.Add(((LinkItem)excelLinkItem).SourcePath);
						E.Add(((LinkItem)excelLinkItem).ModifiedBy);
						excelLinkItem = null;
					}
					return;
				}
				}
			}
		}
		if (A.Type != MsoShapeType.msoGroup)
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
			enumerator = A.GroupItems.GetEnumerator();
			try
			{
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
					this.A(a, B, C, ref D, ref E);
				}
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

	private void wpfManageLinks_Closing(object sender, CancelEventArgs e)
	{
		if (this.m_A != null)
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
			if (this.m_A.IsBusy)
			{
				this.m_A.CancelAsync();
				e.Cancel = true;
				return;
			}
			this.m_A = null;
		}
		base.Activated -= wpfManageLinks_Activated;
		base.Deactivated -= wpfManageLinks_Deactivated;
		chkShapes.Checked -= chkShapes_CheckedChanged;
		chkShapes.Unchecked -= chkShapes_CheckedChanged;
		lvShapes.SelectionChanged -= lvShapes_SelectionChanged;
		radThisPresentation.Checked -= FindReplaceScopeChanged;
		radAllPresentations.Checked -= FindReplaceScopeChanged;
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56688)).RemoveEventHandler(this.m_A, new EApplication_PresentationBeforeCloseEventHandler(A));
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_B = null;
		Properties.ManageLinksHeight = base.Height;
		Properties.ManageLinksWidth = base.Width;
	}

	private void wpfManageLinks_Activated(object sender, EventArgs e)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).RemoveEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void wpfManageLinks_Deactivated(object sender, EventArgs e)
	{
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(12762)).AddEventHandler(this.m_A, new EApplication_WindowSelectionChangeEventHandler(A));
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Presentation A, ref bool B)
	{
		Close();
	}

	private void A(Selection A)
	{
		lvShapes.SelectedItems.Clear();
		try
		{
			PpSelectionType type = A.Type;
			if (type != PpSelectionType.ppSelectionShapes)
			{
				return;
			}
			lvShapes.SelectionChanged -= lvShapes_SelectionChanged;
			try
			{
				if (!this.m_A.ActiveWindow.Selection.HasChildShapeRange)
				{
					{
						IEnumerator enumerator = this.m_A.ActiveWindow.Selection.ShapeRange.GetEnumerator();
						try
						{
							while (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
								B(a);
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_0110;
								}
								continue;
								end_IL_0110:
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
				}
				else
				{
					IEnumerator enumerator2 = default(IEnumerator);
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
						enumerator2 = this.m_A.ActiveWindow.Selection.ChildShapeRange.GetEnumerator();
						try
						{
							while (enumerator2.MoveNext())
							{
								Microsoft.Office.Interop.PowerPoint.Shape a2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
								B(a2);
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									break;
								default:
									goto end_IL_00ac;
								}
								continue;
								end_IL_00ac:
								break;
							}
						}
						finally
						{
							IDisposable disposable2 = enumerator2 as IDisposable;
							if (disposable2 != null)
							{
								disposable2.Dispose();
							}
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
			lvShapes.SelectionChanged += lvShapes_SelectionChanged;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
	}

	private void B(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		if (A.Type != MsoShapeType.msoGroup)
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
					{
						foreach (ExcelLinkItem item in this.m_B)
						{
							if (A == this.A(item))
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
								((LinkItem)item).IsSelected = true;
							}
						}
						return;
					}
				}
			}
		}
		IEnumerator enumerator2 = A.GroupItems.GetEnumerator();
		try
		{
			while (enumerator2.MoveNext())
			{
				Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
				B(a);
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
			IDisposable disposable = enumerator2 as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	private void SpacebarToggleShapes(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Space)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			if (lvShapes.SelectedItems.Count <= 0)
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
				this.m_A = true;
				try
				{
					bool isChecked = !((LinkItem)(ExcelLinkItem)lvShapes.SelectedItems[0]).IsChecked;
					try
					{
						enumerator = lvShapes.SelectedItems.GetEnumerator();
						while (enumerator.MoveNext())
						{
							((LinkItem)(ExcelLinkItem)enumerator.Current).IsChecked = isChecked;
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
					A(this.m_B);
					B(A() > 0);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				this.m_A = false;
				return;
			}
		}
	}

	private void ShapeItemCheckedChanged(object sender, RoutedEventArgs e)
	{
		if (!this.m_A)
		{
			A(this.m_B);
			B(A() > 0);
		}
	}

	private void lvShapes_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		//IL_00e7: Unknown result type (might be due to invalid IL or missing references)
		if (lvShapes.SelectedItems.Count == 1)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
				{
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					btnViewShape.IsEnabled = true;
					ExcelLinkItem excelLinkItem = this.m_B[lvShapes.SelectedIndex];
					object objectValue = RuntimeHelpers.GetObjectValue(A(excelLinkItem));
					if (objectValue != null)
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
						Microsoft.Office.Interop.PowerPoint.Shape linkedShape = excelLinkItem.LinkedShape;
						base.Activated -= wpfManageLinks_Activated;
						base.Deactivated -= wpfManageLinks_Deactivated;
						try
						{
							A((Slide)linkedShape.Parent);
							if (objectValue is Microsoft.Office.Interop.PowerPoint.Shape)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										linkedShape.Select();
										goto end_IL_00c1;
									}
									continue;
									end_IL_00c1:
									break;
								}
							}
							else if (objectValue is TextLink)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										((TextLink)objectValue).TextRange.Select();
										goto end_IL_00dc;
									}
									continue;
									end_IL_00dc:
									break;
								}
							}
							else if (objectValue is Microsoft.Office.Interop.PowerPoint.Hyperlink)
							{
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										Hyperlinks.HyperlinkParentTextRange((Microsoft.Office.Interop.PowerPoint.Hyperlink)objectValue).Select();
										goto end_IL_0102;
									}
									continue;
									end_IL_0102:
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
						Focus();
						base.Activated += wpfManageLinks_Activated;
						base.Deactivated += wpfManageLinks_Deactivated;
						objectValue = null;
						linkedShape = null;
					}
					else
					{
						this.m_A.Remove(excelLinkItem);
						this.m_B.Remove(excelLinkItem);
					}
					excelLinkItem = null;
					return;
				}
				}
			}
		}
		if (lvShapes.SelectedItems.Count <= 1)
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
			btnViewShape.IsEnabled = false;
			return;
		}
	}

	private void A(Slide A)
	{
		Slide slide = null;
		try
		{
			slide = this.m_A.ActiveWindow.Selection.SlideRange[1];
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (slide == A)
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
			this.m_A.Windows[1].Activate();
			this.m_A.ActiveWindow.View.GotoSlide(A.SlideIndex);
			return;
		}
	}

	private void A(ObservableCollection<ExcelLinkItem> A)
	{
		System.Windows.Controls.CheckBox checkBox = chkShapes;
		checkBox.Checked -= chkShapes_CheckedChanged;
		checkBox.Unchecked -= chkShapes_CheckedChanged;
		Func<ExcelLinkItem, bool> predicate;
		if (_Closure_0024__.A == null)
		{
			predicate = (_Closure_0024__.A = [SpecialName] (ExcelLinkItem excelLinkItem) => ((LinkItem)excelLinkItem).IsChecked);
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
		Forms.SynchCheckBox(A.Where(predicate).Count(), lvShapes, chkShapes);
		checkBox.Checked += chkShapes_CheckedChanged;
		checkBox.Unchecked += chkShapes_CheckedChanged;
		_ = null;
	}

	private void chkShapes_CheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkShapes.IsChecked.Value;
		this.m_A = true;
		A(this.m_B, value);
		lvShapes.Focus();
		this.m_A = false;
		B(value);
	}

	private void A(ObservableCollection<ExcelLinkItem> A, bool B)
	{
		IEnumerator<ExcelLinkItem> enumerator = default(IEnumerator<ExcelLinkItem>);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				((LinkItem)enumerator.Current).IsChecked = B;
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

	private void btnVerifyExcel_Click(object sender, RoutedEventArgs e)
	{
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			A();
			return;
		}
	}

	private void A()
	{
		Manage2.GetExcelInstances(ref this.m_A, ref this.m_B);
		if (!Manage2.IsExcelReady((Action<string>)C, this.m_A))
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
			E();
			Manage2.StartProgressBar(pbShapes, tbShapeCount);
			this.m_A = new BackgroundWorker();
			BackgroundWorker a = this.m_A;
			a.WorkerSupportsCancellation = true;
			a.WorkerReportsProgress = true;
			a.DoWork += VerifyExcelLinksDoWork;
			a.ProgressChanged += VerifyExcelLinksProgressChanged;
			a.RunWorkerCompleted += VerifyExcelLinksComplete;
			a.RunWorkerAsync();
			_ = null;
			return;
		}
	}

	private void VerifyExcelLinksDoWork(object sender, DoWorkEventArgs e)
	{
		//IL_008c: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ed: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ef: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f4: Invalid comparison between Unknown and I4
		//IL_0100: Unknown result type (might be due to invalid IL or missing references)
		//IL_0104: Invalid comparison between Unknown and I4
		Workbook workbook = null;
		int num = 0;
		this.m_A = 0;
		this.m_A = new List<object>();
		ObservableCollection<ExcelLinkItem> observableCollection = B();
		this.m_B = observableCollection.Count;
		try
		{
			IEnumerator<ExcelLinkItem> enumerator = default(IEnumerator<ExcelLinkItem>);
			try
			{
				enumerator = observableCollection.GetEnumerator();
				while (true)
				{
					IL_015b:
					if (enumerator.MoveNext())
					{
						ExcelLinkItem current = enumerator.Current;
						if (this.m_A == null)
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
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
							workbook = Manage2.GetSourceWorkbook(this.m_A, current.Link, ref this.m_A);
							if (workbook == null)
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
								A(current, AH.A(96104));
								goto IL_0130;
							}
							if (!Manage2.FindExcelSource(current.Link, workbook))
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
								ImportType type = current.Link.Type;
								if (type - 6 > 2)
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
									if ((int)type != 12)
									{
										A(current, AH.A(96486));
										goto IL_012e;
									}
								}
								A(current, AH.A(96316));
							}
							goto IL_012e;
							IL_0130:
							checked
							{
								num++;
								this.m_A.ReportProgress((int)Math.Round((double)num / (double)this.m_B * 100.0));
								goto IL_015b;
							}
							IL_012e:
							workbook = null;
							goto IL_0130;
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
							goto end_IL_0166;
						}
						continue;
						end_IL_0166:
						break;
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(ex2.Message);
			ProjectData.ClearProjectError();
		}
		observableCollection = null;
	}

	private void VerifyExcelLinksProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbShapes.Value = e.ProgressPercentage;
	}

	private void VerifyExcelLinksComplete(object sender, RunWorkerCompletedEventArgs e)
	{
		Manage2.StopProgressBar(pbShapes, tbShapeCount);
		Manage2.ClosePreviouslyClosedWorkbooks(ref this.m_A, this.m_B);
		ReleaseHelper.ReleaseObjectEnumerable<Microsoft.Office.Interop.Excel.Application>(ref this.m_A, false);
		ReleaseHelper.DoGarbageCollection();
		if (!e.Cancelled)
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
			if (this.m_A > 0)
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
				C(Manage2.LinkValidationFailedMessage());
			}
			else
			{
				B(Manage2.LinkValidationSuccessMessage(this.m_B));
			}
		}
		this.m_A = 0;
		this.m_B = 0;
		this.m_B = false;
	}

	private void btnEditShape_Click(object sender, RoutedEventArgs e)
	{
		//IL_023b: Unknown result type (might be due to invalid IL or missing references)
		//IL_0240: Unknown result type (might be due to invalid IL or missing references)
		//IL_0242: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a7: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ac: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_0245: Unknown result type (might be due to invalid IL or missing references)
		//IL_022d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0232: Unknown result type (might be due to invalid IL or missing references)
		//IL_0223: Unknown result type (might be due to invalid IL or missing references)
		//IL_0229: Unknown result type (might be due to invalid IL or missing references)
		int num = 0;
		if (A() || !Manage2.IsAllExcelReady((Action<string>)C))
		{
			return;
		}
		checked
		{
			IEnumerator<ExcelLinkItem> enumerator = default(IEnumerator<ExcelLinkItem>);
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
				List<object> list = new List<object>();
				ObservableCollection<ExcelLinkItem> observableCollection = B();
				E();
				for (int i = observableCollection.Count - 1; i >= 0; i += -1)
				{
					ExcelLinkItem excelLinkItem = observableCollection[i];
					object objectValue = RuntimeHelpers.GetObjectValue(A(excelLinkItem));
					if (objectValue != null)
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
						list.Add(RuntimeHelpers.GetObjectValue(objectValue));
					}
					else
					{
						this.m_A.Remove(excelLinkItem);
						this.m_B.Remove(excelLinkItem);
					}
					excelLinkItem = null;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					object objectValue;
					Shapes.EditedShapes editedShapes;
					if (list.Any())
					{
						list.Reverse();
						base.Topmost = false;
						editedShapes = Shapes.EditLink(list);
						base.Topmost = true;
						if (list.Count == editedShapes.Objects.Count)
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
							if (editedShapes.IsError != null)
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
								try
								{
									enumerator = observableCollection.GetEnumerator();
									while (enumerator.MoveNext())
									{
										ExcelLinkItem current = enumerator.Current;
										if (editedShapes.IsError[num])
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
											((LinkItem)current).MarkBroken("");
										}
										else
										{
											objectValue = RuntimeHelpers.GetObjectValue(editedShapes.Objects[num]);
											Link link;
											if (objectValue is Microsoft.Office.Interop.PowerPoint.Shape)
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
												link = Shapes.LinkDetails((Microsoft.Office.Interop.PowerPoint.Shape)objectValue);
											}
											else if (objectValue is TextLink)
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
												Type typeFromHandle = typeof(Text);
												string memberName = AH.A(93278);
												object[] obj = new object[1] { objectValue };
												object[] array = obj;
												bool[] obj2 = new bool[1] { true };
												bool[] array2 = obj2;
												object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
												if (array2[0])
												{
													objectValue = RuntimeHelpers.GetObjectValue(array[0]);
												}
												_003F val;
												if (obj3 == null)
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
													val = default(Link);
												}
												else
												{
													val = (Link)obj3;
												}
												link = (Link)val;
											}
											else
											{
												link = Hyperlinks.LinkDetails((Microsoft.Office.Interop.PowerPoint.Hyperlink)objectValue);
											}
											current.Link = link;
											current.LinkedObject = RuntimeHelpers.GetObjectValue(objectValue);
										}
										num++;
									}
									while (true)
									{
										switch (1)
										{
										case 0:
											break;
										default:
											goto end_IL_026e;
										}
										continue;
										end_IL_026e:
										break;
									}
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
							F();
						}
						else
						{
							D(AH.A(96634));
						}
					}
					objectValue = null;
					editedShapes = default(Shapes.EditedShapes);
					list = null;
					observableCollection = null;
					return;
				}
			}
		}
	}

	private void btnViewShape_Click(object sender, RoutedEventArgs e)
	{
		//IL_01ee: Unknown result type (might be due to invalid IL or missing references)
		//IL_01f3: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f6: Unknown result type (might be due to invalid IL or missing references)
		if (!Access.AllowSuiteOperation((PlanType)5, (Restriction)2, false) || A())
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
			if (lvShapes.SelectedItems.Count > 0)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
					{
						object objectValue;
						ExcelLinkItem excelLinkItem;
						try
						{
							excelLinkItem = (ExcelLinkItem)lvShapes.SelectedItems[0];
							objectValue = RuntimeHelpers.GetObjectValue(A(excelLinkItem));
							if (objectValue != null)
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
									{
										string sourcePath;
										if (objectValue is Microsoft.Office.Interop.PowerPoint.Shape)
										{
											Type typeFromHandle = typeof(Shapes);
											string memberName = AH.A(96737);
											object[] obj = new object[1] { objectValue };
											object[] array = obj;
											bool[] obj2 = new bool[1] { true };
											bool[] array2 = obj2;
											NewLateBinding.LateCall(null, typeFromHandle, memberName, obj, null, null, obj2, IgnoreReturn: true);
											if (array2[0])
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
												objectValue = RuntimeHelpers.GetObjectValue(array[0]);
											}
											sourcePath = Shapes.LinkDetails((Microsoft.Office.Interop.PowerPoint.Shape)objectValue).Source;
										}
										else if (objectValue is TextLink)
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
											Type typeFromHandle2 = typeof(Text);
											string memberName2 = AH.A(96737);
											object[] obj3 = new object[1] { objectValue };
											object[] array = obj3;
											bool[] obj4 = new bool[1] { true };
											bool[] array2 = obj4;
											NewLateBinding.LateCall(null, typeFromHandle2, memberName2, obj3, null, null, obj4, IgnoreReturn: true);
											if (array2[0])
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
												objectValue = RuntimeHelpers.GetObjectValue(array[0]);
											}
											object instance = NewLateBinding.LateGet(null, typeof(Text), AH.A(93278), array = new object[1] { objectValue }, null, null, array2 = new bool[1] { true });
											if (array2[0])
											{
												objectValue = RuntimeHelpers.GetObjectValue(array[0]);
											}
											sourcePath = Conversions.ToString(NewLateBinding.LateGet(instance, null, AH.A(96758), new object[0], null, null, null));
										}
										else
										{
											Hyperlinks.ViewSource((Microsoft.Office.Interop.PowerPoint.Hyperlink)objectValue);
											sourcePath = Hyperlinks.LinkDetails((Microsoft.Office.Interop.PowerPoint.Hyperlink)objectValue).Source;
										}
										((LinkItem)excelLinkItem).SourcePath = sourcePath;
										F();
										goto end_IL_0085;
									}
									}
									continue;
									end_IL_0085:
									break;
								}
							}
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							A();
							Interaction.AppActivate(this.m_A.Caption);
							ProjectData.ClearProjectError();
						}
						objectValue = null;
						excelLinkItem = null;
						return;
					}
					}
				}
			}
			C(AH.A(96771));
			return;
		}
	}

	private void btnUpdateShape_Click(object sender, RoutedEventArgs e)
	{
		//IL_006e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0074: Expected O, but got Unknown
		//IL_0075: Expected O, but got Unknown
		//IL_0043: Unknown result type (might be due to invalid IL or missing references)
		//IL_004d: Expected O, but got Unknown
		if (!Access.AllowSuiteOperation((PlanType)5, (Restriction)2, false))
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
			if (A())
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
				try
				{
					this.m_A = new RefreshInstance(System.Windows.Window.GetWindow(this));
					this.m_A.LegalFonts = BrandCompliance.GetLegalFontTypes(this.m_A.ActivePresentation);
				}
				catch (UpdateLinkException ex)
				{
					ProjectData.SetProjectError((Exception)ex);
					UpdateLinkException ex2 = ex;
					C(((Exception)(object)ex2).Message);
					ProjectData.ClearProjectError();
					return;
				}
				E();
				Manage2.StartProgressBar(pbShapes, tbShapeCount);
				base.Activated -= wpfManageLinks_Activated;
				base.Deactivated -= wpfManageLinks_Deactivated;
				this.m_A = new BackgroundWorker();
				BackgroundWorker a = this.m_A;
				a.WorkerSupportsCancellation = true;
				a.WorkerReportsProgress = true;
				a.DoWork += RefreshShapeLinksDoWork;
				a.ProgressChanged += RefreshShapeLinksProgressChanged;
				a.RunWorkerCompleted += RefreshShapeLinksComplete;
				a.RunWorkerAsync();
				_ = null;
				return;
			}
		}
	}

	private void RefreshShapeLinksDoWork(object sender, DoWorkEventArgs e)
	{
		//IL_0043: Unknown result type (might be due to invalid IL or missing references)
		//IL_004d: Expected O, but got Unknown
		//IL_013d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0143: Expected O, but got Unknown
		//IL_0157: Expected O, but got Unknown
		EF eF = new EF(eF);
		eF.A = this;
		eF.A = new List<string>();
		int num = 0;
		this.m_A = 0;
		ObservableCollection<ExcelLinkItem> observableCollection = B();
		this.m_B = observableCollection.Count;
		this.m_A.StartNewUndoEntry();
		eF.A = new CopierAsPicture();
		eF.A = new TimelineRestorer();
		checked
		{
			IEnumerator<ExcelLinkItem> enumerator = default(IEnumerator<ExcelLinkItem>);
			try
			{
				enumerator = observableCollection.GetEnumerator();
				DF dF = default(DF);
				FF fF = default(FF);
				GF gF = default(GF);
				while (enumerator.MoveNext())
				{
					dF = new DF(dF);
					dF.A = eF;
					dF.A = enumerator.Current;
					if (this.m_A == null)
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (this.m_A.CancellationPending)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								e.Cancel = true;
								break;
							}
							break;
						}
						if (this.m_A.Canceled)
						{
							while (true)
							{
								switch (2)
								{
								case 0:
									break;
								default:
									goto end_IL_00d4;
								}
								continue;
								end_IL_00d4:
								break;
							}
							break;
						}
						dF.A.A = RuntimeHelpers.GetObjectValue(A(dF.A));
						if (dF.A.A != null)
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
							try
							{
								base.Dispatcher.Invoke(dF.A);
							}
							catch (UpdateLinkException ex)
							{
								ProjectData.SetProjectError((Exception)ex);
								fF = new FF(fF);
								fF.A = dF;
								UpdateLinkException a = ex;
								fF.A = a;
								this.m_A++;
								base.Dispatcher.Invoke(fF.A);
								fF.A.A.A = null;
								ProjectData.ClearProjectError();
							}
							catch (Exception ex2)
							{
								ProjectData.SetProjectError(ex2);
								gF = new GF(gF);
								gF.A = dF;
								Exception a2 = ex2;
								gF.A = a2;
								this.m_A++;
								base.Dispatcher.Invoke(gF.A);
								gF.A.A.A = null;
								ProjectData.ClearProjectError();
							}
							dF.A.A = null;
						}
						num++;
						this.m_A.ReportProgress((int)Math.Round((double)num / (double)this.m_B * 100.0));
						goto IL_023f;
					}
					break;
					IL_023f:;
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
			eF.A.A();
			observableCollection = null;
			eF.A = null;
		}
	}

	private void RefreshShapeLinksProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbShapes.Value = e.ProgressPercentage;
	}

	private void RefreshShapeLinksComplete(object sender, RunWorkerCompletedEventArgs e)
	{
		Focus();
		Manage2.StopProgressBar(pbShapes, tbShapeCount);
		F();
		Base.ReleaseRefreshInstance(ref this.m_A, true);
		lvShapes.SelectionChanged -= lvShapes_SelectionChanged;
		lvShapes.SelectedItems.Clear();
		lvShapes.SelectionChanged += lvShapes_SelectionChanged;
		btnViewShape.IsEnabled = false;
		if (!e.Cancelled)
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
			if (this.m_A == 0)
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
				E(Manage2.LinkRefreshSuccessMessage(this.m_B));
			}
			else
			{
				D(Manage2.LinkRefreshFailedMessage(this.m_A));
			}
		}
		base.Activated += wpfManageLinks_Activated;
		base.Deactivated += wpfManageLinks_Deactivated;
		this.m_A = 0;
		this.m_B = 0;
		this.m_B = false;
	}

	private void btnUnlinkShape_Click(object sender, RoutedEventArgs e)
	{
		if (A())
		{
			return;
		}
		ObservableCollection<ExcelLinkItem> observableCollection = B();
		checked
		{
			if (observableCollection.Any())
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
				if (Base.ConfirmBreakLink())
				{
					lvShapes.SelectionChanged -= lvShapes_SelectionChanged;
					this.m_A.StartNewUndoEntry();
					for (int i = observableCollection.Count - 1; i >= 0; i += -1)
					{
						ExcelLinkItem excelLinkItem = observableCollection[i];
						object objectValue = RuntimeHelpers.GetObjectValue(A(excelLinkItem));
						if (objectValue != null)
						{
							if (objectValue is Microsoft.Office.Interop.PowerPoint.Shape)
							{
								Shapes.BreakLink((Microsoft.Office.Interop.PowerPoint.Shape)objectValue);
							}
							else if (objectValue is TextLink)
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
								Type typeFromHandle = typeof(Text);
								string memberName = AH.A(96806);
								object[] obj = new object[1] { objectValue };
								object[] array = obj;
								bool[] obj2 = new bool[1] { true };
								bool[] array2 = obj2;
								NewLateBinding.LateCall(null, typeFromHandle, memberName, obj, null, null, obj2, IgnoreReturn: true);
								if (array2[0])
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
									objectValue = RuntimeHelpers.GetObjectValue(array[0]);
								}
							}
							else
							{
								Hyperlinks.BreakLink((Microsoft.Office.Interop.PowerPoint.Hyperlink)objectValue);
							}
							objectValue = null;
						}
						this.m_A.Remove(excelLinkItem);
						this.m_B.Remove(excelLinkItem);
						excelLinkItem = null;
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
					lvShapes.SelectionChanged += lvShapes_SelectionChanged;
					chkShapes.Unchecked -= chkShapes_CheckedChanged;
					chkShapes.IsChecked = false;
					chkShapes.Unchecked += chkShapes_CheckedChanged;
					Manage2.UpdateLinkCount(lvShapes, tbShapeCount);
					B(A: false);
					F();
				}
			}
			else
			{
				G();
			}
			observableCollection = null;
		}
	}

	private void btnExportLinks_Click(object sender, RoutedEventArgs e)
	{
		//IL_01db: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e2: Unknown result type (might be due to invalid IL or missing references)
		ObservableCollection<ExcelLinkItem> observableCollection = B();
		Microsoft.Office.Interop.Excel.Application application = null;
		IEnumerator<ExcelLinkItem> enumerator = default(IEnumerator<ExcelLinkItem>);
		if (observableCollection.Any())
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
					try
					{
						application = InstanceManagement.GetExcelInstance(true);
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
					if (application != null)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
							{
								application.ScreenUpdating = false;
								application.EnableEvents = false;
								Worksheet worksheet;
								try
								{
									worksheet = Manage2.CreateNewWorksheet(application);
									Worksheet worksheet2 = worksheet;
									((Range)worksheet2.Cells[1, 1]).Value2 = AH.A(96825);
									((Range)worksheet2.Cells[1, 2]).Value2 = AH.A(96848);
									((Range)worksheet2.Cells[1, 3]).Value2 = AH.A(96881);
									((Range)worksheet2.Cells[1, 4]).Value2 = AH.A(96916);
									((Range)worksheet2.Cells[1, 5]).Value2 = AH.A(96941);
									((Range)worksheet2.Cells[1, 6]).Value2 = AH.A(96964);
									((Range)worksheet2.Cells[1, 7]).Value2 = AH.A(97011);
									_ = null;
									int num = 2;
									try
									{
										enumerator = observableCollection.GetEnumerator();
										while (enumerator.MoveNext())
										{
											ExcelLinkItem current = enumerator.Current;
											Manage2.SetLinkType(worksheet, num, current.Link.Type);
											((Range)worksheet.Cells[num, 3]).Value2 = current.Slide.SlideIndex;
											Manage2.SetLastRefresh((Range)worksheet.Cells[num, 4], ((LinkItem)current).LastUpdate);
											((Range)worksheet.Cells[num, 5]).Value2 = ((LinkItem)current).ModifiedBy;
											((Range)worksheet.Cells[num, 6]).Value2 = Manage2.GetSourceRangeString(((LinkItem)current).SourceRange);
											((Range)worksheet.Cells[num, 7]).Value2 = ((LinkItem)current).SourcePath;
											num = checked(num + 1);
										}
										while (true)
										{
											switch (4)
											{
											case 0:
												break;
											default:
												goto end_IL_02ed;
											}
											continue;
											end_IL_02ed:
											break;
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
													break;
												default:
													enumerator.Dispose();
													goto end_IL_02fd;
												}
												continue;
												end_IL_02fd:
												break;
											}
										}
									}
									Ranges.PrepareExportedData(worksheet);
								}
								catch (Exception ex3)
								{
									ProjectData.SetProjectError(ex3);
									Exception ex4 = ex3;
									D(ex4.Message);
									ProjectData.ClearProjectError();
								}
								application.ScreenUpdating = true;
								application.EnableEvents = true;
								Interaction.AppActivate(application.Caption);
								JG.A(application);
								application = null;
								worksheet = null;
								return;
							}
							}
						}
					}
					C(AH.A(97034));
					return;
				}
			}
		}
		G();
	}

	private void ToggleShapeFilters(object sender, RoutedEventArgs e)
	{
		Manage2.ToggleFilters(chkFilterShapesToggle, grdShapeFilters);
	}

	private void B()
	{
		cbxFilterShapeSource.SelectionChanged += cbxFilterShapeSource_SelectionChanged;
		cbxFilterModifiedBy.SelectionChanged += cbxFilterModifiedBy_SelectionChanged;
		chkFilterRanges.Checked += ShapesFilterChanged;
		chkFilterRanges.Unchecked += ShapesFilterChanged;
		chkFilterCharts.Checked += ShapesFilterChanged;
		chkFilterCharts.Unchecked += ShapesFilterChanged;
		chkFilterTypeGraphic.Checked += ShapesFilterChanged;
		chkFilterTypeGraphic.Unchecked += ShapesFilterChanged;
		chkFilterTypePicture.Checked += ShapesFilterChanged;
		chkFilterTypePicture.Unchecked += ShapesFilterChanged;
		chkFilterTypeTable.Checked += ShapesFilterChanged;
		chkFilterTypeTable.Unchecked += ShapesFilterChanged;
		chkFilterTypeWorkbook.Checked += ShapesFilterChanged;
		chkFilterTypeWorkbook.Unchecked += ShapesFilterChanged;
		chkFilterTypeText.Checked += ShapesFilterChanged;
		chkFilterTypeText.Unchecked += ShapesFilterChanged;
		chkFilterTypeChart.Checked += ShapesFilterChanged;
		chkFilterTypeChart.Unchecked += ShapesFilterChanged;
	}

	private void C()
	{
		cbxFilterShapeSource.SelectionChanged -= cbxFilterShapeSource_SelectionChanged;
		cbxFilterModifiedBy.SelectionChanged -= cbxFilterModifiedBy_SelectionChanged;
		chkFilterRanges.Checked -= ShapesFilterChanged;
		chkFilterRanges.Unchecked -= ShapesFilterChanged;
		chkFilterCharts.Checked -= ShapesFilterChanged;
		chkFilterCharts.Unchecked -= ShapesFilterChanged;
		chkFilterTypeGraphic.Checked -= ShapesFilterChanged;
		chkFilterTypeGraphic.Unchecked -= ShapesFilterChanged;
		chkFilterTypePicture.Checked -= ShapesFilterChanged;
		chkFilterTypePicture.Unchecked -= ShapesFilterChanged;
		chkFilterTypeTable.Checked -= ShapesFilterChanged;
		chkFilterTypeTable.Unchecked -= ShapesFilterChanged;
		chkFilterTypeWorkbook.Checked -= ShapesFilterChanged;
		chkFilterTypeWorkbook.Unchecked -= ShapesFilterChanged;
		chkFilterTypeText.Checked -= ShapesFilterChanged;
		chkFilterTypeText.Unchecked -= ShapesFilterChanged;
		chkFilterTypeChart.Checked -= ShapesFilterChanged;
		chkFilterTypeChart.Unchecked -= ShapesFilterChanged;
	}

	private void cbxFilterShapeSource_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		D();
	}

	private void cbxFilterModifiedBy_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		D();
	}

	private void ShapesFilterChanged(object sender, RoutedEventArgs e)
	{
		D();
	}

	private void D()
	{
		this.m_A = true;
		lvShapes.SelectionChanged -= lvShapes_SelectionChanged;
		this.m_B = A();
		ShapesCollection = CollectionViewSource.GetDefaultView(this.m_B);
		lvShapes.SelectionChanged += lvShapes_SelectionChanged;
		this.m_A = false;
		Manage2.UpdateLinkCount(lvShapes, tbShapeCount);
		A(this.m_B);
		F();
	}

	private ObservableCollection<ExcelLinkItem> A()
	{
		List<ExcelLinkItem> list = this.m_A.Where([SpecialName] (ExcelLinkItem A) =>
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0006: Unknown result type (might be due to invalid IL or missing references)
			bool num = Manage2.FilterLinks(A.Link, cbxFilterShapeSource, cbxFilterModifiedBy, chkFilterRanges, chkFilterCharts, chkFilterTypeGraphic, chkFilterTypePicture, chkFilterTypeTable, chkFilterTypeWorkbook, chkFilterTypeChart, chkFilterTypeText);
			if (!num)
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
				((LinkItem)A).IsChecked = false;
			}
			return num;
		}).ToList();
		return new ObservableCollection<ExcelLinkItem>(list);
	}

	private bool A(Link A, System.Windows.Controls.ComboBox B)
	{
		//IL_001b: Unknown result type (might be due to invalid IL or missing references)
		if (B.SelectedIndex != 0)
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
					return Operators.CompareString(A.Source, B.SelectedValue.ToString(), TextCompare: false) == 0;
				}
			}
		}
		return true;
	}

	private void btnReset_Click(object sender, RoutedEventArgs e)
	{
		C();
		cbxFilterShapeSource.SelectedIndex = 0;
		cbxFilterModifiedBy.SelectedIndex = 0;
		chkFilterRanges.IsChecked = true;
		chkFilterCharts.IsChecked = true;
		chkFilterTypeChart.IsChecked = true;
		chkFilterTypeGraphic.IsChecked = true;
		chkFilterTypePicture.IsChecked = true;
		chkFilterTypeTable.IsChecked = true;
		chkFilterTypeText.IsChecked = true;
		chkFilterTypeWorkbook.IsChecked = true;
		B();
		D();
	}

	private void FindReplaceScopeChanged(object sender, RoutedEventArgs e)
	{
		if (radThisPresentation.IsChecked == true)
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
					A(A: false);
					txtFolder.Text = "";
					chkSubfolders.IsChecked = false;
					return;
				}
			}
		}
		A(A: true);
	}

	private void A(bool A)
	{
		txtFolder.IsEnabled = A;
		btnBrowse.IsEnabled = A;
		chkSubfolders.IsEnabled = A;
	}

	private void btnBrowse_Click(object sender, RoutedEventArgs e)
	{
		Microsoft.Office.Core.FileDialog fileDialog = ((Microsoft.Office.Interop.PowerPoint._Application)this.m_A).get_FileDialog(MsoFileDialogType.msoFileDialogFolderPicker);
		fileDialog.Title = AH.A(97093);
		fileDialog.Filters.Clear();
		if (this.m_A.Presentations.Count > 0)
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
			fileDialog.InitialFileName = this.m_A.ActivePresentation.Path;
		}
		fileDialog.AllowMultiSelect = false;
		base.Topmost = false;
		fileDialog.Show();
		base.Topmost = true;
		FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
		if (selectedItems.Count > 0)
		{
			txtFolder.Text = selectedItems.Cast<object>().ElementAtOrDefault(0).ToString();
		}
		fileDialog = null;
		selectedItems = null;
	}

	private void btnReplace_Click(object sender, RoutedEventArgs e)
	{
		string text = txtFolder.Text;
		if (txtFind.Text.Length < 5)
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
			if (Forms.OkCancelMessage(AH.A(97120)) == System.Windows.Forms.DialogResult.Cancel)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						txtFind.Focus();
						txtFind.SelectAll();
						return;
					}
				}
			}
		}
		if (txtFind.Text.Length != 0)
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
			if (txtReplace.Text.Length != 0)
			{
				bool? isChecked = radAllPresentations.IsChecked;
				if ((isChecked ?? true) && text.Length == 0)
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
					if (isChecked.HasValue)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								C(AH.A(97455));
								return;
							}
						}
					}
				}
				this.m_C = 0;
				this.m_D = 0;
				if (radThisPresentation.IsChecked == true)
				{
					if (this.m_A.Presentations.Count > 0)
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
						if (!this.m_A.ActivePresentation.Final)
						{
							this.m_A.StartNewUndoEntry();
							A(this.m_A.ActivePresentation);
							E(AH.A(97498) + this.m_D + AH.A(97515));
						}
						else
						{
							C(Common.A);
						}
					}
					else
					{
						C(AH.A(97576));
					}
				}
				else if (Manage2.IsFolderValid(text, (Action<string>)C) && Forms.OkCancelMessage(AH.A(97623)) == System.Windows.Forms.DialogResult.OK)
				{
					this.m_A = Manage2.FindReplaceFiles(txtFolder, chkSubfolders, AH.A(70805));
					if (this.m_A.Count() > 0)
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
						grpFindReplace.IsEnabled = false;
						grpScope.IsEnabled = false;
						btnReplace.IsEnabled = false;
						btnStop.IsEnabled = true;
						pbFindReplace.Maximum = 100.0;
						pbFindReplace.Value = 0.0;
						lblReplacing.Text = "";
						stkFindReplace.Visibility = Visibility.Visible;
						new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56688)).RemoveEventHandler(this.m_A, new EApplication_PresentationBeforeCloseEventHandler(A));
						this.m_A = new BackgroundWorker();
						BackgroundWorker a = this.m_A;
						a.WorkerSupportsCancellation = true;
						a.WorkerReportsProgress = true;
						a.DoWork += bgw_DoWork;
						a.ProgressChanged += bgw_ProgressChanged;
						a.RunWorkerCompleted += bgw_RunWorkerCompleted;
						a.RunWorkerAsync();
						_ = null;
					}
					else
					{
						B(AH.A(97929));
					}
				}
				clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)10, AH.A(97976));
				return;
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
		C(AH.A(97390));
	}

	private void btnStop_Click(object sender, RoutedEventArgs e)
	{
		if (this.m_A.IsBusy)
		{
			this.m_A.CancelAsync();
		}
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		int num = 0;
		int num2 = this.m_A.Count();
		this.m_A = new List<CF>();
		checked
		{
			try
			{
				FileInfo[] a = this.m_A;
				int num3 = 0;
				IEnumerator enumerator = default(IEnumerator);
				while (true)
				{
					if (num3 < a.Length)
					{
						FileInfo fileInfo = a[num3];
						if (this.m_A == null)
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
							if (this.m_A.CancellationPending)
							{
								while (true)
								{
									switch (6)
									{
									case 0:
										break;
									default:
										e.Cancel = true;
										return;
									}
								}
							}
							if (!fileInfo.Name.StartsWith(AH.A(98013)))
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
								bool flag = false;
								{
									enumerator = this.m_A.Presentations.GetEnumerator();
									try
									{
										while (true)
										{
											if (enumerator.MoveNext())
											{
												Microsoft.Office.Interop.PowerPoint.Presentation presentation = (Microsoft.Office.Interop.PowerPoint.Presentation)enumerator.Current;
												try
												{
													if (Operators.CompareString(presentation.FullName, fileInfo.FullName, TextCompare: false) == 0)
													{
														A(lblReplacing, presentation.FullName);
														FindReplaceAsynch(presentation, ref e);
														flag = true;
														break;
													}
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
												switch (7)
												{
												case 0:
													break;
												default:
													goto end_IL_011b;
												}
												continue;
												end_IL_011b:
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
									Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
									try
									{
										presentation = this.m_A.Presentations.Open(fileInfo.FullName, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
										if (presentation.ReadOnly == MsoTriState.msoFalse)
										{
											while (true)
											{
												switch (5)
												{
												case 0:
													continue;
												}
												A(lblReplacing, presentation.FullName);
												FindReplaceAsynch(presentation, ref e);
												break;
											}
										}
										else
										{
											CF item = new CF
											{
												A = fileInfo.FullName,
												B = AH.A(98016)
											};
											this.m_A.Add(item);
										}
									}
									catch (Exception ex3)
									{
										ProjectData.SetProjectError(ex3);
										Exception ex4 = ex3;
										CF item = new CF
										{
											A = fileInfo.FullName,
											B = ex4.Message
										};
										this.m_A.Add(item);
										ProjectData.ClearProjectError();
									}
									finally
									{
										if (presentation != null)
										{
											while (true)
											{
												switch (2)
												{
												case 0:
													continue;
												}
												presentation.Close();
												presentation = null;
												break;
											}
										}
									}
								}
								num++;
								this.m_A.ReportProgress((int)Math.Round((double)num / (double)num2 * 100.0));
							}
							num3++;
							break;
						}
						continue;
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
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				HF a2 = default(HF);
				HF CS_0024_003C_003E8__locals5 = new HF(a2);
				CS_0024_003C_003E8__locals5.A = this;
				Exception a3 = ex5;
				CS_0024_003C_003E8__locals5.A = a3;
				base.Dispatcher.Invoke([SpecialName] () =>
				{
					CS_0024_003C_003E8__locals5.A.D(CS_0024_003C_003E8__locals5.A.Message);
				});
				clsReporting.LogException(CS_0024_003C_003E8__locals5.A);
				ProjectData.ClearProjectError();
			}
		}
	}

	private void bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbFindReplace.Value = e.ProgressPercentage;
	}

	private void bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
	{
		grpFindReplace.IsEnabled = true;
		grpScope.IsEnabled = true;
		btnReplace.IsEnabled = true;
		btnStop.IsEnabled = false;
		stkFindReplace.Visibility = Visibility.Hidden;
		lblReplacing.Text = "";
		new ComAwareEventInfo(typeof(EApplication_Event), AH.A(56688)).AddEventHandler(this.m_A, new EApplication_PresentationBeforeCloseEventHandler(A));
		NumberFormatInfo numberFormatInfo = new NumberFormatInfo();
		numberFormatInfo.NumberDecimalSeparator = clsPublish.SystemDecimalSeparator();
		numberFormatInfo.NumberDecimalDigits = 0;
		string text = this.m_D.ToString(AH.A(7941), numberFormatInfo);
		string text2 = this.m_C.ToString(AH.A(7941), numberFormatInfo);
		string text3 = this.m_A.Count.ToString(AH.A(7941), numberFormatInfo);
		numberFormatInfo = null;
		if (!this.m_A.Any())
		{
			if (!e.Cancelled)
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
				E(AH.A(97498) + text + AH.A(98176) + text2 + AH.A(98197));
			}
			else
			{
				B(AH.A(98232) + text + AH.A(98176) + text2 + AH.A(98197));
			}
		}
		else
		{
			string text4 = (e.Cancelled ? (AH.A(98232) + text + AH.A(98176) + text2 + AH.A(98301) + text3 + AH.A(98368)) : (AH.A(97498) + text + AH.A(98176) + text2 + AH.A(98301) + text3 + AH.A(98368)));
			using (List<CF>.Enumerator enumerator = this.m_A.GetEnumerator())
			{
				while (enumerator.MoveNext())
				{
					CF current = enumerator.Current;
					text4 = text4 + AH.A(98403) + current.A + AH.A(7894) + current.B;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_02e4;
					}
					continue;
					end_IL_02e4:
					break;
				}
			}
			C(text4);
		}
		this.m_A = null;
		this.m_A = null;
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Presentation A)
	{
		string text = txtFind.Text;
		string text2 = txtReplace.Text;
		string text3 = Common.PresentationAuthor(A);
		bool value = chkRegex.IsChecked.Value;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Slides.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Slide slide = (Slide)enumerator.Current;
				try
				{
					enumerator2 = slide.Shapes.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						Microsoft.Office.Interop.PowerPoint.Shape a = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
						this.A(a, value, text, text2, text3);
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				this.A(slide, value, text, text2, text3);
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
		IEnumerator<ExcelLinkItem> enumerator3 = default(IEnumerator<ExcelLinkItem>);
		try
		{
			enumerator3 = this.m_A.GetEnumerator();
			while (enumerator3.MoveNext())
			{
				ExcelLinkItem current = enumerator3.Current;
				((LinkItem)current).SourcePath = Manage2.FindReplaceString(((LinkItem)current).SourcePath, text, text2, value);
				((LinkItem)current).ModifiedBy = text3;
			}
		}
		finally
		{
			if (enumerator3 != null)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					enumerator3.Dispose();
					break;
				}
			}
		}
		F();
		checked
		{
			this.m_C++;
		}
	}

	private void FindReplaceAsynch(Microsoft.Office.Interop.PowerPoint.Presentation pres, ref DoWorkEventArgs e)
	{
		IF a = default(IF);
		IF CS_0024_003C_003E8__locals15 = new IF(a);
		CS_0024_003C_003E8__locals15.A = this;
		string e2 = Common.PresentationAuthor(pres);
		int d = this.m_D;
		CS_0024_003C_003E8__locals15.A = null;
		CS_0024_003C_003E8__locals15.B = null;
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			CS_0024_003C_003E8__locals15.A = CS_0024_003C_003E8__locals15.A.txtFind.Text;
			CS_0024_003C_003E8__locals15.B = CS_0024_003C_003E8__locals15.A.txtReplace.Text;
			CS_0024_003C_003E8__locals15.A = CS_0024_003C_003E8__locals15.A.chkRegex.IsChecked.Value;
		});
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = pres.Slides.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (true)
			{
				IL_014a:
				if (enumerator.MoveNext())
				{
					Slide slide = (Slide)enumerator.Current;
					if (this.m_A == null)
					{
						break;
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
						if (this.m_A.CancellationPending)
						{
							e.Cancel = true;
							break;
						}
						enumerator2 = slide.Shapes.GetEnumerator();
						try
						{
							while (true)
							{
								if (enumerator2.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape a2 = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current;
									if (this.m_A == null)
									{
										break;
									}
									if (this.m_A.CancellationPending)
									{
										while (true)
										{
											switch (3)
											{
											case 0:
												continue;
											}
											e.Cancel = true;
											break;
										}
										break;
									}
									A(a2, CS_0024_003C_003E8__locals15.A, CS_0024_003C_003E8__locals15.A, CS_0024_003C_003E8__locals15.B, e2);
									continue;
								}
								while (true)
								{
									switch (1)
									{
									case 0:
										break;
									default:
										goto end_IL_010d;
									}
									continue;
									end_IL_010d:
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
						A(slide, CS_0024_003C_003E8__locals15.A, CS_0024_003C_003E8__locals15.A, CS_0024_003C_003E8__locals15.B, e2);
						goto IL_014a;
					}
					break;
				}
				while (true)
				{
					switch (7)
					{
					case 0:
						break;
					default:
						goto end_IL_0157;
					}
					continue;
					end_IL_0157:
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
		checked
		{
			if (!e.Cancel)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						pres.Save();
						this.m_C++;
						return;
					}
				}
			}
			this.m_D = d;
		}
	}

	private void A(Microsoft.Office.Interop.PowerPoint.Shape A, bool B, string C, string D, string E)
	{
		//IL_0034: Unknown result type (might be due to invalid IL or missing references)
		//IL_0039: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ef: Unknown result type (might be due to invalid IL or missing references)
		checked
		{
			if (A.Type != MsoShapeType.msoGroup)
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
						if (Shapes.IsLinked(A))
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
							string source = Shapes.LinkDetails(A).Source;
							string text = Manage2.FindReplaceString(source, C, D, B);
							if (Operators.CompareString(source, text, TextCompare: false) != 0)
							{
								Common.UpdateSource(A.Tags, null, text, blnUpdateLastModified: true);
								Common.UpdateUser(A.Tags, E);
								this.m_D++;
							}
						}
						if (Text.ContainsLinks(A))
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									break;
								default:
								{
									foreach (TextLink item in Text.SelectedLinks(A))
									{
										Link val = Text.LinkDetails(item);
										string source2 = val.Source;
										string text = Manage2.FindReplaceString(source2, C, D, B);
										if (Operators.CompareString(source2, text, TextCompare: false) != 0)
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
											Text.UpdateSource(item, val.RangeId, text, blnUpdateLastModified: true);
											Text.UpdateUser(item, val.RangeId, E);
											this.m_D++;
										}
									}
									return;
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
					this.A(a, B, C, D, E);
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
				if (enumerator2 is IDisposable)
				{
					while (true)
					{
						switch (1)
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

	private void A(Slide A, bool B, string C, string D, string E)
	{
		//IL_0027: Unknown result type (might be due to invalid IL or missing references)
		//IL_002c: Unknown result type (might be due to invalid IL or missing references)
		checked
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = A.Hyperlinks.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Microsoft.Office.Interop.PowerPoint.Hyperlink hyp = (Microsoft.Office.Interop.PowerPoint.Hyperlink)enumerator.Current;
					if (!Hyperlinks.IsLinked(hyp))
					{
						continue;
					}
					string source = Hyperlinks.LinkDetails(hyp).Source;
					string text = Manage2.FindReplaceString(source, C, D, B);
					if (Operators.CompareString(source, text, TextCompare: false) == 0)
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
					Hyperlinks.UpdateSource(hyp, null, text, blnUpdateLastModified: true);
					Hyperlinks.UpdateUser(hyp, E);
					this.m_D++;
				}
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
		}
	}

	private void A(TextBlock A, string B)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			A.Text = B;
		});
	}

	private void E()
	{
		IEnumerator<ExcelLinkItem> enumerator = default(IEnumerator<ExcelLinkItem>);
		try
		{
			enumerator = this.m_A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				((LinkItem)enumerator.Current).ResetError();
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

	private void A(LinkItem A, string B)
	{
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			((LinkItem)A).MarkBroken(B);
		});
		checked
		{
			this.m_A++;
		}
	}

	private void ShowLinkErrorMessage(object sender, MouseButtonEventArgs e)
	{
		LinkItem linkItem = (LinkItem)((Border)sender).DataContext;
		if (((LinkItem)linkItem).ErrorTooltip.Length > 0)
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
			D(((LinkItem)linkItem).ErrorTooltip);
		}
		linkItem = null;
	}

	private void SortShapeColumn(object sender, RoutedEventArgs e)
	{
		GridViewColumnHeader gridViewColumnHeader = (GridViewColumnHeader)sender;
		string text = gridViewColumnHeader.Tag.ToString();
		if (this.m_A != null)
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
			AdornerLayer.GetAdornerLayer(this.m_A).Remove(this.m_A);
			ShapesCollection.SortDescriptions.Clear();
		}
		ListSortDirection listSortDirection = ListSortDirection.Descending;
		if (this.m_A == gridViewColumnHeader && this.m_A.Direction == listSortDirection)
		{
			listSortDirection = ListSortDirection.Ascending;
		}
		this.m_A = gridViewColumnHeader;
		this.m_A = new SortAdorner(this.m_A, listSortDirection);
		AdornerLayer.GetAdornerLayer(this.m_A).Add(this.m_A);
		ShapesCollection.SortDescriptions.Add(new SortDescription(text, listSortDirection));
		if (Operators.CompareString(text, AH.A(98412), TextCompare: false) != 0)
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
			if (Operators.CompareString(text, AH.A(98433), TextCompare: false) != 0)
			{
				if (Operators.CompareString(text, AH.A(98464), TextCompare: false) != 0)
				{
					goto IL_0182;
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
		}
		AdornerPadding = new Thickness(0.0, 0.0, 10.0, 0.0);
		GridViewColumn column = gridViewColumnHeader.Column;
		if (double.IsNaN(column.Width))
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
			column.Width = column.ActualWidth;
		}
		column.Width = double.NaN;
		column = null;
		goto IL_0182;
		IL_0182:
		gridViewColumnHeader = null;
	}

	private ExcelLinkItem A(Slide A, object B, string C)
	{
		//IL_000e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c6: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cb: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cd: Unknown result type (might be due to invalid IL or missing references)
		//IL_012a: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_009b: Unknown result type (might be due to invalid IL or missing references)
		//IL_00a1: Unknown result type (might be due to invalid IL or missing references)
		//IL_00aa: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ac: Unknown result type (might be due to invalid IL or missing references)
		ExcelLinkItem result;
		try
		{
			Link lnk;
			Microsoft.Office.Interop.PowerPoint.Shape shpLinked;
			if (B is Microsoft.Office.Interop.PowerPoint.Shape)
			{
				lnk = Shapes.LinkDetails((Microsoft.Office.Interop.PowerPoint.Shape)B);
				shpLinked = (Microsoft.Office.Interop.PowerPoint.Shape)B;
			}
			else if (B is TextLink)
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
				object[] array;
				bool[] array2;
				object obj = NewLateBinding.LateGet(null, typeof(Text), AH.A(93278), array = new object[1] { B }, null, null, array2 = new bool[1] { true });
				if (array2[0])
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
					B = RuntimeHelpers.GetObjectValue(array[0]);
				}
				lnk = ((obj != null) ? ((Link)obj) : default(Link));
				shpLinked = Text.TextRangeParentShape(((TextLink)B).TextRange);
			}
			else
			{
				lnk = Hyperlinks.LinkDetails((Microsoft.Office.Interop.PowerPoint.Hyperlink)B);
				Type typeFromHandle = typeof(Hyperlinks);
				string memberName = AH.A(98485);
				object[] obj2 = new object[2] { B, false };
				object[] array = obj2;
				bool[] obj3 = new bool[2] { true, false };
				bool[] array2 = obj3;
				object obj4 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj2, null, null, obj3);
				if (array2[0])
				{
					B = RuntimeHelpers.GetObjectValue(array[0]);
				}
				shpLinked = (Microsoft.Office.Interop.PowerPoint.Shape)obj4;
			}
			result = new ExcelLinkItem(lnk, C, A, RuntimeHelpers.GetObjectValue(B), shpLinked);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = null;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private void btnClose_Click(object sender, RoutedEventArgs e)
	{
		if (this.m_A != null)
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
			if (this.m_A.IsBusy)
			{
				this.m_A.CancelAsync();
				return;
			}
		}
		try
		{
			this.m_A.Windows[1].Activate();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		Close();
	}

	private void B(bool A)
	{
		btnVerifyExcel.IsEnabled = A;
		btnEditShape.IsEnabled = A;
		btnUnlinkShape.IsEnabled = A;
		btnUpdateShape.IsEnabled = A;
		btnExportLinks.IsEnabled = A;
	}

	private void F()
	{
		Manage2.ForceColumnWidthUpdate(gvShapes);
	}

	private object A(ExcelLinkItem A)
	{
		//IL_005a: Unknown result type (might be due to invalid IL or missing references)
		object obj = RuntimeHelpers.GetObjectValue(A.LinkedObject);
		try
		{
			if (obj is Microsoft.Office.Interop.PowerPoint.Shape)
			{
				Conversions.ToString(NewLateBinding.LateGet(obj, null, AH.A(63335), new object[0], null, null, null));
			}
			else if (obj is TextLink)
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
					_ = ((TextLink)obj).TextRange.Text;
					break;
				}
			}
			else
			{
				Conversions.ToString(NewLateBinding.LateGet(obj, null, AH.A(98514), new object[0], null, null, null));
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			obj = null;
			ProjectData.ClearProjectError();
		}
		return obj;
	}

	private ObservableCollection<ExcelLinkItem> B()
	{
		return new ObservableCollection<ExcelLinkItem>(this.m_B.Where([SpecialName] (ExcelLinkItem A) => ((LinkItem)A).IsChecked));
	}

	private void G()
	{
		C(AH.A(98529));
	}

	private bool A()
	{
		if (this.m_A.Final)
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
					C(Common.A);
					return true;
				}
			}
		}
		return false;
	}

	private int A()
	{
		ObservableCollection<ExcelLinkItem> b = this.m_B;
		Func<ExcelLinkItem, bool> predicate;
		if (_Closure_0024__.C == null)
		{
			predicate = (_Closure_0024__.C = [SpecialName] (ExcelLinkItem A) => ((LinkItem)A).IsChecked);
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
			predicate = _Closure_0024__.C;
		}
		return b.Where(predicate).Count();
	}

	private void B(string A)
	{
		Forms.InfoMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void C(string A)
	{
		Forms.WarningMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void D(string A)
	{
		Forms.ErrorMessage(System.Windows.Window.GetWindow(this), A);
	}

	private void E(string A)
	{
		Forms.SuccessMessage(System.Windows.Window.GetWindow(this), A);
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (this.m_C)
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
			this.m_C = true;
			Uri resourceLocator = new Uri(AH.A(98564), UriKind.Relative);
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
	[DebuggerNonUserCode]
	[EditorBrowsable(EditorBrowsableState.Never)]
	public void System_Windows_Markup_IComponentConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 1)
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
					TabControl1 = (System.Windows.Controls.TabControl)target;
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
					tabShapes = (TabItem)target;
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
					lvShapes = (System.Windows.Controls.ListView)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					gvShapes = (GridView)target;
					return;
				}
			}
		}
		if (connectionId == 5)
		{
			chkShapes = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 7)
		{
			((GridViewColumnHeader)target).Click += SortShapeColumn;
			return;
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
					((GridViewColumnHeader)target).Click += SortShapeColumn;
					return;
				}
			}
		}
		if (connectionId == 9)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					((GridViewColumnHeader)target).Click += SortShapeColumn;
					return;
				}
			}
		}
		if (connectionId == 11)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					((GridViewColumnHeader)target).Click += SortShapeColumn;
					return;
				}
			}
		}
		if (connectionId == 12)
		{
			((GridViewColumnHeader)target).Click += SortShapeColumn;
			return;
		}
		if (connectionId == 13)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					chkFilterShapesToggle = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			btnUpdateShape = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 15)
		{
			btnVerifyExcel = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnEditShape = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 17)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnUnlinkShape = (System.Windows.Controls.Button)target;
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
					btnViewShape = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 19)
		{
			btnExportLinks = (System.Windows.Controls.Button)target;
			return;
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
					pbShapes = (System.Windows.Controls.ProgressBar)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			tbShapeCount = (TextBlock)target;
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
					grdShapeFilters = (Grid)target;
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
					cbxFilterShapeSource = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					cbxFilterModifiedBy = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 25)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkFilterRanges = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkFilterCharts = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkFilterTypeGraphic = (System.Windows.Controls.CheckBox)target;
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
					chkFilterTypePicture = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 29)
		{
			chkFilterTypeTable = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 30)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkFilterTypeWorkbook = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 31)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkFilterTypeText = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 32)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					chkFilterTypeChart = (System.Windows.Controls.CheckBox)target;
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
					btnReset = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 34)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					grpFindReplace = (System.Windows.Controls.GroupBox)target;
					return;
				}
			}
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
					txtFind = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 36)
		{
			txtReplace = (System.Windows.Controls.TextBox)target;
			return;
		}
		if (connectionId == 37)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					grpScope = (System.Windows.Controls.GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 38)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					radThisPresentation = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 39)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					radAllPresentations = (System.Windows.Controls.RadioButton)target;
					return;
				}
			}
		}
		if (connectionId == 40)
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					break;
				default:
					txtFolder = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 41)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					btnBrowse = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 42)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					chkSubfolders = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 43)
		{
			btnReplace = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 44)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					btnStop = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 45)
		{
			chkRegex = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 46)
		{
			stkFindReplace = (StackPanel)target;
			return;
		}
		if (connectionId == 47)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					pbFindReplace = (System.Windows.Controls.ProgressBar)target;
					return;
				}
			}
		}
		if (connectionId == 48)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					lblReplacing = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 49)
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
		this.m_C = true;
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
		if (connectionId == 6)
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
			((System.Windows.Controls.CheckBox)target).Checked += ShapeItemCheckedChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += ShapeItemCheckedChanged;
		}
		if (connectionId != 10)
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
			((Border)target).MouseLeftButtonUp += ShowLinkErrorMessage;
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
	private bool A(ExcelLinkItem A)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		bool num = Manage2.FilterLinks(A.Link, cbxFilterShapeSource, cbxFilterModifiedBy, chkFilterRanges, chkFilterCharts, chkFilterTypeGraphic, chkFilterTypePicture, chkFilterTypeTable, chkFilterTypeWorkbook, chkFilterTypeChart, chkFilterTypeText);
		if (!num)
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
			((LinkItem)A).IsChecked = false;
		}
		return num;
	}
}
