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
using System.Reflection;
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
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Links;

[DesignerGenerated]
public sealed class wpfManageLinks : System.Windows.Window, INotifyPropertyChanged, IComponentConnector, IStyleConnector
{
	private struct IB
	{
		public string A;

		public string B;
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<LinkItem, bool> A;

		public static Func<LinkItem, bool> B;

		public static Func<LinkItem, bool> C;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(LinkItem A)
		{
			return ((LinkItem)A).IsChecked;
		}

		[SpecialName]
		internal bool B(LinkItem A)
		{
			return ((LinkItem)A).IsChecked;
		}

		[SpecialName]
		internal bool C(LinkItem A)
		{
			return ((LinkItem)A).IsChecked;
		}
	}

	[CompilerGenerated]
	internal sealed class JB
	{
		public bool A;

		public object A;

		public bool B;

		public bool C;

		public CopierAsPicture A;

		public wpfManageLinks A;

		public JB(JB A)
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
				this.A = A.A;
				B = A.B;
				C = A.C;
				this.A = A.A;
				return;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class KB
	{
		public LinkItem A;

		public JB A;

		public KB(KB A)
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
			//IL_0268: Unknown result type (might be due to invalid IL or missing references)
			//IL_026e: Expected O, but got Unknown
			//IL_02dc: Unknown result type (might be due to invalid IL or missing references)
			//IL_02e6: Expected O, but got Unknown
			//IL_0384: Unknown result type (might be due to invalid IL or missing references)
			//IL_037a: Unknown result type (might be due to invalid IL or missing references)
			//IL_0380: Unknown result type (might be due to invalid IL or missing references)
			if (this.A.A)
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
				Microsoft.Office.Interop.Word.Window activeWindow = this.A.A.m_A.ActiveWindow;
				if (this.A.A is InlineShape)
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
					Microsoft.Office.Interop.Word.Window window = activeWindow;
					object objectValue = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(this.A.A, null, XC.A(44175), new object[0], null, null, null));
					object Start = RuntimeHelpers.GetObjectValue(Missing.Value);
					window.ScrollIntoView(objectValue, ref Start);
				}
				else if (this.A.A is Microsoft.Office.Interop.Word.Table)
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
					Microsoft.Office.Interop.Word.Window window2 = activeWindow;
					object objectValue2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(this.A.A, null, XC.A(44175), new object[0], null, null, null));
					object Start = RuntimeHelpers.GetObjectValue(Missing.Value);
					window2.ScrollIntoView(objectValue2, ref Start);
				}
				else if (this.A.A is Microsoft.Office.Interop.Word.ContentControl)
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
					Microsoft.Office.Interop.Word.Window window3 = activeWindow;
					object objectValue3 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(this.A.A, null, XC.A(44175), new object[0], null, null, null));
					object Start = RuntimeHelpers.GetObjectValue(Missing.Value);
					window3.ScrollIntoView(objectValue3, ref Start);
					this.A.B = true;
				}
				else
				{
					Microsoft.Office.Interop.Word.Window window4 = activeWindow;
					object objectValue4 = RuntimeHelpers.GetObjectValue(this.A.A);
					object Start = RuntimeHelpers.GetObjectValue(Missing.Value);
					window4.ScrollIntoView(objectValue4, ref Start);
				}
				activeWindow = null;
			}
			LinkItem linkItem = this.A;
			Type typeFromHandle = typeof(Refresh);
			string memberName = XC.A(44186);
			object[] obj = new object[4]
			{
				this.A.A,
				null,
				null,
				null
			};
			ref RefreshInstance a = ref this.A.A.m_A;
			obj[1] = a;
			obj[2] = this.A.C;
			obj[3] = this.A.A;
			object[] array = obj;
			bool[] array2;
			object obj2 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, array2 = new bool[4] { true, true, true, true });
			if (array2[0])
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
				this.A.A = RuntimeHelpers.GetObjectValue(array[0]);
			}
			if (array2[1])
			{
				a = (RefreshInstance)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[1]), typeof(RefreshInstance));
			}
			if (array2[2])
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
				this.A.C = (bool)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[2]), typeof(bool));
			}
			if (array2[3])
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
				this.A.A = (CopierAsPicture)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[3]), typeof(CopierAsPicture));
			}
			linkItem.LinkedObject = RuntimeHelpers.GetObjectValue(obj2);
			LinkItem linkItem2 = this.A;
			LinkItem linkItem3;
			object obj3 = NewLateBinding.LateGet(null, typeof(Common), XC.A(11777), array = new object[1] { (linkItem3 = this.A).LinkedObject }, null, null, array2 = new bool[1] { true });
			if (array2[0])
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
				linkItem3.LinkedObject = RuntimeHelpers.GetObjectValue(RuntimeHelpers.GetObjectValue(array[0]));
			}
			_003F link;
			if (obj3 == null)
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
				link = default(Link);
			}
			else
			{
				link = (Link)obj3;
			}
			linkItem2.Link = (Link)link;
		}
	}

	[CompilerGenerated]
	internal sealed class LB
	{
		public UpdateLinkException A;

		public KB A;

		public LB(LB A)
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
			((LinkItem)this.A.A).MarkBroken(((Exception)(object)this.A).Message);
		}
	}

	[CompilerGenerated]
	internal sealed class MB
	{
		public Exception A;

		public KB A;

		public MB(MB A)
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
			((LinkItem)this.A.A).MarkBroken(this.A.Message);
		}
	}

	[CompilerGenerated]
	internal sealed class NB
	{
		public Exception A;

		public wpfManageLinks A;

		public NB(NB A)
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
			this.A.D(this.A.Message);
		}
	}

	[CompilerGenerated]
	internal sealed class OB
	{
		public string A;

		public string B;

		public bool A;

		public wpfManageLinks A;

		public OB(OB A)
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
	internal sealed class PB
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
	internal sealed class QB
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

	private Microsoft.Office.Interop.Word.Application m_A;

	private Document m_A;

	private bool m_A;

	private ObservableCollection<LinkItem> m_A;

	private ObservableCollection<LinkItem> m_B;

	private BackgroundWorker m_A;

	private bool m_B;

	private RefreshInstance m_A;

	private IEnumerable<Microsoft.Office.Interop.Excel.Application> m_A;

	private List<string> m_A;

	private List<object> m_A;

	private int m_A;

	private int m_B;

	private bool m_C;

	private FileInfo[] m_A;

	private int m_C;

	private int m_D;

	private List<IB> m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("TabControl1")]
	private System.Windows.Controls.TabControl m_A;

	[AccessedThroughProperty("tabShapes")]
	[CompilerGenerated]
	private TabItem m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lvShapes")]
	private System.Windows.Controls.ListView m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("gvShapes")]
	private GridView m_A;

	[AccessedThroughProperty("chkShapes")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_A;

	[AccessedThroughProperty("chkFilterShapesToggle")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_B;

	[AccessedThroughProperty("btnUpdateShape")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_A;

	[AccessedThroughProperty("btnVerifyShape")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnEditShape")]
	private System.Windows.Controls.Button m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnUnlinkShape")]
	private System.Windows.Controls.Button m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnViewShape")]
	private System.Windows.Controls.Button m_E;

	[AccessedThroughProperty("btnExportLinks")]
	[CompilerGenerated]
	private System.Windows.Controls.Button m_F;

	[AccessedThroughProperty("pbShapes")]
	[CompilerGenerated]
	private System.Windows.Controls.ProgressBar m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("tbShapeCount")]
	private TextBlock m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("grdShapeFilters")]
	private Grid m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxFilterShapeSource")]
	private System.Windows.Controls.ComboBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cbxFilterModifiedBy")]
	private System.Windows.Controls.ComboBox m_B;

	[AccessedThroughProperty("chkFilterRanges")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterCharts")]
	private System.Windows.Controls.CheckBox m_D;

	[AccessedThroughProperty("chkFilterTypeGraphic")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterTypePicture")]
	private System.Windows.Controls.CheckBox m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterTypeTable")]
	private System.Windows.Controls.CheckBox m_G;

	[AccessedThroughProperty("chkFilterTypeWorkbook")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox m_H;

	[CompilerGenerated]
	[AccessedThroughProperty("chkFilterTypeText")]
	private System.Windows.Controls.CheckBox I;

	[AccessedThroughProperty("chkFilterTypeChart")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox J;

	[CompilerGenerated]
	[AccessedThroughProperty("btnReset")]
	private System.Windows.Controls.Button m_G;

	[AccessedThroughProperty("grpFindReplace")]
	[CompilerGenerated]
	private System.Windows.Controls.GroupBox m_A;

	[AccessedThroughProperty("txtFind")]
	[CompilerGenerated]
	private System.Windows.Controls.TextBox m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("txtReplace")]
	private System.Windows.Controls.TextBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("grpScope")]
	private System.Windows.Controls.GroupBox m_B;

	[AccessedThroughProperty("radThisDocument")]
	[CompilerGenerated]
	private System.Windows.Controls.RadioButton m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("radAllDocuments")]
	private System.Windows.Controls.RadioButton m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("txtFolder")]
	private System.Windows.Controls.TextBox m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("btnBrowse")]
	private System.Windows.Controls.Button m_H;

	[AccessedThroughProperty("chkSubfolders")]
	[CompilerGenerated]
	private System.Windows.Controls.CheckBox K;

	[CompilerGenerated]
	[AccessedThroughProperty("btnReplace")]
	private System.Windows.Controls.Button I;

	[AccessedThroughProperty("btnStop")]
	[CompilerGenerated]
	private System.Windows.Controls.Button J;

	[CompilerGenerated]
	[AccessedThroughProperty("chkRegex")]
	private System.Windows.Controls.CheckBox L;

	[AccessedThroughProperty("stkFindReplace")]
	[CompilerGenerated]
	private StackPanel m_A;

	[AccessedThroughProperty("pbFindReplace")]
	[CompilerGenerated]
	private System.Windows.Controls.ProgressBar m_B;

	[AccessedThroughProperty("lblReplacing")]
	[CompilerGenerated]
	private TextBlock m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnClose")]
	private System.Windows.Controls.Button K;

	private bool m_D;

	public ICollectionView ShapesCollection
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(XC.A(15360));
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
			A(XC.A(3447));
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
				switch (2)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
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
			this.m_B = value;
			checkBox = this.m_B;
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

	internal virtual System.Windows.Controls.Button btnVerifyShape
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
			RoutedEventHandler value2 = btnVerifyShape_Click;
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
			this.m_D = value;
			button = this.m_D;
			if (button != null)
			{
				button.Click += value2;
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
			this.m_F = value;
			button = this.m_F;
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
			return this.m_H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			this.m_H = value;
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
			this.m_G = value;
			button = this.m_G;
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

	internal virtual System.Windows.Controls.RadioButton radThisDocument
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

	internal virtual System.Windows.Controls.RadioButton radAllDocuments
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
			return this.m_H;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			RoutedEventHandler value2 = btnBrowse_Click;
			System.Windows.Controls.Button button = this.m_H;
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
			this.m_H = value;
			button = this.m_H;
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
			I = value;
			button = I;
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
			J = value;
			button = J;
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
	}

	public wpfManageLinks()
	{
		base.Loaded += wpfManageLinks_Loaded;
		base.Closing += wpfManageLinks_Closing;
		this.m_A = null;
		this.m_A = null;
		this.m_A = true;
		this.m_B = false;
		InitializeComponent();
		base.Icon = Forms.GetIcon();
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
			switch (1)
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

	private void wpfManageLinks_Loaded(object sender, RoutedEventArgs e)
	{
		List<string> B = new List<string>();
		List<string> C = new List<string>();
		this.m_A = PC.A.Application;
		this.m_A = this.m_A.ActiveDocument;
		Microsoft.Office.Interop.Word.View view = this.m_A.ActiveWindow.View;
		WdViewType type = view.Type;
		if ((uint)(type - 4) <= 1u)
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
			view.Type = WdViewType.wdPrintView;
		}
		view = null;
		this.m_A = new ObservableCollection<LinkItem>();
		B.Add(XC.A(15393));
		C.Add(XC.A(15393));
		_ = this.m_A.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
		foreach (Microsoft.Office.Interop.Word.Range storyRange in this.m_A.StoryRanges)
		{
			Microsoft.Office.Interop.Word.Range range = storyRange;
			do
			{
				A(range, ref B, ref C);
				range = range.NextStoryRange;
			}
			while (range != null);
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
		this.m_B = new ObservableCollection<LinkItem>(this.m_A);
		ShapesCollection = CollectionViewSource.GetDefaultView(this.m_B);
		Manage2.UpdateLinkCount(lvShapes, tbShapeCount);
		System.Windows.Controls.ComboBox comboBox = cbxFilterShapeSource;
		comboBox.ItemsSource = B.Distinct();
		comboBox.SelectedIndex = 0;
		_ = null;
		System.Windows.Controls.ComboBox comboBox2 = cbxFilterModifiedBy;
		comboBox2.ItemsSource = C.Distinct();
		comboBox2.SelectedIndex = 0;
		_ = null;
		B = null;
		C = null;
		this.C();
		A();
		lvShapes.Focus();
		chkShapes.IsEnabled = lvShapes.Items.Count > 0;
		A(this.m_A, lvShapes, chkShapes);
		radAllDocuments.IsEnabled = Base.IsUserAdmin();
		base.Activated += wpfManageLinks_Activated;
		base.Deactivated += wpfManageLinks_Deactivated;
		chkShapes.Checked += chkShapes_CheckedChanged;
		chkShapes.Unchecked += chkShapes_CheckedChanged;
		lvShapes.SelectionChanged += lvShapes_SelectionChanged;
		radThisDocument.Checked += FindReplaceScopeChanged;
		radAllDocuments.Checked += FindReplaceScopeChanged;
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(13443)).AddEventHandler(this.m_A, new ApplicationEvents4_DocumentBeforeCloseEventHandler(A));
	}

	private void A()
	{
		Selection selection = this.m_A.Selection;
		try
		{
			IEnumerator enumerator = selection.InlineShapes.GetEnumerator();
			try
			{
				IEnumerator<LinkItem> enumerator2 = default(IEnumerator<LinkItem>);
				while (enumerator.MoveNext())
				{
					InlineShape inlineShape = (InlineShape)enumerator.Current;
					try
					{
						enumerator2 = this.m_A.GetEnumerator();
						while (enumerator2.MoveNext())
						{
							LinkItem current = enumerator2.Current;
							if (!(B(current) is InlineShape))
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							if (Operators.CompareString(((InlineShape)B(current)).AlternativeText, inlineShape.AlternativeText, TextCompare: false) == 0)
							{
								((LinkItem)current).IsChecked = true;
								B(A: true);
								break;
							}
						}
					}
					finally
					{
						if (enumerator2 != null)
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								enumerator2.Dispose();
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
						goto end_IL_00d3;
					}
					continue;
					end_IL_00d3:
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		try
		{
			IEnumerator enumerator3 = default(IEnumerator);
			try
			{
				enumerator3 = Helpers.SelectedShapes(selection).GetEnumerator();
				while (enumerator3.MoveNext())
				{
					Microsoft.Office.Interop.Word.Shape a = (Microsoft.Office.Interop.Word.Shape)enumerator3.Current;
					A(a);
				}
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						goto end_IL_013b;
					}
					continue;
					end_IL_013b:
					break;
				}
			}
			finally
			{
				if (enumerator3 is IDisposable)
				{
					while (true)
					{
						switch (1)
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
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			ProjectData.ClearProjectError();
		}
		try
		{
			IEnumerator enumerator4 = default(IEnumerator);
			try
			{
				enumerator4 = selection.Tables.GetEnumerator();
				while (enumerator4.MoveNext())
				{
					Microsoft.Office.Interop.Word.Table table = (Microsoft.Office.Interop.Word.Table)enumerator4.Current;
					foreach (LinkItem item in this.m_A)
					{
						if (!(B(item) is Microsoft.Office.Interop.Word.Table))
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
						if (Operators.CompareString(((Microsoft.Office.Interop.Word.Table)B(item)).Descr, table.Descr, TextCompare: false) == 0)
						{
							((LinkItem)item).IsChecked = true;
							B(A: true);
							break;
						}
					}
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
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			ProjectData.ClearProjectError();
		}
		try
		{
			using List<Microsoft.Office.Interop.Word.ContentControl>.Enumerator enumerator6 = Common.LinkedContentControlsInSelection(selection).GetEnumerator();
			IEnumerator<LinkItem> enumerator7 = default(IEnumerator<LinkItem>);
			while (enumerator6.MoveNext())
			{
				Microsoft.Office.Interop.Word.ContentControl current3 = enumerator6.Current;
				try
				{
					enumerator7 = this.m_A.GetEnumerator();
					while (true)
					{
						if (enumerator7.MoveNext())
						{
							LinkItem current4 = enumerator7.Current;
							object objectValue = RuntimeHelpers.GetObjectValue(B(current4));
							if (objectValue is Microsoft.Office.Interop.Word.ContentControl && Operators.CompareString(((Microsoft.Office.Interop.Word.ContentControl)objectValue).ID, current3.ID, TextCompare: false) == 0)
							{
								((LinkItem)current4).IsChecked = true;
								B(A: true);
								break;
							}
							objectValue = null;
							continue;
						}
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								goto end_IL_02fd;
							}
							continue;
							end_IL_02fd:
							break;
						}
						break;
					}
				}
				finally
				{
					if (enumerator7 != null)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							enumerator7.Dispose();
							break;
						}
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
					goto end_IL_032b;
				}
				continue;
				end_IL_032b:
				break;
			}
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		selection = null;
	}

	private void A(Microsoft.Office.Interop.Word.Shape A)
	{
		if (A.Type == MsoShapeType.msoGroup)
		{
			{
				IEnumerator enumerator = A.GroupItems.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						Microsoft.Office.Interop.Word.Shape a = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
						this.A(a);
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
					IDisposable disposable = enumerator as IDisposable;
					if (disposable != null)
					{
						disposable.Dispose();
					}
				}
			}
		}
		using IEnumerator<LinkItem> enumerator2 = this.m_A.GetEnumerator();
		while (enumerator2.MoveNext())
		{
			LinkItem current = enumerator2.Current;
			if (!(B(current) is Microsoft.Office.Interop.Word.Shape))
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (Operators.CompareString(((Microsoft.Office.Interop.Word.Shape)B(current)).AlternativeText, A.AlternativeText, TextCompare: false) != 0)
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
				((LinkItem)current).IsChecked = true;
				B(A: true);
				return;
			}
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

	private void A(Microsoft.Office.Interop.Word.Range A, ref List<string> B, ref List<string> C)
	{
		//IL_0042: Unknown result type (might be due to invalid IL or missing references)
		IEnumerator enumerator = default(IEnumerator);
		LinkItem linkItem;
		try
		{
			enumerator = A.InlineShapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				InlineShape inlineShape = (InlineShape)enumerator.Current;
				if (!Common.IsLinked(inlineShape))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (Operators.CompareString(Common.LinkDetails(inlineShape).Source, string.Empty, TextCompare: false) == 0)
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
				linkItem = this.A(inlineShape);
				if (linkItem == null)
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
				this.m_A.Add(linkItem);
				B.Add(((LinkItem)linkItem).SourcePath);
				C.Add(((LinkItem)linkItem).ModifiedBy);
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					goto end_IL_00b3;
				}
				continue;
				end_IL_00b3:
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
		if (Helpers.A(A))
		{
			IEnumerator enumerator2 = default(IEnumerator);
			try
			{
				enumerator2 = A.ShapeRange.GetEnumerator();
				while (enumerator2.MoveNext())
				{
					Microsoft.Office.Interop.Word.Shape a = (Microsoft.Office.Interop.Word.Shape)enumerator2.Current;
					this.A(a, ref B);
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_0119;
					}
					continue;
					end_IL_0119:
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
		}
		IEnumerator enumerator3 = default(IEnumerator);
		try
		{
			enumerator3 = A.Tables.GetEnumerator();
			while (enumerator3.MoveNext())
			{
				Microsoft.Office.Interop.Word.Table table = (Microsoft.Office.Interop.Word.Table)enumerator3.Current;
				if (!Common.IsLinked(table))
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
				linkItem = this.A(table);
				if (linkItem == null)
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
				this.m_A.Add(linkItem);
				B.Add(((LinkItem)linkItem).SourcePath);
			}
		}
		finally
		{
			if (enumerator3 is IDisposable)
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					(enumerator3 as IDisposable).Dispose();
					break;
				}
			}
		}
		IEnumerator enumerator4 = default(IEnumerator);
		try
		{
			enumerator4 = A.ContentControls.GetEnumerator();
			while (enumerator4.MoveNext())
			{
				Microsoft.Office.Interop.Word.ContentControl contentControl = (Microsoft.Office.Interop.Word.ContentControl)enumerator4.Current;
				if (!Common.IsLinked(contentControl))
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
				linkItem = this.A(contentControl);
				if (linkItem == null)
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
				this.m_A.Add(linkItem);
				B.Add(((LinkItem)linkItem).SourcePath);
			}
		}
		finally
		{
			if (enumerator4 is IDisposable)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator4 as IDisposable).Dispose();
					break;
				}
			}
		}
		linkItem = null;
	}

	private void A(Microsoft.Office.Interop.Word.Shape A, ref List<string> B)
	{
		if (A.Type != MsoShapeType.msoGroup)
		{
			if (!Common.IsLinked(A))
			{
				return;
			}
			LinkItem linkItem = this.A((object)A);
			if (linkItem == null)
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
				this.m_A.Add(linkItem);
				B.Add(((LinkItem)linkItem).SourcePath);
				return;
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape a = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
				this.A(a, ref B);
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

	private void wpfManageLinks_Closing(object sender, CancelEventArgs e)
	{
		if (this.m_A != null)
		{
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
		radThisDocument.Checked -= FindReplaceScopeChanged;
		radAllDocuments.Checked -= FindReplaceScopeChanged;
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(13443)).RemoveEventHandler(this.m_A, new ApplicationEvents4_DocumentBeforeCloseEventHandler(A));
		this.m_A = null;
		this.m_A = null;
		this.m_A = null;
		this.m_B = null;
		Properties.ManageLinksHeight = base.Height;
		Properties.ManageLinksWidth = base.Width;
	}

	private void wpfManageLinks_Activated(object sender, EventArgs e)
	{
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).RemoveEventHandler(this.m_A, new ApplicationEvents4_WindowSelectionChangeEventHandler(A));
	}

	private void wpfManageLinks_Deactivated(object sender, EventArgs e)
	{
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).AddEventHandler(this.m_A, new ApplicationEvents4_WindowSelectionChangeEventHandler(A));
	}

	private void A(Document A, ref bool B)
	{
		Close();
	}

	private void A(Selection A)
	{
		lvShapes.SelectedItems.Clear();
		try
		{
			WdSelectionType type = A.Type;
			IEnumerator<LinkItem> enumerator2 = default(IEnumerator<LinkItem>);
			IEnumerator enumerator3 = default(IEnumerator);
			IEnumerator<LinkItem> enumerator4 = default(IEnumerator<LinkItem>);
			IEnumerator enumerator5 = default(IEnumerator);
			IEnumerator<LinkItem> enumerator6 = default(IEnumerator<LinkItem>);
			if (type != WdSelectionType.wdSelectionInlineShape)
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
						if (type != WdSelectionType.wdSelectionShape)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									if (Common.IsContentControlSelected(A))
									{
										while (true)
										{
											switch (7)
											{
											case 0:
												break;
											default:
											{
												using List<Microsoft.Office.Interop.Word.ContentControl>.Enumerator enumerator = Common.LinkedContentControlsInSelection(A).GetEnumerator();
												while (enumerator.MoveNext())
												{
													Microsoft.Office.Interop.Word.ContentControl current = enumerator.Current;
													try
													{
														enumerator2 = this.m_B.GetEnumerator();
														while (enumerator2.MoveNext())
														{
															LinkItem current2 = enumerator2.Current;
															try
															{
																if (Operators.CompareString(current.ID, ((Microsoft.Office.Interop.Word.ContentControl)B(current2)).ID, TextCompare: false) == 0)
																{
																	while (true)
																	{
																		switch (2)
																		{
																		case 0:
																			break;
																		default:
																			((LinkItem)current2).IsSelected = true;
																			goto end_IL_027c;
																		}
																		continue;
																		end_IL_027c:
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
														}
													}
													finally
													{
														if (enumerator2 != null)
														{
															while (true)
															{
																switch (1)
																{
																case 0:
																	break;
																default:
																	enumerator2.Dispose();
																	goto end_IL_02ae;
																}
																continue;
																end_IL_02ae:
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
														return;
													}
												}
											}
											}
										}
									}
									if (Common.IsTableSelected(A))
									{
										while (true)
										{
											switch (6)
											{
											case 0:
												break;
											default:
												try
												{
													enumerator3 = A.Tables.GetEnumerator();
													while (enumerator3.MoveNext())
													{
														Microsoft.Office.Interop.Word.Table table = (Microsoft.Office.Interop.Word.Table)enumerator3.Current;
														try
														{
															enumerator4 = this.m_B.GetEnumerator();
															while (enumerator4.MoveNext())
															{
																LinkItem current3 = enumerator4.Current;
																try
																{
																	if (Operators.CompareString(table.Range.ID, ((Microsoft.Office.Interop.Word.Table)B(current3)).Range.ID, TextCompare: false) == 0)
																	{
																		while (true)
																		{
																			switch (4)
																			{
																			case 0:
																				break;
																			default:
																				((LinkItem)current3).IsSelected = true;
																				goto end_IL_0377;
																			}
																			continue;
																			end_IL_0377:
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
															while (true)
															{
																switch (1)
																{
																case 0:
																	break;
																default:
																	goto end_IL_03a5;
																}
																continue;
																end_IL_03a5:
																break;
															}
														}
														finally
														{
															if (enumerator4 != null)
															{
																while (true)
																{
																	switch (4)
																	{
																	case 0:
																		break;
																	default:
																		enumerator4.Dispose();
																		goto end_IL_03b5;
																	}
																	continue;
																	end_IL_03b5:
																	break;
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
															return;
														}
													}
												}
												finally
												{
													if (enumerator3 is IDisposable)
													{
														while (true)
														{
															switch (1)
															{
															case 0:
																break;
															default:
																(enumerator3 as IDisposable).Dispose();
																goto end_IL_03ea;
															}
															continue;
															end_IL_03ea:
															break;
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
						try
						{
							enumerator5 = A.ShapeRange.GetEnumerator();
							while (enumerator5.MoveNext())
							{
								Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)enumerator5.Current;
								try
								{
									enumerator6 = this.m_B.GetEnumerator();
									while (enumerator6.MoveNext())
									{
										LinkItem current4 = enumerator6.Current;
										if (Operators.CompareString(shape.AlternativeText, ((Microsoft.Office.Interop.Word.Shape)B(current4)).AlternativeText, TextCompare: false) == 0)
										{
											((LinkItem)current4).IsSelected = true;
										}
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_01a7;
										}
										continue;
										end_IL_01a7:
										break;
									}
								}
								finally
								{
									if (enumerator6 != null)
									{
										while (true)
										{
											switch (5)
											{
											case 0:
												break;
											default:
												enumerator6.Dispose();
												goto end_IL_01b7;
											}
											continue;
											end_IL_01b7:
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
									return;
								}
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
										break;
									default:
										(enumerator5 as IDisposable).Dispose();
										goto end_IL_01ed;
									}
									continue;
									end_IL_01ed:
									break;
								}
							}
						}
					}
				}
			}
			IEnumerator enumerator7 = default(IEnumerator);
			try
			{
				enumerator7 = A.InlineShapes.GetEnumerator();
				IEnumerator<LinkItem> enumerator8 = default(IEnumerator<LinkItem>);
				while (enumerator7.MoveNext())
				{
					InlineShape inlineShape = (InlineShape)enumerator7.Current;
					try
					{
						enumerator8 = this.m_B.GetEnumerator();
						while (enumerator8.MoveNext())
						{
							LinkItem current5 = enumerator8.Current;
							try
							{
								if (Operators.CompareString(inlineShape.AlternativeText, ((InlineShape)B(current5)).AlternativeText, TextCompare: false) != 0)
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
									((LinkItem)current5).IsSelected = true;
									return;
								}
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
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_00dc;
							}
							continue;
							end_IL_00dc:
							break;
						}
					}
					finally
					{
						if (enumerator8 != null)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								enumerator8.Dispose();
								break;
							}
						}
					}
				}
			}
			finally
			{
				if (enumerator7 is IDisposable)
				{
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						(enumerator7 as IDisposable).Dispose();
						break;
					}
				}
			}
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
	}

	private void SpacebarToggleShapes(object sender, System.Windows.Input.KeyEventArgs e)
	{
		if (e.Key != Key.Space || lvShapes.SelectedItems.Count <= 0)
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
			this.m_B = true;
			try
			{
				bool isChecked = !((LinkItem)(LinkItem)lvShapes.SelectedItems[0]).IsChecked;
				enumerator = lvShapes.SelectedItems.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						((LinkItem)(LinkItem)enumerator.Current).IsChecked = isChecked;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							break;
						default:
							goto end_IL_0096;
						}
						continue;
						end_IL_0096:
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
				A(this.m_B, lvShapes, chkShapes);
				B(A() > 0);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			this.m_B = false;
			return;
		}
	}

	private void ShapeItemCheckedChanged(object sender, RoutedEventArgs e)
	{
		if (this.m_B)
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
			A(this.m_B, lvShapes, chkShapes);
			B(A() > 0);
			return;
		}
	}

	private void lvShapes_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (lvShapes.SelectedItems.Count == 1)
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
					btnViewShape.IsEnabled = true;
					LinkItem linkItem = this.m_B[lvShapes.SelectedIndex];
					Microsoft.Office.Interop.Word.Windows windows = this.m_A.Windows;
					object Index = 1;
					windows[ref Index].Activate();
					object objectValue = RuntimeHelpers.GetObjectValue(B(linkItem));
					if (objectValue != null)
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
						base.Activated -= wpfManageLinks_Activated;
						base.Deactivated -= wpfManageLinks_Deactivated;
						Common.NavigateLink(RuntimeHelpers.GetObjectValue(objectValue));
						Focus();
						base.Activated += wpfManageLinks_Activated;
						base.Deactivated += wpfManageLinks_Deactivated;
						objectValue = null;
					}
					else
					{
						this.m_A.Remove(linkItem);
						this.m_B.Remove(linkItem);
					}
					linkItem = null;
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
			switch (1)
			{
			case 0:
				continue;
			}
			btnViewShape.IsEnabled = false;
			return;
		}
	}

	private void A(ObservableCollection<LinkItem> A, System.Windows.Controls.ListView B, System.Windows.Controls.CheckBox C)
	{
		chkShapes.Checked -= chkShapes_CheckedChanged;
		chkShapes.Unchecked -= chkShapes_CheckedChanged;
		Forms.SynchCheckBox(A.Where([SpecialName] (LinkItem linkItem) => ((LinkItem)linkItem).IsChecked).Count(), B, C);
		chkShapes.Checked += chkShapes_CheckedChanged;
		chkShapes.Unchecked += chkShapes_CheckedChanged;
	}

	private void chkShapes_CheckedChanged(object sender, RoutedEventArgs e)
	{
		bool value = chkShapes.IsChecked.Value;
		this.m_B = true;
		A(this.m_B, value);
		lvShapes.Focus();
		this.m_B = false;
		B(value);
	}

	private void A(ObservableCollection<LinkItem> A, bool B)
	{
		foreach (LinkItem item in A)
		{
			((LinkItem)item).IsChecked = B;
		}
	}

	private void btnVerifyShape_Click(object sender, RoutedEventArgs e)
	{
		B();
	}

	private void B()
	{
		Manage2.GetExcelInstances(ref this.m_A, ref this.m_C);
		if (Manage2.IsExcelReady((Action<string>)C, this.m_A))
		{
			F();
			Manage2.StartProgressBar(pbShapes, tbShapeCount);
			this.m_A = new BackgroundWorker();
			BackgroundWorker a = this.m_A;
			a.WorkerSupportsCancellation = true;
			a.WorkerReportsProgress = true;
			a.DoWork += VerifyShapeLinksDoWork;
			a.ProgressChanged += VerifyShapeLinksProgressChanged;
			a.RunWorkerCompleted += VerifyShapeLinksComplete;
			a.RunWorkerAsync();
			_ = null;
		}
	}

	private void VerifyShapeLinksDoWork(object sender, DoWorkEventArgs e)
	{
		//IL_0090: Unknown result type (might be due to invalid IL or missing references)
		//IL_0095: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ca: Unknown result type (might be due to invalid IL or missing references)
		//IL_00cf: Unknown result type (might be due to invalid IL or missing references)
		//IL_00db: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e0: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00e9: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ee: Invalid comparison between Unknown and I4
		//IL_00fa: Unknown result type (might be due to invalid IL or missing references)
		//IL_00fe: Invalid comparison between Unknown and I4
		Workbook workbook = null;
		int num = 0;
		this.m_A = 0;
		this.m_A = new List<object>();
		ObservableCollection<LinkItem> observableCollection = B();
		this.m_B = observableCollection.Count;
		try
		{
			IEnumerator<LinkItem> enumerator = default(IEnumerator<LinkItem>);
			try
			{
				enumerator = observableCollection.GetEnumerator();
				while (true)
				{
					IL_0161:
					if (enumerator.MoveNext())
					{
						LinkItem current = enumerator.Current;
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
									switch (7)
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
								A(current, XC.A(15400));
								goto IL_0136;
							}
							if (!Manage2.FindExcelSource(current.Link, workbook))
							{
								ImportType type = current.Link.Type;
								if (type - 6 > 2)
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
									if ((int)type != 12)
									{
										A(current, XC.A(15782));
										goto IL_0134;
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
								A(current, XC.A(15612));
							}
							goto IL_0134;
							IL_0136:
							checked
							{
								num++;
								this.m_A.ReportProgress((int)Math.Round((double)num / (double)this.m_B * 100.0));
								goto IL_0161;
							}
							IL_0134:
							workbook = null;
							goto IL_0136;
						}
						break;
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_016e;
						}
						continue;
						end_IL_016e:
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			D(ex2.Message);
			ProjectData.ClearProjectError();
		}
		observableCollection = null;
	}

	private void VerifyShapeLinksProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbShapes.Value = e.ProgressPercentage;
	}

	private void VerifyShapeLinksComplete(object sender, RunWorkerCompletedEventArgs e)
	{
		Manage2.StopProgressBar(pbShapes, tbShapeCount);
		Manage2.ClosePreviouslyClosedWorkbooks(ref this.m_A, this.m_C);
		ReleaseHelper.ReleaseObjectEnumerable<Microsoft.Office.Interop.Excel.Application>(ref this.m_A, false);
		ReleaseHelper.DoGarbageCollection();
		if (!e.Cancelled)
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
		this.m_C = false;
	}

	private void btnEditShape_Click(object sender, RoutedEventArgs e)
	{
		//IL_01e5: Unknown result type (might be due to invalid IL or missing references)
		//IL_01db: Unknown result type (might be due to invalid IL or missing references)
		//IL_01e1: Unknown result type (might be due to invalid IL or missing references)
		int num = 0;
		if (!Manage2.IsAllExcelReady((Action<string>)C))
		{
			return;
		}
		checked
		{
			IEnumerator<LinkItem> enumerator = default(IEnumerator<LinkItem>);
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
				List<object> list = new List<object>();
				ObservableCollection<LinkItem> observableCollection = B();
				F();
				for (int i = observableCollection.Count - 1; i >= 0; i += -1)
				{
					LinkItem linkItem = observableCollection[i];
					object objectValue = RuntimeHelpers.GetObjectValue(B(linkItem));
					if (objectValue != null)
					{
						list.Add(RuntimeHelpers.GetObjectValue(objectValue));
					}
					else
					{
						this.m_A.Remove(linkItem);
						this.m_B.Remove(linkItem);
					}
					linkItem = null;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					object objectValue;
					Edit.EditedShapes editedShapes;
					if (list.Any())
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
						list.Reverse();
						base.Topmost = false;
						editedShapes = Edit.EditLink(list);
						base.Topmost = true;
						if (list.Count == editedShapes.Objects.Count)
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
							if (editedShapes.IsError != null)
							{
								try
								{
									enumerator = observableCollection.GetEnumerator();
									while (enumerator.MoveNext())
									{
										LinkItem current = enumerator.Current;
										if (editedShapes.IsError[num])
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
											((LinkItem)current).MarkBroken("");
										}
										else
										{
											objectValue = RuntimeHelpers.GetObjectValue(editedShapes.Objects[num]);
											Type typeFromHandle = typeof(Common);
											string memberName = XC.A(11777);
											object[] obj = new object[1] { objectValue };
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
												objectValue = RuntimeHelpers.GetObjectValue(array[0]);
											}
											_003F link;
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
												link = default(Link);
											}
											else
											{
												link = (Link)obj3;
											}
											current.Link = (Link)link;
											current.LinkedObject = RuntimeHelpers.GetObjectValue(objectValue);
										}
										num++;
									}
									while (true)
									{
										switch (4)
										{
										case 0:
											break;
										default:
											goto end_IL_0210;
										}
										continue;
										end_IL_0210:
										break;
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
							G();
						}
						else
						{
							D(XC.A(15930));
						}
					}
					objectValue = null;
					editedShapes = default(Edit.EditedShapes);
					list = null;
					observableCollection = null;
					return;
				}
			}
		}
	}

	private void btnViewShape_Click(object sender, RoutedEventArgs e)
	{
		if (lvShapes.SelectedItems.Count > 0)
		{
			object objectValue;
			LinkItem linkItem;
			try
			{
				linkItem = (LinkItem)lvShapes.SelectedItems[0];
				objectValue = RuntimeHelpers.GetObjectValue(B(linkItem));
				if (objectValue != null)
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
						object[] array;
						bool[] array2;
						NewLateBinding.LateCall(null, typeof(View), XC.A(6092), array = new object[1] { objectValue }, null, null, array2 = new bool[1] { true }, IgnoreReturn: true);
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
						LinkItem linkItem2 = linkItem;
						object instance = NewLateBinding.LateGet(null, typeof(Common), XC.A(11777), array = new object[1] { objectValue }, null, null, array2 = new bool[1] { true });
						if (array2[0])
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
							objectValue = RuntimeHelpers.GetObjectValue(array[0]);
						}
						((LinkItem)linkItem2).SourcePath = Conversions.ToString(NewLateBinding.LateGet(instance, null, XC.A(13872), new object[0], null, null, null));
						G();
						break;
					}
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				B();
				Interaction.AppActivate(this.m_A.Caption);
				ProjectData.ClearProjectError();
			}
			objectValue = null;
			linkItem = null;
		}
		else
		{
			C(XC.A(16033));
		}
	}

	private void btnUpdateShape_Click(object sender, RoutedEventArgs e)
	{
		//IL_0007: Unknown result type (might be due to invalid IL or missing references)
		//IL_0011: Expected O, but got Unknown
		//IL_0013: Unknown result type (might be due to invalid IL or missing references)
		//IL_0019: Expected O, but got Unknown
		//IL_001a: Expected O, but got Unknown
		try
		{
			this.m_A = new RefreshInstance(System.Windows.Window.GetWindow(this));
		}
		catch (UpdateLinkException ex)
		{
			ProjectData.SetProjectError((Exception)ex);
			UpdateLinkException ex2 = ex;
			C(((Exception)(object)ex2).Message);
			ProjectData.ClearProjectError();
			return;
		}
		F();
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
	}

	private void RefreshShapeLinksDoWork(object sender, DoWorkEventArgs e)
	{
		//IL_0176: Unknown result type (might be due to invalid IL or missing references)
		//IL_017c: Expected O, but got Unknown
		//IL_0190: Expected O, but got Unknown
		//IL_0085: Unknown result type (might be due to invalid IL or missing references)
		//IL_008f: Expected O, but got Unknown
		JB jB = new JB(jB);
		jB.A = this;
		int num = 0;
		UndoRecord undoRecord = this.m_A.UndoRecord;
		jB.B = false;
		this.m_A = 0;
		ObservableCollection<LinkItem> observableCollection = B();
		this.m_B = observableCollection.Count;
		jB.C = this.m_B == 1;
		jB.A = jB.C;
		undoRecord.StartCustomRecord(XC.A(16068));
		if (!jB.A)
		{
			this.m_A.ScreenUpdating = false;
		}
		jB.A = new CopierAsPicture();
		checked
		{
			IEnumerator<LinkItem> enumerator = default(IEnumerator<LinkItem>);
			try
			{
				enumerator = observableCollection.GetEnumerator();
				KB kB = default(KB);
				LB lB = default(LB);
				MB mB = default(MB);
				while (true)
				{
					IL_0278:
					if (enumerator.MoveNext())
					{
						kB = new KB(kB);
						kB.A = jB;
						kB.A = enumerator.Current;
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
									switch (4)
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
									switch (7)
									{
									case 0:
										break;
									default:
										goto end_IL_0111;
									}
									continue;
									end_IL_0111:
									break;
								}
								break;
							}
							kB.A.A = RuntimeHelpers.GetObjectValue(B(kB.A));
							if (kB.A.A != null)
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
								try
								{
									base.Dispatcher.Invoke(kB.A);
								}
								catch (UpdateLinkException ex)
								{
									ProjectData.SetProjectError((Exception)ex);
									lB = new LB(lB);
									lB.A = kB;
									UpdateLinkException a = ex;
									lB.A = a;
									this.m_A++;
									base.Dispatcher.Invoke(lB.A);
									lB.A.A.A = null;
									ProjectData.ClearProjectError();
								}
								catch (Exception ex2)
								{
									ProjectData.SetProjectError(ex2);
									mB = new MB(mB);
									mB.A = kB;
									Exception a2 = ex2;
									mB.A = a2;
									this.m_A++;
									base.Dispatcher.Invoke(mB.A);
									mB.A.A.A = null;
									ProjectData.ClearProjectError();
								}
								kB.A.A = null;
							}
							num++;
							this.m_A.ReportProgress((int)Math.Round((double)num / (double)this.m_B * 100.0));
							goto IL_0278;
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
							goto end_IL_0284;
						}
						continue;
						end_IL_0284:
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
			undoRecord.EndCustomRecord();
			undoRecord = null;
			this.m_A.ScreenUpdating = true;
			observableCollection = null;
		}
	}

	private void RefreshShapeLinksProgressChanged(object sender, ProgressChangedEventArgs e)
	{
		pbShapes.Value = e.ProgressPercentage;
	}

	private void RefreshShapeLinksComplete(object sender, RunWorkerCompletedEventArgs e)
	{
		Common.CloseHeaderFooterView(this.m_A);
		Focus();
		Manage2.StopProgressBar(pbShapes, tbShapeCount);
		G();
		Base.ReleaseRefreshInstance(ref this.m_A, true);
		lvShapes.SelectionChanged -= lvShapes_SelectionChanged;
		lvShapes.SelectedItems.Clear();
		lvShapes.SelectionChanged += lvShapes_SelectionChanged;
		btnViewShape.IsEnabled = false;
		if (!e.Cancelled)
		{
			if (this.m_A == 0)
			{
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
		this.m_C = false;
	}

	private void btnUnlinkShape_Click(object sender, RoutedEventArgs e)
	{
		ObservableCollection<LinkItem> observableCollection = B();
		checked
		{
			if (observableCollection.Any())
			{
				if (Base.ConfirmBreakLink())
				{
					lvShapes.SelectionChanged -= lvShapes_SelectionChanged;
					for (int i = observableCollection.Count - 1; i >= 0; i += -1)
					{
						LinkItem linkItem = observableCollection[i];
						object objectValue = RuntimeHelpers.GetObjectValue(B(linkItem));
						if (objectValue != null)
						{
							Type typeFromHandle = typeof(Break);
							string memberName = XC.A(16095);
							object[] obj = new object[1] { objectValue };
							object[] array = obj;
							bool[] obj2 = new bool[1] { true };
							bool[] array2 = obj2;
							NewLateBinding.LateCall(null, typeFromHandle, memberName, obj, null, null, obj2, IgnoreReturn: true);
							if (array2[0])
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
								objectValue = RuntimeHelpers.GetObjectValue(array[0]);
							}
							objectValue = null;
						}
						this.m_A.Remove(linkItem);
						this.m_B.Remove(linkItem);
						linkItem = null;
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
					lvShapes.SelectionChanged += lvShapes_SelectionChanged;
					chkShapes.Unchecked -= chkShapes_CheckedChanged;
					chkShapes.IsChecked = false;
					chkShapes.Unchecked += chkShapes_CheckedChanged;
					Manage2.UpdateLinkCount(lvShapes, tbShapeCount);
					B(A: false);
					G();
				}
			}
			else
			{
				H();
			}
			observableCollection = null;
		}
	}

	private void btnExportLinks_Click(object sender, RoutedEventArgs e)
	{
		//IL_019a: Unknown result type (might be due to invalid IL or missing references)
		//IL_019f: Unknown result type (might be due to invalid IL or missing references)
		//IL_01a1: Unknown result type (might be due to invalid IL or missing references)
		ObservableCollection<LinkItem> observableCollection = B();
		Microsoft.Office.Interop.Excel.Application application = null;
		if (observableCollection.Any())
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
				application.ScreenUpdating = false;
				application.EnableEvents = false;
				Worksheet worksheet;
				try
				{
					worksheet = Manage2.CreateNewWorksheet(application);
					Worksheet worksheet2 = worksheet;
					((Microsoft.Office.Interop.Excel.Range)worksheet2.Cells[1, 1]).Value2 = XC.A(16114);
					((Microsoft.Office.Interop.Excel.Range)worksheet2.Cells[1, 2]).Value2 = XC.A(16137);
					((Microsoft.Office.Interop.Excel.Range)worksheet2.Cells[1, 3]).Value2 = XC.A(16170);
					((Microsoft.Office.Interop.Excel.Range)worksheet2.Cells[1, 4]).Value2 = XC.A(16195);
					((Microsoft.Office.Interop.Excel.Range)worksheet2.Cells[1, 5]).Value2 = XC.A(16218);
					((Microsoft.Office.Interop.Excel.Range)worksheet2.Cells[1, 6]).Value2 = XC.A(16265);
					_ = null;
					int num = 2;
					using (IEnumerator<LinkItem> enumerator = observableCollection.GetEnumerator())
					{
						while (enumerator.MoveNext())
						{
							LinkItem current = enumerator.Current;
							Manage2.SetLinkType(worksheet, num, current.Link.Type);
							Manage2.SetLastRefresh((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[num, 3], ((LinkItem)current).LastUpdate);
							((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[num, 4]).Value2 = ((LinkItem)current).ModifiedBy;
							((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[num, 5]).Value2 = Manage2.GetSourceRangeString(((LinkItem)current).SourceRange);
							((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[num, 6]).Value2 = ((LinkItem)current).SourcePath;
							num = checked(num + 1);
						}
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								goto end_IL_0276;
							}
							continue;
							end_IL_0276:
							break;
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
				MC.A(application);
				application = null;
				worksheet = null;
			}
			else
			{
				C(XC.A(16288));
			}
		}
		else
		{
			H();
		}
		observableCollection = null;
	}

	private void ToggleShapeFilters(object sender, RoutedEventArgs e)
	{
		Manage2.ToggleFilters(chkFilterShapesToggle, grdShapeFilters);
	}

	private void C()
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

	private void D()
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
		E();
	}

	private void cbxFilterModifiedBy_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		E();
	}

	private void ShapesFilterChanged(object sender, RoutedEventArgs e)
	{
		E();
	}

	private void E()
	{
		this.m_B = true;
		lvShapes.SelectionChanged -= lvShapes_SelectionChanged;
		this.m_B = A();
		ShapesCollection = CollectionViewSource.GetDefaultView(this.m_B);
		lvShapes.SelectionChanged += lvShapes_SelectionChanged;
		this.m_B = false;
		Manage2.UpdateLinkCount(lvShapes, tbShapeCount);
		A(this.m_B, lvShapes, chkShapes);
		G();
	}

	private ObservableCollection<LinkItem> A()
	{
		List<LinkItem> list = this.m_A.Where([SpecialName] (LinkItem A) =>
		{
			//IL_0001: Unknown result type (might be due to invalid IL or missing references)
			//IL_0006: Unknown result type (might be due to invalid IL or missing references)
			bool num = Manage2.FilterLinks(A.Link, cbxFilterShapeSource, cbxFilterModifiedBy, chkFilterRanges, chkFilterCharts, chkFilterTypeGraphic, chkFilterTypePicture, chkFilterTypeTable, chkFilterTypeWorkbook, chkFilterTypeChart, chkFilterTypeText);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				((LinkItem)A).IsChecked = false;
			}
			return num;
		}).ToList();
		return new ObservableCollection<LinkItem>(list);
	}

	private bool A(Link A, System.Windows.Controls.ComboBox B)
	{
		//IL_001d: Unknown result type (might be due to invalid IL or missing references)
		if (B.SelectedIndex != 0)
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
					return Operators.CompareString(A.Source, B.SelectedValue.ToString(), TextCompare: false) == 0;
				}
			}
		}
		return true;
	}

	private void btnReset_Click(object sender, RoutedEventArgs e)
	{
		D();
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
		C();
		E();
	}

	private void FindReplaceScopeChanged(object sender, RoutedEventArgs e)
	{
		if (radThisDocument.IsChecked == true)
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
		Microsoft.Office.Core.FileDialog fileDialog = ((Microsoft.Office.Interop.Word._Application)this.m_A).get_FileDialog(MsoFileDialogType.msoFileDialogFolderPicker);
		fileDialog.Title = XC.A(16347);
		fileDialog.Filters.Clear();
		if (this.m_A.Documents.Count > 0)
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
			fileDialog.InitialFileName = this.m_A.ActiveDocument.Path;
		}
		fileDialog.AllowMultiSelect = false;
		base.Topmost = false;
		fileDialog.Show();
		base.Topmost = true;
		FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
		if (selectedItems.Count > 0)
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
			if (Forms.OkCancelMessage(XC.A(16374)) == System.Windows.Forms.DialogResult.Cancel)
			{
				while (true)
				{
					switch (4)
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
				switch (5)
				{
				case 0:
					continue;
				}
				break;
			}
			if (txtReplace.Text.Length != 0)
			{
				bool? isChecked = radAllDocuments.IsChecked;
				if (isChecked.HasValue)
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
					if (isChecked != true)
					{
						goto IL_012f;
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
				if (text.Length == 0 && isChecked.HasValue)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							break;
						default:
							C(XC.A(16709));
							return;
						}
					}
				}
				goto IL_012f;
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
		C(XC.A(16644));
		return;
		IL_012f:
		this.m_C = 0;
		this.m_D = 0;
		if (radThisDocument.IsChecked == true)
		{
			if (this.m_A.Documents.Count > 0)
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
				UndoRecord undoRecord = this.m_A.UndoRecord;
				undoRecord.StartCustomRecord(XC.A(16752));
				A(this.m_A.ActiveDocument);
				undoRecord.EndCustomRecord();
				E(XC.A(16781) + this.m_D + XC.A(16798));
			}
			else
			{
				C(XC.A(16851));
			}
		}
		else if (Manage2.IsFolderValid(text, (Action<string>)C))
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
			if (Forms.OkCancelMessage(XC.A(16890)) == System.Windows.Forms.DialogResult.OK)
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
				this.m_A = Manage2.FindReplaceFiles(txtFolder, chkSubfolders, XC.A(17226));
				if (this.m_A.Count() > 0)
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
					grpFindReplace.IsEnabled = false;
					grpScope.IsEnabled = false;
					btnReplace.IsEnabled = false;
					btnStop.IsEnabled = true;
					pbFindReplace.Maximum = 100.0;
					pbFindReplace.Value = 0.0;
					lblReplacing.Text = "";
					stkFindReplace.Visibility = System.Windows.Visibility.Visible;
					new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(13443)).RemoveEventHandler(this.m_A, new ApplicationEvents4_DocumentBeforeCloseEventHandler(A));
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
					B(XC.A(17239));
				}
			}
		}
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)10, XC.A(17278));
	}

	private void btnStop_Click(object sender, RoutedEventArgs e)
	{
		if (!this.m_A.IsBusy)
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
			this.m_A.CancelAsync();
			return;
		}
	}

	private void bgw_DoWork(object sender, DoWorkEventArgs e)
	{
		int num = 0;
		int num2 = this.m_A.Count();
		this.m_A = new List<IB>();
		this.m_A.DisplayAlerts = WdAlertLevel.wdAlertsNone;
		this.m_A.ScreenUpdating = false;
		checked
		{
			try
			{
				FileInfo[] a = this.m_A;
				int num3 = 0;
				IEnumerator enumerator = default(IEnumerator);
				while (true)
				{
					IL_0367:
					if (num3 < a.Length)
					{
						FileInfo fileInfo = a[num3];
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
								e.Cancel = true;
								break;
							}
							if (!fileInfo.Name.StartsWith(XC.A(17315)))
							{
								bool flag = false;
								try
								{
									enumerator = this.m_A.Documents.GetEnumerator();
									while (true)
									{
										if (enumerator.MoveNext())
										{
											Document document = (Document)enumerator.Current;
											if (Operators.CompareString(document.FullName, fileInfo.FullName, TextCompare: false) != 0)
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
												A(lblReplacing, document.FullName);
												FindReplaceAsynch(document, ref e);
												flag = true;
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
												goto end_IL_0116;
											}
											continue;
											end_IL_0116:
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
									Document document = null;
									try
									{
										Documents documents = this.m_A.Documents;
										object FileName = fileInfo.FullName;
										object ConfirmConversions = RuntimeHelpers.GetObjectValue(Missing.Value);
										object ReadOnly = false;
										object AddToRecentFiles = false;
										object PasswordDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
										object PasswordTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
										object Revert = RuntimeHelpers.GetObjectValue(Missing.Value);
										object WritePasswordDocument = RuntimeHelpers.GetObjectValue(Missing.Value);
										object WritePasswordTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
										object Format = RuntimeHelpers.GetObjectValue(Missing.Value);
										object Encoding = RuntimeHelpers.GetObjectValue(Missing.Value);
										object Visible = false;
										object OpenAndRepair = false;
										object DocumentDirection = RuntimeHelpers.GetObjectValue(Missing.Value);
										object NoEncodingDialog = RuntimeHelpers.GetObjectValue(Missing.Value);
										object XMLTransform = RuntimeHelpers.GetObjectValue(Missing.Value);
										document = documents.Open(ref FileName, ref ConfirmConversions, ref ReadOnly, ref AddToRecentFiles, ref PasswordDocument, ref PasswordTemplate, ref Revert, ref WritePasswordDocument, ref WritePasswordTemplate, ref Format, ref Encoding, ref Visible, ref OpenAndRepair, ref DocumentDirection, ref NoEncodingDialog, ref XMLTransform);
										if (!document.ReadOnly)
										{
											while (true)
											{
												switch (5)
												{
												case 0:
													continue;
												}
												A(lblReplacing, document.FullName);
												FindReplaceAsynch(document, ref e);
												break;
											}
										}
										else
										{
											IB item = new IB
											{
												A = fileInfo.FullName,
												B = XC.A(17318)
											};
											this.m_A.Add(item);
										}
									}
									catch (Exception ex)
									{
										ProjectData.SetProjectError(ex);
										Exception ex2 = ex;
										IB item = new IB
										{
											A = fileInfo.FullName,
											B = ex2.Message
										};
										this.m_A.Add(item);
										ProjectData.ClearProjectError();
									}
									finally
									{
										if (document != null)
										{
											while (true)
											{
												switch (3)
												{
												case 0:
													continue;
												}
												Document document2 = document;
												object XMLTransform = RuntimeHelpers.GetObjectValue(Missing.Value);
												object NoEncodingDialog = RuntimeHelpers.GetObjectValue(Missing.Value);
												object DocumentDirection = RuntimeHelpers.GetObjectValue(Missing.Value);
												document2.Close(ref XMLTransform, ref NoEncodingDialog, ref DocumentDirection);
												document = null;
												break;
											}
										}
									}
								}
								num++;
								this.m_A.ReportProgress((int)Math.Round((double)num / (double)num2 * 100.0));
							}
							num3++;
							goto IL_0367;
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
							goto end_IL_0372;
						}
						continue;
						end_IL_0372:
						break;
					}
					break;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				NB a2 = default(NB);
				NB CS_0024_003C_003E8__locals5 = new NB(a2);
				CS_0024_003C_003E8__locals5.A = this;
				Exception a3 = ex3;
				CS_0024_003C_003E8__locals5.A = a3;
				base.Dispatcher.Invoke([SpecialName] () =>
				{
					CS_0024_003C_003E8__locals5.A.D(CS_0024_003C_003E8__locals5.A.Message);
				});
				clsReporting.LogException(CS_0024_003C_003E8__locals5.A);
				ProjectData.ClearProjectError();
			}
			this.m_A.ScreenUpdating = true;
			this.m_A.DisplayAlerts = WdAlertLevel.wdAlertsAll;
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
		stkFindReplace.Visibility = System.Windows.Visibility.Hidden;
		lblReplacing.Text = "";
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(13443)).AddEventHandler(this.m_A, new ApplicationEvents4_DocumentBeforeCloseEventHandler(A));
		NumberFormatInfo numberFormatInfo = new NumberFormatInfo();
		numberFormatInfo.NumberDecimalSeparator = clsPublish.SystemDecimalSeparator();
		numberFormatInfo.NumberDecimalDigits = 0;
		string text = this.m_D.ToString(XC.A(17470), numberFormatInfo);
		string text2 = this.m_C.ToString(XC.A(17470), numberFormatInfo);
		string text3 = this.m_A.Count.ToString(XC.A(17470), numberFormatInfo);
		numberFormatInfo = null;
		if (!this.m_A.Any())
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
			if (!e.Cancelled)
			{
				E(XC.A(16781) + text + XC.A(17473) + text2 + XC.A(17494));
			}
			else
			{
				B(XC.A(17521) + text + XC.A(17473) + text2 + XC.A(17494));
			}
		}
		else
		{
			string text4;
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
				text4 = XC.A(16781) + text + XC.A(17473) + text2 + XC.A(17590) + text3 + XC.A(17649);
			}
			else
			{
				text4 = XC.A(17521) + text + XC.A(17473) + text2 + XC.A(17590) + text3 + XC.A(17649);
			}
			foreach (IB item in this.m_A)
			{
				text4 = text4 + XC.A(17676) + item.A + XC.A(17685) + item.B;
			}
			C(text4);
		}
		this.m_A = null;
		this.m_A = null;
	}

	private void A(Document A)
	{
		string text = txtFind.Text;
		string text2 = txtReplace.Text;
		string userName = A.Application.UserName;
		bool value = chkRegex.IsChecked.Value;
		_ = A.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.StoryRanges.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			IEnumerator enumerator3 = default(IEnumerator);
			IEnumerator enumerator4 = default(IEnumerator);
			IEnumerator enumerator5 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Range range = (Microsoft.Office.Interop.Word.Range)enumerator.Current;
				do
				{
					enumerator2 = range.InlineShapes.GetEnumerator();
					try
					{
						while (enumerator2.MoveNext())
						{
							InlineShape a = (InlineShape)enumerator2.Current;
							this.A(a, value, text, text2);
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
						IDisposable disposable = enumerator2 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
					if (Helpers.A(range))
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
							enumerator3 = range.ShapeRange.GetEnumerator();
							while (enumerator3.MoveNext())
							{
								Microsoft.Office.Interop.Word.Shape a2 = (Microsoft.Office.Interop.Word.Shape)enumerator3.Current;
								this.A(a2, value, text, text2);
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_0147;
								}
								continue;
								end_IL_0147:
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
					}
					try
					{
						enumerator4 = range.Tables.GetEnumerator();
						while (enumerator4.MoveNext())
						{
							Microsoft.Office.Interop.Word.Table a3 = (Microsoft.Office.Interop.Word.Table)enumerator4.Current;
							this.A(a3, value, text, text2);
						}
					}
					finally
					{
						if (enumerator4 is IDisposable)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								(enumerator4 as IDisposable).Dispose();
								break;
							}
						}
					}
					enumerator5 = range.ContentControls.GetEnumerator();
					try
					{
						while (enumerator5.MoveNext())
						{
							Microsoft.Office.Interop.Word.ContentControl a4 = (Microsoft.Office.Interop.Word.ContentControl)enumerator5.Current;
							this.A(a4, value, text, text2);
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								goto end_IL_0203;
							}
							continue;
							end_IL_0203:
							break;
						}
					}
					finally
					{
						IDisposable disposable2 = enumerator5 as IDisposable;
						if (disposable2 != null)
						{
							disposable2.Dispose();
						}
					}
					range = range.NextStoryRange;
				}
				while (range != null);
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
			while (true)
			{
				switch (5)
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
		IEnumerator<LinkItem> enumerator6 = default(IEnumerator<LinkItem>);
		try
		{
			enumerator6 = this.m_A.GetEnumerator();
			while (enumerator6.MoveNext())
			{
				LinkItem current = enumerator6.Current;
				((LinkItem)current).SourcePath = Manage2.FindReplaceString(((LinkItem)current).SourcePath, text, text2, value);
				((LinkItem)current).ModifiedBy = userName;
			}
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					goto end_IL_02bb;
				}
				continue;
				end_IL_02bb:
				break;
			}
		}
		finally
		{
			if (enumerator6 != null)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					enumerator6.Dispose();
					break;
				}
			}
		}
		G();
		checked
		{
			this.m_C++;
		}
	}

	private void FindReplaceAsynch(Document doc, ref DoWorkEventArgs e)
	{
		OB a = default(OB);
		OB CS_0024_003C_003E8__locals21 = new OB(a);
		CS_0024_003C_003E8__locals21.A = this;
		int d = this.m_D;
		CS_0024_003C_003E8__locals21.A = null;
		CS_0024_003C_003E8__locals21.B = null;
		base.Dispatcher.Invoke([SpecialName] () =>
		{
			CS_0024_003C_003E8__locals21.A = CS_0024_003C_003E8__locals21.A.txtFind.Text;
			CS_0024_003C_003E8__locals21.B = CS_0024_003C_003E8__locals21.A.txtReplace.Text;
			CS_0024_003C_003E8__locals21.A = CS_0024_003C_003E8__locals21.A.chkRegex.IsChecked.Value;
		});
		_ = doc.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;
		IEnumerator enumerator2 = default(IEnumerator);
		IEnumerator enumerator3 = default(IEnumerator);
		IEnumerator enumerator4 = default(IEnumerator);
		IEnumerator enumerator5 = default(IEnumerator);
		foreach (Microsoft.Office.Interop.Word.Range storyRange in doc.StoryRanges)
		{
			Microsoft.Office.Interop.Word.Range range = storyRange;
			do
			{
				try
				{
					enumerator2 = range.InlineShapes.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						InlineShape a2 = (InlineShape)enumerator2.Current;
						if (this.m_A == null)
						{
							break;
						}
						if (this.m_A.CancellationPending)
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
								e.Cancel = true;
								break;
							}
							break;
						}
						A(a2, CS_0024_003C_003E8__locals21.A, CS_0024_003C_003E8__locals21.A, CS_0024_003C_003E8__locals21.B);
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
								continue;
							}
							(enumerator2 as IDisposable).Dispose();
							break;
						}
					}
				}
				if (!e.Cancel)
				{
					if (Helpers.A(range))
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
							enumerator3 = range.ShapeRange.GetEnumerator();
							while (true)
							{
								IL_01aa:
								if (enumerator3.MoveNext())
								{
									Microsoft.Office.Interop.Word.Shape a3 = (Microsoft.Office.Interop.Word.Shape)enumerator3.Current;
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
										if (this.m_A.CancellationPending)
										{
											e.Cancel = true;
											break;
										}
										A(a3, CS_0024_003C_003E8__locals21.A, CS_0024_003C_003E8__locals21.A, CS_0024_003C_003E8__locals21.B);
										goto IL_01aa;
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
										goto end_IL_01b5;
									}
									continue;
									end_IL_01b5:
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
									switch (4)
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
							enumerator4 = range.Tables.GetEnumerator();
							while (true)
							{
								if (enumerator4.MoveNext())
								{
									Microsoft.Office.Interop.Word.Table a4 = (Microsoft.Office.Interop.Word.Table)enumerator4.Current;
									if (this.m_A == null)
									{
										break;
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
									A(a4, CS_0024_003C_003E8__locals21.A, CS_0024_003C_003E8__locals21.A, CS_0024_003C_003E8__locals21.B);
									continue;
								}
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_0266;
									}
									continue;
									end_IL_0266:
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
									switch (6)
									{
									case 0:
										continue;
									}
									(enumerator4 as IDisposable).Dispose();
									break;
								}
							}
						}
						if (!e.Cancel)
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
								enumerator5 = range.ContentControls.GetEnumerator();
								while (enumerator5.MoveNext())
								{
									Microsoft.Office.Interop.Word.ContentControl a5 = (Microsoft.Office.Interop.Word.ContentControl)enumerator5.Current;
									if (this.m_A == null)
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
										A(a5, CS_0024_003C_003E8__locals21.A, CS_0024_003C_003E8__locals21.A, CS_0024_003C_003E8__locals21.B);
										goto IL_031a;
									}
									break;
									IL_031a:;
								}
							}
							finally
							{
								if (enumerator5 is IDisposable)
								{
									while (true)
									{
										switch (1)
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
					}
				}
				range = range.NextStoryRange;
			}
			while (range != null);
		}
		checked
		{
			if (!e.Cancel)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						doc.Save();
						this.m_C++;
						return;
					}
				}
			}
			this.m_D = d;
		}
	}

	private void A(Microsoft.Office.Interop.Word.Shape A, bool B, string C, string D)
	{
		if (A.Type != MsoShapeType.msoGroup)
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
					this.A((object)A, B, C, D);
					return;
				}
			}
		}
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape a = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
				this.A(a, B, C, D);
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

	private void A(object A, bool B, string C, string D)
	{
		Type typeFromHandle = typeof(Common);
		string memberName = XC.A(13018);
		object[] obj = new object[1] { A };
		object[] array = obj;
		bool[] obj2 = new bool[1] { true };
		bool[] array2 = obj2;
		object value = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
		if (array2[0])
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
			A = RuntimeHelpers.GetObjectValue(array[0]);
		}
		if (!Conversions.ToBoolean(value))
		{
			return;
		}
		Type typeFromHandle2 = typeof(Common);
		string memberName2 = XC.A(11777);
		object[] obj3 = new object[1] { A };
		array = obj3;
		bool[] obj4 = new bool[1] { true };
		array2 = obj4;
		object instance = NewLateBinding.LateGet(null, typeFromHandle2, memberName2, obj3, null, null, obj4);
		if (array2[0])
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
			A = RuntimeHelpers.GetObjectValue(array[0]);
		}
		string text = Conversions.ToString(NewLateBinding.LateGet(instance, null, XC.A(13872), new object[0], null, null, null));
		if (Operators.CompareString(text, string.Empty, TextCompare: false) == 0)
		{
			return;
		}
		checked
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				string text2 = Manage2.FindReplaceString(text, C, D, B);
				if (Operators.CompareString(text, text2, TextCompare: false) == 0)
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
					Type typeFromHandle3 = typeof(Update);
					string memberName3 = XC.A(13872);
					object[] obj5 = new object[4] { A, null, text2, true };
					array = obj5;
					bool[] obj6 = new bool[4] { true, false, true, false };
					array2 = obj6;
					NewLateBinding.LateCall(null, typeFromHandle3, memberName3, obj5, null, null, obj6, IgnoreReturn: true);
					if (array2[0])
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
						A = RuntimeHelpers.GetObjectValue(array[0]);
					}
					if (array2[2])
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
						text2 = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[2]), typeof(string));
					}
					NewLateBinding.LateCall(null, typeof(Update), XC.A(17690), array = new object[1] { A }, null, null, array2 = new bool[1] { true }, IgnoreReturn: true);
					if (array2[0])
					{
						A = RuntimeHelpers.GetObjectValue(array[0]);
					}
					this.m_D++;
					return;
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

	private void F()
	{
		IEnumerator<LinkItem> enumerator = default(IEnumerator<LinkItem>);
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
		LinkItem linkItem = (LinkItem)((System.Windows.Controls.Border)sender).DataContext;
		if (((LinkItem)linkItem).ErrorTooltip.Length > 0)
		{
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
			AdornerLayer.GetAdornerLayer(this.m_A).Remove(this.m_A);
			ShapesCollection.SortDescriptions.Clear();
		}
		ListSortDirection listSortDirection = ListSortDirection.Descending;
		if (this.m_A == gridViewColumnHeader && this.m_A.Direction == listSortDirection)
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
			listSortDirection = ListSortDirection.Ascending;
		}
		this.m_A = gridViewColumnHeader;
		this.m_A = new SortAdorner(this.m_A, listSortDirection);
		AdornerLayer.GetAdornerLayer(this.m_A).Add(this.m_A);
		ShapesCollection.SortDescriptions.Add(new SortDescription(text, listSortDirection));
		if (Operators.CompareString(text, XC.A(17699), TextCompare: false) != 0)
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
			if (Operators.CompareString(text, XC.A(17720), TextCompare: false) != 0)
			{
				goto IL_0171;
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
		AdornerPadding = new Thickness(0.0, 0.0, 10.0, 0.0);
		GridViewColumn column = gridViewColumnHeader.Column;
		if (double.IsNaN(column.Width))
		{
			column.Width = column.ActualWidth;
		}
		column.Width = double.NaN;
		column = null;
		goto IL_0171;
		IL_0171:
		gridViewColumnHeader = null;
	}

	private LinkItem A(object A)
	{
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_0070: Unknown result type (might be due to invalid IL or missing references)
		//IL_0076: Unknown result type (might be due to invalid IL or missing references)
		LinkItem result;
		try
		{
			Type typeFromHandle = typeof(Common);
			string memberName = XC.A(11777);
			object[] obj = new object[1] { A };
			object[] array = obj;
			bool[] obj2 = new bool[1] { true };
			bool[] array2 = obj2;
			object obj3 = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
			if (array2[0])
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
				A = RuntimeHelpers.GetObjectValue(array[0]);
			}
			_003F link;
			if (obj3 == null)
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
				link = default(Link);
			}
			else
			{
				link = (Link)obj3;
			}
			result = new LinkItem((Link)link, RuntimeHelpers.GetObjectValue(A));
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
			if (this.m_A.IsBusy)
			{
				this.m_A.CancelAsync();
				return;
			}
		}
		try
		{
			Microsoft.Office.Interop.Word.Windows windows = this.m_A.Windows;
			object Index = 1;
			windows[ref Index].Activate();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		new ComAwareEventInfo(typeof(ApplicationEvents4_Event), XC.A(1839)).RemoveEventHandler(this.m_A, new ApplicationEvents4_WindowSelectionChangeEventHandler(A));
		Close();
	}

	private void B(bool A)
	{
		btnVerifyShape.IsEnabled = A;
		btnEditShape.IsEnabled = A;
		btnUnlinkShape.IsEnabled = A;
		btnUpdateShape.IsEnabled = A;
		btnExportLinks.IsEnabled = A;
	}

	private void G()
	{
		Manage2.ForceColumnWidthUpdate(gvShapes);
	}

	private object B(LinkItem A)
	{
		object obj = RuntimeHelpers.GetObjectValue(A.LinkedObject);
		try
		{
			Conversions.ToString(NewLateBinding.LateGet(obj, null, XC.A(1509), new object[0], null, null, null));
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

	private ObservableCollection<LinkItem> B()
	{
		ObservableCollection<LinkItem> b = this.m_B;
		Func<LinkItem, bool> predicate;
		if (_Closure_0024__.B == null)
		{
			predicate = (_Closure_0024__.B = [SpecialName] (LinkItem A) => ((LinkItem)A).IsChecked);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			predicate = _Closure_0024__.B;
		}
		return new ObservableCollection<LinkItem>(b.Where(predicate));
	}

	private void H()
	{
		C(XC.A(17751));
	}

	private int A()
	{
		return this.m_B.Where([SpecialName] (LinkItem A) => ((LinkItem)A).IsChecked).Count();
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

	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[DebuggerNonUserCode]
	public void InitializeComponent()
	{
		if (!this.m_D)
		{
			this.m_D = true;
			Uri resourceLocator = new Uri(XC.A(17786), UriKind.Relative);
			System.Windows.Application.LoadComponent(this, resourceLocator);
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
					TabControl1 = (System.Windows.Controls.TabControl)target;
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
					tabShapes = (TabItem)target;
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
					lvShapes = (System.Windows.Controls.ListView)target;
					return;
				}
			}
		}
		if (connectionId == 4)
		{
			gvShapes = (GridView)target;
			return;
		}
		if (connectionId == 5)
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
		if (connectionId == 7)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					((GridViewColumnHeader)target).Click += SortShapeColumn;
					return;
				}
			}
		}
		if (connectionId == 8)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					((GridViewColumnHeader)target).Click += SortShapeColumn;
					return;
				}
			}
		}
		if (connectionId == 10)
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
			((GridViewColumnHeader)target).Click += SortShapeColumn;
			return;
		}
		if (connectionId == 12)
		{
			chkFilterShapesToggle = (System.Windows.Controls.CheckBox)target;
			return;
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
					btnUpdateShape = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 14)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					btnVerifyShape = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 15)
		{
			btnEditShape = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 16)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					btnUnlinkShape = (System.Windows.Controls.Button)target;
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
					btnViewShape = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 18)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					btnExportLinks = (System.Windows.Controls.Button)target;
					return;
				}
			}
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
					pbShapes = (System.Windows.Controls.ProgressBar)target;
					return;
				}
			}
		}
		if (connectionId == 20)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					tbShapeCount = (TextBlock)target;
					return;
				}
			}
		}
		if (connectionId == 21)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					grdShapeFilters = (Grid)target;
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
					cbxFilterShapeSource = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 23)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					cbxFilterModifiedBy = (System.Windows.Controls.ComboBox)target;
					return;
				}
			}
		}
		if (connectionId == 24)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkFilterRanges = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
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
					chkFilterCharts = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 26)
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					chkFilterTypeGraphic = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 27)
		{
			chkFilterTypePicture = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 28)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					chkFilterTypeTable = (System.Windows.Controls.CheckBox)target;
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
					chkFilterTypeWorkbook = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 30)
		{
			chkFilterTypeText = (System.Windows.Controls.CheckBox)target;
			return;
		}
		if (connectionId == 31)
		{
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					chkFilterTypeChart = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
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
					btnReset = (System.Windows.Controls.Button)target;
					return;
				}
			}
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
					grpFindReplace = (System.Windows.Controls.GroupBox)target;
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
					txtFind = (System.Windows.Controls.TextBox)target;
					return;
				}
			}
		}
		if (connectionId == 35)
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					txtReplace = (System.Windows.Controls.TextBox)target;
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
					grpScope = (System.Windows.Controls.GroupBox)target;
					return;
				}
			}
		}
		if (connectionId == 37)
		{
			radThisDocument = (System.Windows.Controls.RadioButton)target;
			return;
		}
		if (connectionId == 38)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					radAllDocuments = (System.Windows.Controls.RadioButton)target;
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
					txtFolder = (System.Windows.Controls.TextBox)target;
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
					btnBrowse = (System.Windows.Controls.Button)target;
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
					chkSubfolders = (System.Windows.Controls.CheckBox)target;
					return;
				}
			}
		}
		if (connectionId == 42)
		{
			btnReplace = (System.Windows.Controls.Button)target;
			return;
		}
		if (connectionId == 43)
		{
			while (true)
			{
				switch (5)
				{
				case 0:
					break;
				default:
					btnStop = (System.Windows.Controls.Button)target;
					return;
				}
			}
		}
		if (connectionId == 44)
		{
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					chkRegex = (System.Windows.Controls.CheckBox)target;
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
					stkFindReplace = (StackPanel)target;
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
					pbFindReplace = (System.Windows.Controls.ProgressBar)target;
					return;
				}
			}
		}
		if (connectionId == 47)
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
		if (connectionId == 48)
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
		this.m_D = true;
	}

	void IComponentConnector.Connect(int connectionId, object target)
	{
		//ILSpy generated this explicit interface implementation from .override directive in System_Windows_Markup_IComponentConnector_Connect
		this.System_Windows_Markup_IComponentConnector_Connect(connectionId, target);
	}

	[EditorBrowsable(EditorBrowsableState.Never)]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void System_Windows_Markup_IStyleConnector_Connect(int connectionId, object target)
	{
		if (connectionId == 6)
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
			((System.Windows.Controls.CheckBox)target).Checked += ShapeItemCheckedChanged;
			((System.Windows.Controls.CheckBox)target).Unchecked += ShapeItemCheckedChanged;
		}
		if (connectionId != 9)
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
			((System.Windows.Controls.Border)target).MouseLeftButtonUp += ShowLinkErrorMessage;
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
	private bool A(LinkItem A)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0006: Unknown result type (might be due to invalid IL or missing references)
		bool num = Manage2.FilterLinks(A.Link, cbxFilterShapeSource, cbxFilterModifiedBy, chkFilterRanges, chkFilterCharts, chkFilterTypeGraphic, chkFilterTypePicture, chkFilterTypeTable, chkFilterTypeWorkbook, chkFilterTypeChart, chkFilterTypeText);
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			((LinkItem)A).IsChecked = false;
		}
		return num;
	}
}
