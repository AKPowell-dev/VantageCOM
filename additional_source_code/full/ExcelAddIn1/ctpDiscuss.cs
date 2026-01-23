using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Forms;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using MacabacusMacros.Links;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

[DesignerGenerated]
public sealed class ctpDiscuss : UserControl
{
	public enum MessageType
	{
		Text = 1,
		File,
		Link,
		ScreenShot
	}

	private struct DH
	{
		public Range A;

		public CustomXMLPart A;
	}

	private struct EH
	{
		public string A;

		public DateTime A;
	}

	private struct FH
	{
		public Image A;

		public string A;
	}

	[CompilerGenerated]
	internal sealed class GH
	{
		public List<FileLinkButton> A;

		public ctpDiscuss A;

		public GH(GH A)
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
			foreach (FileLinkButton item in this.A)
			{
				bool flag = false;
				string toolTip = this.A.ToolTip1.GetToolTip(item);
				string text;
				Image image;
				if (!this.A.m_A.ContainsKey(toolTip))
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
					Uri uri = new Uri(toolTip);
					image = this.A.Icons.Images[10];
					try
					{
						if (uri.HostNameType == UriHostNameType.Dns)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								HttpWebResponse httpWebResponse = (HttpWebResponse)WebRequest.Create(VH.A(212142) + uri.Host + VH.A(212157)).GetResponse();
								Stream responseStream = httpWebResponse.GetResponseStream();
								image = new Bitmap(Image.FromStream(responseStream), 16, 16);
								httpWebResponse.Close();
								responseStream.Close();
								httpWebResponse = null;
								break;
							}
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						flag = true;
						ProjectData.ClearProjectError();
					}
					text = toolTip;
					try
					{
						text = Regex.Match(new WebClient().DownloadString(uri), VH.A(212182)).Groups[1].ToString();
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						flag = true;
						ProjectData.ClearProjectError();
					}
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
						FH value = new FH
						{
							A = image,
							A = text
						};
						this.A.m_A.Add(toolTip, value);
					}
				}
				else
				{
					FH fH = this.A.m_A[toolTip];
					image = fH.A;
					text = fH.A;
				}
				item.Icon = image;
				item.Text = text;
				image = null;
			}
		}
	}

	private IContainer m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("lvDiscussions")]
	private ListView m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("colSheet")]
	private ColumnHeader m_A;

	[AccessedThroughProperty("colCell")]
	[CompilerGenerated]
	private ColumnHeader m_B;

	[AccessedThroughProperty("colValue")]
	[CompilerGenerated]
	private ColumnHeader m_C;

	[AccessedThroughProperty("colUser")]
	[CompilerGenerated]
	private ColumnHeader m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("SplitContainer1")]
	private SplitContainer m_A;

	[AccessedThroughProperty("colDate")]
	[CompilerGenerated]
	private ColumnHeader m_E;

	[AccessedThroughProperty("btnNew")]
	[CompilerGenerated]
	private Button m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnDelete")]
	private Button m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("rtbComment")]
	private RichTextBox m_A;

	[AccessedThroughProperty("pnlRichTextBox")]
	[CompilerGenerated]
	private Panel m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("Panel2")]
	private Panel m_B;

	[AccessedThroughProperty("ToolTip1")]
	[CompilerGenerated]
	private ToolTip m_A;

	[AccessedThroughProperty("Icons")]
	[CompilerGenerated]
	private ImageList m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("cmsFile")]
	private ContextMenuStrip m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFileOpen")]
	private ToolStripMenuItem m_A;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFileDelete")]
	private ToolStripMenuItem m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFileShow")]
	private ToolStripMenuItem m_C;

	[CompilerGenerated]
	[AccessedThroughProperty("cmsLink")]
	private ContextMenuStrip m_B;

	[AccessedThroughProperty("btnLinkFollow")]
	[CompilerGenerated]
	private ToolStripMenuItem m_D;

	[CompilerGenerated]
	[AccessedThroughProperty("btnLinkDelete")]
	private ToolStripMenuItem m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("btnFiles")]
	private Button m_C;

	[AccessedThroughProperty("btnMessage")]
	[CompilerGenerated]
	private Button m_D;

	[AccessedThroughProperty("cmsComment")]
	[CompilerGenerated]
	private ContextMenuStrip m_C;

	[AccessedThroughProperty("btnCommentDelete")]
	[CompilerGenerated]
	private ToolStripMenuItem m_F;

	[AccessedThroughProperty("btnDeleteAll")]
	[CompilerGenerated]
	private Button m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("cmsPicture")]
	private ContextMenuStrip m_D;

	[AccessedThroughProperty("btnImageView")]
	[CompilerGenerated]
	private ToolStripMenuItem m_G;

	[CompilerGenerated]
	[AccessedThroughProperty("btnImageDelete")]
	private ToolStripMenuItem m_H;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPen")]
	private System.Windows.Forms.CheckBox m_A;

	[AccessedThroughProperty("chkHighlighter")]
	[CompilerGenerated]
	private System.Windows.Forms.CheckBox m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPenMenu")]
	private System.Windows.Forms.CheckBox m_C;

	[AccessedThroughProperty("chkHighlighterMenu")]
	[CompilerGenerated]
	private System.Windows.Forms.CheckBox m_D;

	[AccessedThroughProperty("cmsPens")]
	[CompilerGenerated]
	private ContextMenuStrip m_E;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPenRed")]
	private ToolStripMenuItem m_I;

	[AccessedThroughProperty("chkPenGreen")]
	[CompilerGenerated]
	private ToolStripMenuItem m_J;

	[AccessedThroughProperty("chkPenBlue")]
	[CompilerGenerated]
	private ToolStripMenuItem m_K;

	[AccessedThroughProperty("chkPenYellow")]
	[CompilerGenerated]
	private ToolStripMenuItem m_L;

	[CompilerGenerated]
	[AccessedThroughProperty("ToolStripSeparator1")]
	private ToolStripSeparator m_A;

	[AccessedThroughProperty("chkPenThin")]
	[CompilerGenerated]
	private ToolStripMenuItem m_M;

	[CompilerGenerated]
	[AccessedThroughProperty("chkPenMedium")]
	private ToolStripMenuItem m_N;

	[AccessedThroughProperty("chkPenThick")]
	[CompilerGenerated]
	private ToolStripMenuItem m_O;

	[AccessedThroughProperty("cmsHighlighters")]
	[CompilerGenerated]
	private ContextMenuStrip m_F;

	[AccessedThroughProperty("chkHighlighterYellow")]
	[CompilerGenerated]
	private ToolStripMenuItem m_P;

	[AccessedThroughProperty("chkHighlighterOrange")]
	[CompilerGenerated]
	private ToolStripMenuItem m_Q;

	[CompilerGenerated]
	[AccessedThroughProperty("chkHighlighterPink")]
	private ToolStripMenuItem m_R;

	[CompilerGenerated]
	[AccessedThroughProperty("chkHighlighterBlue")]
	private ToolStripMenuItem m_S;

	[CompilerGenerated]
	[AccessedThroughProperty("chkHighlighterGreen")]
	private ToolStripMenuItem m_T;

	[CompilerGenerated]
	[AccessedThroughProperty("colMessages")]
	private ColumnHeader m_F;

	[CompilerGenerated]
	[AccessedThroughProperty("chkEmbed")]
	private System.Windows.Forms.CheckBox m_E;

	[AccessedThroughProperty("btnFileEmbed")]
	[CompilerGenerated]
	private ToolStripMenuItem m_U;

	[CompilerGenerated]
	[AccessedThroughProperty("flpMessages")]
	private FlowLayoutPanel m_A;

	[AccessedThroughProperty("chkControls")]
	[CompilerGenerated]
	private System.Windows.Forms.CheckBox m_F;

	[AccessedThroughProperty("flpControls")]
	[CompilerGenerated]
	private FlowLayoutPanel m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("pnlControls")]
	private Panel m_C;

	[AccessedThroughProperty("TableLayoutPanel2")]
	[CompilerGenerated]
	private TableLayoutPanel m_A;

	[AccessedThroughProperty("btnImageCopy")]
	[CompilerGenerated]
	private ToolStripMenuItem m_V;

	[CompilerGenerated]
	[AccessedThroughProperty("TableLayoutPanel1")]
	private TableLayoutPanel m_B;

	[CompilerGenerated]
	[AccessedThroughProperty("btnScreenShot")]
	private Button m_F;

	[AccessedThroughProperty("btnLink")]
	[CompilerGenerated]
	private Button m_G;

	private Microsoft.Office.Interop.Excel.Application m_A;

	private string m_A;

	private Dictionary<Balloon, EH> m_A;

	private Dictionary<string, FH> m_A;

	private int m_A;

	private System.Timers.Timer m_A;

	private readonly int m_B;

	private readonly string m_B;

	private readonly string m_C;

	private readonly string m_D;

	private readonly string m_E;

	private readonly string m_F;

	private readonly string m_G;

	private readonly string m_H;

	private readonly string m_I;

	private readonly string m_J;

	private readonly string m_K;

	private readonly int m_C;

	private readonly int m_D;

	private readonly int m_E;

	private readonly int m_F;

	private readonly string m_L;

	private readonly Color m_A;

	private readonly Color m_B;

	private readonly Color m_C;

	private readonly Color m_D;

	private readonly double m_A;

	private readonly int m_G;

	private readonly int m_H;

	private bool m_A;

	private System.Drawing.Point m_A;

	private System.Drawing.Point m_B;

	private List<Rectangle> m_A;

	private bool m_B;

	private bool m_C;

	private Color m_E;

	private int m_I;

	private Color m_F;

	private Pen m_A;

	private SolidBrush m_A;

	private readonly Color m_G;

	private readonly Color m_H;

	private readonly Color m_I;

	private readonly Color m_J;

	private readonly int m_J;

	private readonly int m_K;

	private readonly int m_L;

	private readonly Color m_K;

	private readonly Color m_L;

	private readonly Color m_M;

	private readonly Color m_N;

	private readonly Color m_O;

	internal virtual ListView lvDiscussions
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
			ColumnClickEventHandler value2 = A;
			ListView listView = this.m_A;
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
				listView.ColumnClick -= value2;
			}
			this.m_A = value;
			listView = this.m_A;
			if (listView != null)
			{
				listView.ColumnClick += value2;
			}
		}
	}

	internal virtual ColumnHeader colSheet
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

	internal virtual ColumnHeader colCell
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

	internal virtual ColumnHeader colValue
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

	internal virtual ColumnHeader colUser
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

	internal virtual SplitContainer SplitContainer1
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

	internal virtual ColumnHeader colDate
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

	internal virtual Button btnNew
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
			EventHandler value2 = E;
			Button button = this.m_A;
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

	internal virtual Button btnDelete
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
			EventHandler value2 = F;
			Button button = this.m_B;
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
			this.m_B = value;
			button = this.m_B;
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

	internal virtual RichTextBox rtbComment
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
			KeyEventHandler value2 = A;
			KeyEventHandler value3 = B;
			RichTextBox richTextBox = this.m_A;
			if (richTextBox != null)
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
				richTextBox.KeyDown -= value2;
				richTextBox.KeyUp -= value3;
			}
			this.m_A = value;
			richTextBox = this.m_A;
			if (richTextBox == null)
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
				richTextBox.KeyDown += value2;
				richTextBox.KeyUp += value3;
				return;
			}
		}
	}

	internal virtual Panel pnlRichTextBox
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

	internal virtual Panel Panel2
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

	internal virtual ToolTip ToolTip1
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

	internal virtual ImageList Icons
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

	internal virtual ContextMenuStrip cmsFile
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
			CancelEventHandler value2 = B;
			ContextMenuStrip contextMenuStrip = this.m_A;
			if (contextMenuStrip != null)
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
				contextMenuStrip.Opening -= value2;
			}
			this.m_A = value;
			contextMenuStrip = this.m_A;
			if (contextMenuStrip != null)
			{
				contextMenuStrip.Opening += value2;
			}
		}
	}

	internal virtual ToolStripMenuItem btnFileOpen
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
			EventHandler value2 = Q;
			ToolStripMenuItem toolStripMenuItem = this.m_A;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_A = value;
			toolStripMenuItem = this.m_A;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem btnFileDelete
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
			EventHandler value2 = O;
			ToolStripMenuItem toolStripMenuItem = this.m_B;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_B = value;
			toolStripMenuItem = this.m_B;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem btnFileShow
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
			EventHandler value2 = S;
			ToolStripMenuItem toolStripMenuItem = this.m_C;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click -= value2;
			}
			this.m_C = value;
			toolStripMenuItem = this.m_C;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ContextMenuStrip cmsLink
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

	internal virtual ToolStripMenuItem btnLinkFollow
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
			EventHandler value2 = T;
			ToolStripMenuItem toolStripMenuItem = this.m_D;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_D = value;
			toolStripMenuItem = this.m_D;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click += value2;
			}
		}
	}

	internal virtual ToolStripMenuItem btnLinkDelete
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
			EventHandler value2 = P;
			ToolStripMenuItem toolStripMenuItem = this.m_E;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_E = value;
			toolStripMenuItem = this.m_E;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnFiles
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
			EventHandler value2 = J;
			System.Windows.Forms.DragEventHandler value3 = A;
			System.Windows.Forms.DragEventHandler value4 = B;
			Button button = this.m_C;
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
				button.DragDrop -= value3;
				button.DragEnter -= value4;
			}
			this.m_C = value;
			button = this.m_C;
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
				button.DragDrop += value3;
				button.DragEnter += value4;
				return;
			}
		}
	}

	internal virtual Button btnMessage
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
			MouseEventHandler value2 = I;
			Button button = this.m_D;
			if (button != null)
			{
				button.MouseDown -= value2;
			}
			this.m_D = value;
			button = this.m_D;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				button.MouseDown += value2;
				return;
			}
		}
	}

	internal virtual ContextMenuStrip cmsComment
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
			CancelEventHandler value2 = A;
			ContextMenuStrip contextMenuStrip = this.m_C;
			if (contextMenuStrip != null)
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
				contextMenuStrip.Opening -= value2;
			}
			this.m_C = value;
			contextMenuStrip = this.m_C;
			if (contextMenuStrip == null)
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
				contextMenuStrip.Opening += value2;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem btnCommentDelete
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
			EventHandler value2 = N;
			ToolStripMenuItem toolStripMenuItem = this.m_F;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_F = value;
			toolStripMenuItem = this.m_F;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual Button btnDeleteAll
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
			EventHandler value2 = G;
			Button button = this.m_E;
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

	internal virtual ContextMenuStrip cmsPicture
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

	internal virtual ToolStripMenuItem btnImageView
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
			EventHandler value2 = W;
			ToolStripMenuItem toolStripMenuItem = this.m_G;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click -= value2;
			}
			this.m_G = value;
			toolStripMenuItem = this.m_G;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem btnImageDelete
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
			EventHandler value2 = X;
			ToolStripMenuItem toolStripMenuItem = this.m_H;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_H = value;
			toolStripMenuItem = this.m_H;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Forms.CheckBox chkPen
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
			EventHandler value2 = Y;
			System.Windows.Forms.CheckBox checkBox = this.m_A;
			if (checkBox != null)
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
				checkBox.Click -= value2;
			}
			this.m_A = value;
			checkBox = this.m_A;
			if (checkBox != null)
			{
				checkBox.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Forms.CheckBox chkHighlighter
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
			EventHandler value2 = Z;
			System.Windows.Forms.CheckBox checkBox = this.m_B;
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
				checkBox.Click -= value2;
			}
			this.m_B = value;
			checkBox = this.m_B;
			if (checkBox != null)
			{
				checkBox.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Forms.CheckBox chkPenMenu
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
			EventHandler value2 = AB;
			System.Windows.Forms.CheckBox checkBox = this.m_C;
			if (checkBox != null)
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
				checkBox.Click -= value2;
			}
			this.m_C = value;
			checkBox = this.m_C;
			if (checkBox != null)
			{
				checkBox.Click += value2;
			}
		}
	}

	internal virtual System.Windows.Forms.CheckBox chkHighlighterMenu
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
			EventHandler value2 = BB;
			System.Windows.Forms.CheckBox checkBox = this.m_D;
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
				checkBox.Click -= value2;
			}
			this.m_D = value;
			checkBox = this.m_D;
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
				checkBox.Click += value2;
				return;
			}
		}
	}

	internal virtual ContextMenuStrip cmsPens
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
			ToolStripDropDownClosingEventHandler value2 = A;
			CancelEventHandler value3 = C;
			ContextMenuStrip contextMenuStrip = this.m_E;
			if (contextMenuStrip != null)
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
				contextMenuStrip.Closing -= value2;
				contextMenuStrip.Opening -= value3;
			}
			this.m_E = value;
			contextMenuStrip = this.m_E;
			if (contextMenuStrip != null)
			{
				contextMenuStrip.Closing += value2;
				contextMenuStrip.Opening += value3;
			}
		}
	}

	internal virtual ToolStripMenuItem chkPenRed
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
			EventHandler value2 = CB;
			ToolStripMenuItem toolStripMenuItem = this.m_I;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click -= value2;
			}
			this.m_I = value;
			toolStripMenuItem = this.m_I;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem chkPenGreen
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
			EventHandler value2 = CB;
			ToolStripMenuItem toolStripMenuItem = this.m_J;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_J = value;
			toolStripMenuItem = this.m_J;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click += value2;
			}
		}
	}

	internal virtual ToolStripMenuItem chkPenBlue
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
			EventHandler value2 = CB;
			ToolStripMenuItem toolStripMenuItem = this.m_K;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_K = value;
			toolStripMenuItem = this.m_K;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem chkPenYellow
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
			EventHandler value2 = CB;
			ToolStripMenuItem toolStripMenuItem = this.m_L;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click -= value2;
			}
			this.m_L = value;
			toolStripMenuItem = this.m_L;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click += value2;
			}
		}
	}

	internal virtual ToolStripSeparator ToolStripSeparator1
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

	internal virtual ToolStripMenuItem chkPenThin
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
			EventHandler value2 = DB;
			ToolStripMenuItem toolStripMenuItem = this.m_M;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_M = value;
			toolStripMenuItem = this.m_M;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click += value2;
			}
		}
	}

	internal virtual ToolStripMenuItem chkPenMedium
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
			EventHandler value2 = DB;
			ToolStripMenuItem toolStripMenuItem = this.m_N;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_N = value;
			toolStripMenuItem = this.m_N;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click += value2;
			}
		}
	}

	internal virtual ToolStripMenuItem chkPenThick
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
			EventHandler value2 = DB;
			ToolStripMenuItem toolStripMenuItem = this.m_O;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_O = value;
			toolStripMenuItem = this.m_O;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ContextMenuStrip cmsHighlighters
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
			ToolStripDropDownClosingEventHandler value2 = B;
			CancelEventHandler value3 = D;
			ContextMenuStrip contextMenuStrip = this.m_F;
			if (contextMenuStrip != null)
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
				contextMenuStrip.Closing -= value2;
				contextMenuStrip.Opening -= value3;
			}
			this.m_F = value;
			contextMenuStrip = this.m_F;
			if (contextMenuStrip == null)
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
				contextMenuStrip.Closing += value2;
				contextMenuStrip.Opening += value3;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem chkHighlighterYellow
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
			EventHandler value2 = EB;
			ToolStripMenuItem toolStripMenuItem = this.m_P;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_P = value;
			toolStripMenuItem = this.m_P;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem chkHighlighterOrange
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
			EventHandler value2 = EB;
			ToolStripMenuItem toolStripMenuItem = this.m_Q;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_Q = value;
			toolStripMenuItem = this.m_Q;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem chkHighlighterPink
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
			EventHandler value2 = EB;
			ToolStripMenuItem toolStripMenuItem = this.m_R;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_R = value;
			toolStripMenuItem = this.m_R;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem chkHighlighterBlue
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
			EventHandler value2 = EB;
			ToolStripMenuItem toolStripMenuItem = this.m_S;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_S = value;
			toolStripMenuItem = this.m_S;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ToolStripMenuItem chkHighlighterGreen
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
			EventHandler value2 = EB;
			ToolStripMenuItem toolStripMenuItem = this.m_T;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_T = value;
			toolStripMenuItem = this.m_T;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual ColumnHeader colMessages
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

	internal virtual System.Windows.Forms.CheckBox chkEmbed
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

	internal virtual ToolStripMenuItem btnFileEmbed
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
			EventHandler value2 = R;
			ToolStripMenuItem toolStripMenuItem = this.m_U;
			if (toolStripMenuItem != null)
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
				toolStripMenuItem.Click -= value2;
			}
			this.m_U = value;
			toolStripMenuItem = this.m_U;
			if (toolStripMenuItem == null)
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
				toolStripMenuItem.Click += value2;
				return;
			}
		}
	}

	internal virtual FlowLayoutPanel flpMessages
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
			EventHandler value2 = GB;
			EventHandler value3 = HB;
			FlowLayoutPanel flowLayoutPanel = this.m_A;
			if (flowLayoutPanel != null)
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
				flowLayoutPanel.MouseEnter -= value2;
				flowLayoutPanel.Resize -= value3;
			}
			this.m_A = value;
			flowLayoutPanel = this.m_A;
			if (flowLayoutPanel == null)
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
				flowLayoutPanel.MouseEnter += value2;
				flowLayoutPanel.Resize += value3;
				return;
			}
		}
	}

	internal virtual System.Windows.Forms.CheckBox chkControls
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
			EventHandler value2 = FB;
			System.Windows.Forms.CheckBox checkBox = this.m_F;
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
				checkBox.CheckedChanged -= value2;
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
				checkBox.CheckedChanged += value2;
				return;
			}
		}
	}

	internal virtual FlowLayoutPanel flpControls
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

	internal virtual Panel pnlControls
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

	internal virtual TableLayoutPanel TableLayoutPanel2
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

	internal virtual ToolStripMenuItem btnImageCopy
	{
		[CompilerGenerated]
		get
		{
			return this.m_V;
		}
		[MethodImpl(MethodImplOptions.Synchronized)]
		[CompilerGenerated]
		set
		{
			EventHandler value2 = U;
			ToolStripMenuItem toolStripMenuItem = this.m_V;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click -= value2;
			}
			this.m_V = value;
			toolStripMenuItem = this.m_V;
			if (toolStripMenuItem != null)
			{
				toolStripMenuItem.Click += value2;
			}
		}
	}

	internal virtual TableLayoutPanel TableLayoutPanel1
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

	internal virtual Button btnScreenShot
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
			EventHandler value2 = M;
			Button button = this.m_F;
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
				switch (5)
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

	internal virtual Button btnLink
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
			EventHandler value2 = L;
			Button button = this.m_G;
			if (button != null)
			{
				button.Click -= value2;
			}
			this.m_G = value;
			button = this.m_G;
			if (button != null)
			{
				button.Click += value2;
			}
		}
	}

	public ctpDiscuss()
	{
		//IL_02b2: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b8: Expected O, but got Unknown
		base.Load += A;
		base.Disposed += C;
		this.m_A = -1;
		this.m_B = 0;
		this.m_B = VH.A(204576);
		this.m_C = VH.A(198035);
		this.m_D = VH.A(204585);
		this.m_E = VH.A(19019);
		this.m_F = VH.A(102634);
		this.m_G = VH.A(102662);
		this.m_H = VH.A(204600);
		this.m_I = VH.A(198664);
		this.m_J = VH.A(204613);
		this.m_K = VH.A(204630);
		this.m_C = 1;
		this.m_D = 2;
		this.m_E = 3;
		this.m_F = 4;
		this.m_L = VH.A(204645);
		this.m_A = Color.FromKnownColor(KnownColor.ActiveBorder);
		this.m_B = Color.FromKnownColor(KnownColor.ControlLight);
		this.m_C = Color.FromKnownColor(KnownColor.Control);
		this.m_D = Color.FromArgb(255, 255, 225);
		this.m_A = 700.0;
		this.m_G = 12;
		this.m_H = 13;
		this.m_A = false;
		this.m_A = new List<Rectangle>();
		this.m_B = false;
		this.m_C = false;
		this.m_G = Color.Red;
		this.m_H = Color.Lime;
		this.m_I = Color.Blue;
		this.m_J = Color.Yellow;
		this.m_J = 1;
		this.m_K = 2;
		this.m_L = 3;
		this.m_K = Color.Yellow;
		this.m_L = Color.Orange;
		this.m_M = Color.HotPink;
		this.m_N = Color.Cyan;
		this.m_O = Color.Chartreuse;
		System.Windows.Forms.Application.EnableVisualStyles();
		A();
		clsDisplay.RescaleImages(new List<Control>(new Control[4] { btnMessage, btnFiles, btnLink, btnScreenShot }), 24, 24);
		clsDisplay.RescaleImages(new List<Control>(new Control[2] { chkPen, chkHighlighter }), 16, 16);
		clsDisplay.RescaleImages(new List<Control>(new Control[2] { chkPenMenu, chkHighlighterMenu }), 8, 5);
		clsDisplay val = new clsDisplay();
		checked
		{
			ImageList imageList;
			if (val.Y == 1.0)
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
				imageList = Icons;
			}
			else
			{
				imageList = new ImageList
				{
					ImageSize = new System.Drawing.Size((int)Math.Round(16.0 * val.X), (int)Math.Round(16.0 * val.Y))
				};
				ImageList icons = Icons;
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = icons.Images.GetEnumerator();
					while (enumerator.MoveNext())
					{
						Image value = (Image)enumerator.Current;
						imageList.Images.Add(value);
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
				icons = null;
			}
			lvDiscussions.SmallImageList = imageList;
			SplitContainer1.SplitterDistance = (int)Math.Round((double)SplitContainer1.SplitterDistance * val.Y);
			val = null;
			imageList = null;
			this.m_A = MH.A.Application;
			this.m_A = new Dictionary<string, FH>();
			try
			{
				clsApis.ConfigureListView(lvDiscussions);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			I();
		}
	}

	[DebuggerNonUserCode]
	protected override void Dispose(bool disposing)
	{
		try
		{
			if (!disposing)
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
				if (this.m_A == null)
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
					this.m_A.Dispose();
					return;
				}
			}
		}
		finally
		{
			base.Dispose(disposing);
		}
	}

	[DebuggerStepThrough]
	private void A()
	{
		this.m_A = new Container();
		ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof(ctpDiscuss));
		lvDiscussions = new ListView();
		colUser = new ColumnHeader();
		colDate = new ColumnHeader();
		colMessages = new ColumnHeader();
		colSheet = new ColumnHeader();
		colCell = new ColumnHeader();
		colValue = new ColumnHeader();
		Icons = new ImageList(this.m_A);
		SplitContainer1 = new SplitContainer();
		btnDeleteAll = new Button();
		btnDelete = new Button();
		btnNew = new Button();
		TableLayoutPanel2 = new TableLayoutPanel();
		flpControls = new FlowLayoutPanel();
		chkControls = new System.Windows.Forms.CheckBox();
		pnlControls = new Panel();
		btnMessage = new Button();
		btnFiles = new Button();
		TableLayoutPanel1 = new TableLayoutPanel();
		btnScreenShot = new Button();
		btnLink = new Button();
		chkPen = new System.Windows.Forms.CheckBox();
		chkHighlighter = new System.Windows.Forms.CheckBox();
		chkPenMenu = new System.Windows.Forms.CheckBox();
		chkHighlighterMenu = new System.Windows.Forms.CheckBox();
		pnlRichTextBox = new Panel();
		Panel2 = new Panel();
		rtbComment = new RichTextBox();
		chkEmbed = new System.Windows.Forms.CheckBox();
		flpMessages = new FlowLayoutPanel();
		ToolTip1 = new ToolTip(this.m_A);
		cmsFile = new ContextMenuStrip(this.m_A);
		btnFileOpen = new ToolStripMenuItem();
		btnFileDelete = new ToolStripMenuItem();
		btnFileEmbed = new ToolStripMenuItem();
		btnFileShow = new ToolStripMenuItem();
		cmsLink = new ContextMenuStrip(this.m_A);
		btnLinkFollow = new ToolStripMenuItem();
		btnLinkDelete = new ToolStripMenuItem();
		cmsComment = new ContextMenuStrip(this.m_A);
		btnCommentDelete = new ToolStripMenuItem();
		cmsPicture = new ContextMenuStrip(this.m_A);
		btnImageCopy = new ToolStripMenuItem();
		btnImageView = new ToolStripMenuItem();
		btnImageDelete = new ToolStripMenuItem();
		cmsPens = new ContextMenuStrip(this.m_A);
		chkPenRed = new ToolStripMenuItem();
		chkPenGreen = new ToolStripMenuItem();
		chkPenBlue = new ToolStripMenuItem();
		chkPenYellow = new ToolStripMenuItem();
		ToolStripSeparator1 = new ToolStripSeparator();
		chkPenThin = new ToolStripMenuItem();
		chkPenMedium = new ToolStripMenuItem();
		chkPenThick = new ToolStripMenuItem();
		cmsHighlighters = new ContextMenuStrip(this.m_A);
		chkHighlighterYellow = new ToolStripMenuItem();
		chkHighlighterOrange = new ToolStripMenuItem();
		chkHighlighterPink = new ToolStripMenuItem();
		chkHighlighterBlue = new ToolStripMenuItem();
		chkHighlighterGreen = new ToolStripMenuItem();
		((ISupportInitialize)SplitContainer1).BeginInit();
		SplitContainer1.Panel1.SuspendLayout();
		SplitContainer1.Panel2.SuspendLayout();
		SplitContainer1.SuspendLayout();
		TableLayoutPanel2.SuspendLayout();
		flpControls.SuspendLayout();
		pnlControls.SuspendLayout();
		TableLayoutPanel1.SuspendLayout();
		pnlRichTextBox.SuspendLayout();
		Panel2.SuspendLayout();
		cmsFile.SuspendLayout();
		cmsLink.SuspendLayout();
		cmsComment.SuspendLayout();
		cmsPicture.SuspendLayout();
		cmsPens.SuspendLayout();
		cmsHighlighters.SuspendLayout();
		SuspendLayout();
		lvDiscussions.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
		lvDiscussions.Columns.AddRange(new ColumnHeader[6] { colUser, colDate, colMessages, colSheet, colCell, colValue });
		lvDiscussions.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		lvDiscussions.FullRowSelect = true;
		lvDiscussions.Location = new System.Drawing.Point(0, 0);
		lvDiscussions.MultiSelect = false;
		lvDiscussions.Name = VH.A(198637);
		lvDiscussions.Size = new System.Drawing.Size(404, 121);
		lvDiscussions.SmallImageList = Icons;
		lvDiscussions.TabIndex = 0;
		lvDiscussions.UseCompatibleStateImageBehavior = false;
		lvDiscussions.View = System.Windows.Forms.View.Details;
		colUser.Text = VH.A(198664);
		colDate.Text = VH.A(198673);
		colMessages.Text = VH.A(49303);
		colMessages.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
		colSheet.Text = VH.A(175976);
		colCell.Text = VH.A(198696);
		colValue.Text = VH.A(41636);
		Icons.ImageStream = (ImageListStreamer)componentResourceManager.GetObject(VH.A(198705));
		Icons.TransparentColor = Color.Transparent;
		Icons.Images.SetKeyName(0, VH.A(198740));
		Icons.Images.SetKeyName(1, VH.A(198755));
		Icons.Images.SetKeyName(2, VH.A(198780));
		Icons.Images.SetKeyName(3, VH.A(198813));
		Icons.Images.SetKeyName(4, VH.A(198838));
		Icons.Images.SetKeyName(5, VH.A(198861));
		Icons.Images.SetKeyName(6, VH.A(198878));
		Icons.Images.SetKeyName(7, VH.A(198909));
		Icons.Images.SetKeyName(8, VH.A(198940));
		Icons.Images.SetKeyName(9, VH.A(198987));
		Icons.Images.SetKeyName(10, VH.A(199024));
		Icons.Images.SetKeyName(11, VH.A(199051));
		Icons.Images.SetKeyName(12, VH.A(199068));
		Icons.Images.SetKeyName(13, VH.A(199097));
		SplitContainer1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
		SplitContainer1.FixedPanel = FixedPanel.Panel1;
		SplitContainer1.Location = new System.Drawing.Point(3, 3);
		SplitContainer1.Name = VH.A(199130);
		SplitContainer1.Orientation = Orientation.Horizontal;
		SplitContainer1.Panel1.Controls.Add(btnDeleteAll);
		SplitContainer1.Panel1.Controls.Add(btnDelete);
		SplitContainer1.Panel1.Controls.Add(lvDiscussions);
		SplitContainer1.Panel1.Controls.Add(btnNew);
		SplitContainer1.Panel2.AutoScroll = true;
		SplitContainer1.Panel2.Controls.Add(TableLayoutPanel2);
		SplitContainer1.Size = new System.Drawing.Size(404, 958);
		SplitContainer1.SplitterDistance = 156;
		SplitContainer1.TabIndex = 1;
		btnDeleteAll.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		btnDeleteAll.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		btnDeleteAll.Location = new System.Drawing.Point(329, 127);
		btnDeleteAll.Name = VH.A(199161);
		btnDeleteAll.Size = new System.Drawing.Size(75, 27);
		btnDeleteAll.TabIndex = 5;
		btnDeleteAll.Text = VH.A(199186);
		ToolTip1.SetToolTip(btnDeleteAll, VH.A(199207));
		btnDeleteAll.UseVisualStyleBackColor = true;
		btnDelete.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		btnDelete.Enabled = false;
		btnDelete.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		btnDelete.Location = new System.Drawing.Point(248, 127);
		btnDelete.Name = VH.A(199286);
		btnDelete.Size = new System.Drawing.Size(75, 27);
		btnDelete.TabIndex = 3;
		btnDelete.Text = VH.A(60691);
		ToolTip1.SetToolTip(btnDelete, VH.A(199305));
		btnDelete.UseVisualStyleBackColor = true;
		btnNew.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
		btnNew.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		btnNew.Location = new System.Drawing.Point(0, 127);
		btnNew.Name = VH.A(199366);
		btnNew.Size = new System.Drawing.Size(109, 27);
		btnNew.TabIndex = 2;
		btnNew.Text = VH.A(199379);
		ToolTip1.SetToolTip(btnNew, VH.A(199408));
		btnNew.UseVisualStyleBackColor = true;
		TableLayoutPanel2.ColumnCount = 1;
		TableLayoutPanel2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
		TableLayoutPanel2.Controls.Add(flpControls, 0, 1);
		TableLayoutPanel2.Controls.Add(flpMessages, 0, 0);
		TableLayoutPanel2.Dock = DockStyle.Fill;
		TableLayoutPanel2.Location = new System.Drawing.Point(0, 0);
		TableLayoutPanel2.Margin = new Padding(0);
		TableLayoutPanel2.Name = VH.A(199509);
		TableLayoutPanel2.RowCount = 2;
		TableLayoutPanel2.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
		TableLayoutPanel2.RowStyles.Add(new RowStyle());
		TableLayoutPanel2.Size = new System.Drawing.Size(404, 798);
		TableLayoutPanel2.TabIndex = 0;
		flpControls.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
		flpControls.AutoSize = true;
		flpControls.AutoSizeMode = AutoSizeMode.GrowAndShrink;
		flpControls.BackColor = System.Drawing.SystemColors.Control;
		flpControls.Controls.Add(chkControls);
		flpControls.Controls.Add(pnlControls);
		flpControls.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
		flpControls.Location = new System.Drawing.Point(0, 525);
		flpControls.Margin = new Padding(0);
		flpControls.Name = VH.A(199544);
		flpControls.Size = new System.Drawing.Size(404, 273);
		flpControls.TabIndex = 15;
		chkControls.Appearance = Appearance.Button;
		chkControls.Checked = true;
		chkControls.CheckState = CheckState.Checked;
		chkControls.FlatAppearance.BorderSize = 0;
		chkControls.FlatAppearance.CheckedBackColor = System.Drawing.SystemColors.Control;
		chkControls.FlatAppearance.MouseDownBackColor = System.Drawing.SystemColors.ActiveBorder;
		chkControls.FlatAppearance.MouseOverBackColor = System.Drawing.SystemColors.ControlLight;
		chkControls.FlatStyle = FlatStyle.Flat;
		chkControls.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		chkControls.Image = global::A.J.TreeNodeExpanded;
		chkControls.ImageAlign = ContentAlignment.MiddleLeft;
		chkControls.Location = new System.Drawing.Point(0, 0);
		chkControls.Margin = new Padding(0);
		chkControls.Name = VH.A(199567);
		chkControls.Size = new System.Drawing.Size(404, 27);
		chkControls.TabIndex = 14;
		chkControls.Text = VH.A(199590);
		chkControls.UseVisualStyleBackColor = true;
		pnlControls.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		pnlControls.BackColor = System.Drawing.SystemColors.Control;
		pnlControls.Controls.Add(btnMessage);
		pnlControls.Controls.Add(btnFiles);
		pnlControls.Controls.Add(TableLayoutPanel1);
		pnlControls.Controls.Add(chkPen);
		pnlControls.Controls.Add(chkHighlighter);
		pnlControls.Controls.Add(chkPenMenu);
		pnlControls.Controls.Add(chkHighlighterMenu);
		pnlControls.Controls.Add(pnlRichTextBox);
		pnlControls.Controls.Add(chkEmbed);
		pnlControls.Location = new System.Drawing.Point(0, 27);
		pnlControls.Margin = new Padding(0);
		pnlControls.Name = VH.A(199637);
		pnlControls.Size = new System.Drawing.Size(404, 246);
		pnlControls.TabIndex = 15;
		btnMessage.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		btnMessage.BackColor = Color.White;
		btnMessage.Enabled = false;
		btnMessage.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
		btnMessage.FlatAppearance.MouseDownBackColor = Color.FromArgb(157, 214, 182);
		btnMessage.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 240, 224);
		btnMessage.FlatStyle = FlatStyle.Flat;
		btnMessage.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		btnMessage.Image = (Image)componentResourceManager.GetObject(VH.A(199660));
		btnMessage.ImageAlign = ContentAlignment.TopCenter;
		btnMessage.Location = new System.Drawing.Point(0, 0);
		btnMessage.Margin = new Padding(0, 3, 0, 3);
		btnMessage.Name = VH.A(199693);
		btnMessage.Padding = new Padding(0, 5, 0, 5);
		btnMessage.Size = new System.Drawing.Size(404, 66);
		btnMessage.TabIndex = 0;
		btnMessage.Text = VH.A(199714);
		btnMessage.TextAlign = ContentAlignment.BottomCenter;
		btnMessage.UseVisualStyleBackColor = false;
		btnFiles.AllowDrop = true;
		btnFiles.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		btnFiles.BackColor = Color.White;
		btnFiles.Enabled = false;
		btnFiles.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
		btnFiles.FlatAppearance.MouseDownBackColor = Color.FromArgb(157, 214, 182);
		btnFiles.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 240, 224);
		btnFiles.FlatStyle = FlatStyle.Flat;
		btnFiles.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		btnFiles.Image = (Image)componentResourceManager.GetObject(VH.A(199737));
		btnFiles.ImageAlign = ContentAlignment.TopCenter;
		btnFiles.Location = new System.Drawing.Point(0, 72);
		btnFiles.Margin = new Padding(0, 3, 0, 3);
		btnFiles.Name = VH.A(199766);
		btnFiles.Padding = new Padding(0, 5, 0, 5);
		btnFiles.Size = new System.Drawing.Size(404, 66);
		btnFiles.TabIndex = 8;
		btnFiles.Text = VH.A(199783);
		btnFiles.TextAlign = ContentAlignment.BottomCenter;
		btnFiles.UseVisualStyleBackColor = false;
		TableLayoutPanel1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		TableLayoutPanel1.ColumnCount = 2;
		TableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
		TableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
		TableLayoutPanel1.Controls.Add(btnScreenShot, 1, 0);
		TableLayoutPanel1.Controls.Add(btnLink, 0, 0);
		TableLayoutPanel1.Location = new System.Drawing.Point(0, 141);
		TableLayoutPanel1.Margin = new Padding(0, 3, 0, 3);
		TableLayoutPanel1.Name = VH.A(199856);
		TableLayoutPanel1.RowCount = 1;
		TableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 72f));
		TableLayoutPanel1.Size = new System.Drawing.Size(404, 72);
		TableLayoutPanel1.TabIndex = 13;
		btnScreenShot.BackColor = Color.White;
		btnScreenShot.Dock = DockStyle.Fill;
		btnScreenShot.Enabled = false;
		btnScreenShot.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
		btnScreenShot.FlatAppearance.MouseDownBackColor = Color.FromArgb(157, 214, 182);
		btnScreenShot.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 240, 224);
		btnScreenShot.FlatStyle = FlatStyle.Flat;
		btnScreenShot.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		btnScreenShot.Image = (Image)componentResourceManager.GetObject(VH.A(199891));
		btnScreenShot.ImageAlign = ContentAlignment.TopCenter;
		btnScreenShot.Location = new System.Drawing.Point(205, 3);
		btnScreenShot.Margin = new Padding(3, 3, 0, 3);
		btnScreenShot.Name = VH.A(199930);
		btnScreenShot.Padding = new Padding(0, 5, 0, 5);
		btnScreenShot.Size = new System.Drawing.Size(199, 66);
		btnScreenShot.TabIndex = 10;
		btnScreenShot.Text = VH.A(199957);
		btnScreenShot.TextAlign = ContentAlignment.BottomCenter;
		btnScreenShot.UseVisualStyleBackColor = false;
		btnLink.BackColor = Color.White;
		btnLink.Dock = DockStyle.Fill;
		btnLink.Enabled = false;
		btnLink.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
		btnLink.FlatAppearance.MouseDownBackColor = Color.FromArgb(157, 214, 182);
		btnLink.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 240, 224);
		btnLink.FlatStyle = FlatStyle.Flat;
		btnLink.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		btnLink.Image = (Image)componentResourceManager.GetObject(VH.A(199992));
		btnLink.ImageAlign = ContentAlignment.TopCenter;
		btnLink.Location = new System.Drawing.Point(0, 3);
		btnLink.Margin = new Padding(0, 3, 3, 3);
		btnLink.Name = VH.A(200019);
		btnLink.Padding = new Padding(0, 5, 0, 5);
		btnLink.Size = new System.Drawing.Size(199, 66);
		btnLink.TabIndex = 9;
		btnLink.Text = VH.A(200034);
		btnLink.TextAlign = ContentAlignment.BottomCenter;
		btnLink.UseVisualStyleBackColor = false;
		chkPen.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
		chkPen.Appearance = Appearance.Button;
		chkPen.BackColor = Color.White;
		chkPen.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
		chkPen.FlatAppearance.CheckedBackColor = Color.FromArgb(210, 240, 224);
		chkPen.FlatAppearance.MouseDownBackColor = Color.FromArgb(157, 214, 182);
		chkPen.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 240, 224);
		chkPen.FlatStyle = FlatStyle.Flat;
		chkPen.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		chkPen.Image = (Image)componentResourceManager.GetObject(VH.A(200069));
		chkPen.ImageAlign = ContentAlignment.TopLeft;
		chkPen.Location = new System.Drawing.Point(0, 217);
		chkPen.Name = VH.A(200094);
		chkPen.Size = new System.Drawing.Size(27, 27);
		chkPen.TabIndex = 5;
		ToolTip1.SetToolTip(chkPen, VH.A(200107));
		chkPen.UseVisualStyleBackColor = false;
		chkHighlighter.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
		chkHighlighter.Appearance = Appearance.Button;
		chkHighlighter.BackColor = Color.White;
		chkHighlighter.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
		chkHighlighter.FlatAppearance.CheckedBackColor = Color.FromArgb(210, 240, 224);
		chkHighlighter.FlatAppearance.MouseDownBackColor = Color.FromArgb(157, 214, 182);
		chkHighlighter.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 240, 224);
		chkHighlighter.FlatStyle = FlatStyle.Flat;
		chkHighlighter.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		chkHighlighter.Image = (Image)componentResourceManager.GetObject(VH.A(200174));
		chkHighlighter.ImageAlign = ContentAlignment.TopLeft;
		chkHighlighter.Location = new System.Drawing.Point(48, 217);
		chkHighlighter.Name = VH.A(200215);
		chkHighlighter.Size = new System.Drawing.Size(27, 27);
		chkHighlighter.TabIndex = 6;
		ToolTip1.SetToolTip(chkHighlighter, VH.A(200244));
		chkHighlighter.UseVisualStyleBackColor = false;
		chkPenMenu.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
		chkPenMenu.Appearance = Appearance.Button;
		chkPenMenu.BackColor = Color.White;
		chkPenMenu.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
		chkPenMenu.FlatAppearance.CheckedBackColor = Color.FromArgb(210, 240, 224);
		chkPenMenu.FlatAppearance.MouseDownBackColor = Color.FromArgb(157, 214, 182);
		chkPenMenu.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 240, 224);
		chkPenMenu.FlatStyle = FlatStyle.Flat;
		chkPenMenu.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		chkPenMenu.Image = (Image)componentResourceManager.GetObject(VH.A(200327));
		chkPenMenu.Location = new System.Drawing.Point(26, 217);
		chkPenMenu.Name = VH.A(200360);
		chkPenMenu.Size = new System.Drawing.Size(15, 27);
		chkPenMenu.TabIndex = 7;
		chkPenMenu.UseVisualStyleBackColor = false;
		chkHighlighterMenu.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
		chkHighlighterMenu.Appearance = Appearance.Button;
		chkHighlighterMenu.BackColor = Color.White;
		chkHighlighterMenu.FlatAppearance.BorderColor = System.Drawing.SystemColors.ActiveBorder;
		chkHighlighterMenu.FlatAppearance.CheckedBackColor = Color.FromArgb(210, 240, 224);
		chkHighlighterMenu.FlatAppearance.MouseDownBackColor = Color.FromArgb(157, 214, 182);
		chkHighlighterMenu.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 240, 224);
		chkHighlighterMenu.FlatStyle = FlatStyle.Flat;
		chkHighlighterMenu.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		chkHighlighterMenu.Image = (Image)componentResourceManager.GetObject(VH.A(200381));
		chkHighlighterMenu.Location = new System.Drawing.Point(74, 217);
		chkHighlighterMenu.Name = VH.A(200430);
		chkHighlighterMenu.Size = new System.Drawing.Size(15, 27);
		chkHighlighterMenu.TabIndex = 8;
		chkHighlighterMenu.UseVisualStyleBackColor = false;
		pnlRichTextBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
		pnlRichTextBox.AutoSize = true;
		pnlRichTextBox.BackColor = System.Drawing.SystemColors.ActiveBorder;
		pnlRichTextBox.Controls.Add(Panel2);
		pnlRichTextBox.Location = new System.Drawing.Point(263, 28);
		pnlRichTextBox.Margin = new Padding(0, 0, 0, 6);
		pnlRichTextBox.Name = VH.A(200467);
		pnlRichTextBox.Padding = new Padding(1);
		pnlRichTextBox.Size = new System.Drawing.Size(470, 66);
		pnlRichTextBox.TabIndex = 2;
		pnlRichTextBox.Visible = false;
		Panel2.AutoSize = true;
		Panel2.BackColor = Color.White;
		Panel2.Controls.Add(rtbComment);
		Panel2.Dock = DockStyle.Fill;
		Panel2.Location = new System.Drawing.Point(1, 1);
		Panel2.Margin = new Padding(0);
		Panel2.Name = VH.A(200496);
		Panel2.Size = new System.Drawing.Size(468, 64);
		Panel2.TabIndex = 3;
		rtbComment.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
		rtbComment.BorderStyle = BorderStyle.None;
		rtbComment.Cursor = Cursors.IBeam;
		rtbComment.Font = new System.Drawing.Font(VH.A(50021), 10f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		rtbComment.Location = new System.Drawing.Point(6, 6);
		rtbComment.Margin = new Padding(6);
		rtbComment.Name = VH.A(200509);
		rtbComment.ScrollBars = RichTextBoxScrollBars.Vertical;
		rtbComment.Size = new System.Drawing.Size(456, 52);
		rtbComment.TabIndex = 1;
		rtbComment.Text = "";
		ToolTip1.SetToolTip(rtbComment, VH.A(200530));
		chkEmbed.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		chkEmbed.AutoSize = true;
		chkEmbed.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		chkEmbed.Location = new System.Drawing.Point(316, 222);
		chkEmbed.Name = VH.A(200806);
		chkEmbed.Size = new System.Drawing.Size(87, 19);
		chkEmbed.TabIndex = 9;
		chkEmbed.Text = VH.A(200823);
		ToolTip1.SetToolTip(chkEmbed, VH.A(200846));
		chkEmbed.UseVisualStyleBackColor = true;
		flpMessages.AutoScroll = true;
		flpMessages.AutoSize = true;
		flpMessages.BackColor = System.Drawing.SystemColors.Control;
		flpMessages.Dock = DockStyle.Fill;
		flpMessages.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
		flpMessages.Font = new System.Drawing.Font(VH.A(50021), 9f, System.Drawing.FontStyle.Regular, GraphicsUnit.Point, 0);
		flpMessages.Location = new System.Drawing.Point(0, 0);
		flpMessages.Margin = new Padding(0, 0, 0, 9);
		flpMessages.Name = VH.A(201028);
		flpMessages.Size = new System.Drawing.Size(404, 516);
		flpMessages.TabIndex = 0;
		flpMessages.WrapContents = false;
		cmsFile.Items.AddRange(new ToolStripItem[4] { btnFileOpen, btnFileDelete, btnFileEmbed, btnFileShow });
		cmsFile.Name = VH.A(201051);
		cmsFile.ShowImageMargin = false;
		cmsFile.Size = new System.Drawing.Size(128, 92);
		btnFileOpen.Name = VH.A(201066);
		btnFileOpen.Size = new System.Drawing.Size(127, 22);
		btnFileOpen.Text = VH.A(201089);
		btnFileDelete.Name = VH.A(201108);
		btnFileDelete.Size = new System.Drawing.Size(127, 22);
		btnFileDelete.Text = VH.A(201135);
		btnFileEmbed.Name = VH.A(201158);
		btnFileEmbed.Size = new System.Drawing.Size(127, 22);
		btnFileEmbed.Text = VH.A(201183);
		btnFileEmbed.Visible = false;
		btnFileShow.Name = VH.A(201204);
		btnFileShow.Size = new System.Drawing.Size(127, 22);
		btnFileShow.Text = VH.A(201227);
		cmsLink.Items.AddRange(new ToolStripItem[2] { btnLinkFollow, btnLinkDelete });
		cmsLink.Name = VH.A(201256);
		cmsLink.ShowImageMargin = false;
		cmsLink.Size = new System.Drawing.Size(110, 48);
		btnLinkFollow.Name = VH.A(201271);
		btnLinkFollow.Size = new System.Drawing.Size(109, 22);
		btnLinkFollow.Text = VH.A(201298);
		btnLinkDelete.Name = VH.A(201321);
		btnLinkDelete.Size = new System.Drawing.Size(109, 22);
		btnLinkDelete.Text = VH.A(201348);
		cmsComment.Items.AddRange(new ToolStripItem[1] { btnCommentDelete });
		cmsComment.Name = VH.A(201371);
		cmsComment.ShowImageMargin = false;
		cmsComment.Size = new System.Drawing.Size(140, 26);
		btnCommentDelete.Name = VH.A(201392);
		btnCommentDelete.Size = new System.Drawing.Size(139, 22);
		btnCommentDelete.Text = VH.A(201425);
		cmsPicture.Items.AddRange(new ToolStripItem[3] { btnImageCopy, btnImageView, btnImageDelete });
		cmsPicture.Name = VH.A(201454);
		cmsPicture.ShowImageMargin = false;
		cmsPicture.Size = new System.Drawing.Size(119, 70);
		btnImageCopy.Name = VH.A(201475);
		btnImageCopy.Size = new System.Drawing.Size(118, 22);
		btnImageCopy.Text = VH.A(201500);
		btnImageView.Name = VH.A(201521);
		btnImageView.Size = new System.Drawing.Size(118, 22);
		btnImageView.Text = VH.A(201546);
		btnImageDelete.Name = VH.A(201567);
		btnImageDelete.Size = new System.Drawing.Size(118, 22);
		btnImageDelete.Text = VH.A(201596);
		cmsPens.Items.AddRange(new ToolStripItem[8] { chkPenRed, chkPenGreen, chkPenBlue, chkPenYellow, ToolStripSeparator1, chkPenThin, chkPenMedium, chkPenThick });
		cmsPens.Name = VH.A(201621);
		cmsPens.Size = new System.Drawing.Size(120, 164);
		chkPenRed.CheckOnClick = true;
		chkPenRed.Name = VH.A(201636);
		chkPenRed.Size = new System.Drawing.Size(119, 22);
		chkPenRed.Text = VH.A(201655);
		chkPenGreen.CheckOnClick = true;
		chkPenGreen.Name = VH.A(201664);
		chkPenGreen.Size = new System.Drawing.Size(119, 22);
		chkPenGreen.Text = VH.A(201687);
		chkPenBlue.CheckOnClick = true;
		chkPenBlue.Name = VH.A(201700);
		chkPenBlue.Size = new System.Drawing.Size(119, 22);
		chkPenBlue.Text = VH.A(201721);
		chkPenYellow.CheckOnClick = true;
		chkPenYellow.Name = VH.A(201732);
		chkPenYellow.Size = new System.Drawing.Size(119, 22);
		chkPenYellow.Text = VH.A(201757);
		ToolStripSeparator1.Name = VH.A(201772);
		ToolStripSeparator1.Size = new System.Drawing.Size(116, 6);
		chkPenThin.CheckOnClick = true;
		chkPenThin.Name = VH.A(201811);
		chkPenThin.Size = new System.Drawing.Size(119, 22);
		chkPenThin.Text = VH.A(201832);
		chkPenMedium.CheckOnClick = true;
		chkPenMedium.Name = VH.A(201843);
		chkPenMedium.Size = new System.Drawing.Size(119, 22);
		chkPenMedium.Text = VH.A(201868);
		chkPenThick.CheckOnClick = true;
		chkPenThick.Name = VH.A(201883);
		chkPenThick.Size = new System.Drawing.Size(119, 22);
		chkPenThick.Text = VH.A(201906);
		cmsHighlighters.Items.AddRange(new ToolStripItem[5] { chkHighlighterYellow, chkHighlighterOrange, chkHighlighterPink, chkHighlighterBlue, chkHighlighterGreen });
		cmsHighlighters.Name = VH.A(201919);
		cmsHighlighters.Size = new System.Drawing.Size(114, 114);
		chkHighlighterYellow.CheckOnClick = true;
		chkHighlighterYellow.Name = VH.A(201950);
		chkHighlighterYellow.Size = new System.Drawing.Size(113, 22);
		chkHighlighterYellow.Text = VH.A(201757);
		chkHighlighterOrange.CheckOnClick = true;
		chkHighlighterOrange.Name = VH.A(201991);
		chkHighlighterOrange.Size = new System.Drawing.Size(113, 22);
		chkHighlighterOrange.Text = VH.A(202032);
		chkHighlighterPink.CheckOnClick = true;
		chkHighlighterPink.Name = VH.A(202047);
		chkHighlighterPink.Size = new System.Drawing.Size(113, 22);
		chkHighlighterPink.Text = VH.A(202084);
		chkHighlighterBlue.CheckOnClick = true;
		chkHighlighterBlue.Name = VH.A(202095);
		chkHighlighterBlue.Size = new System.Drawing.Size(113, 22);
		chkHighlighterBlue.Text = VH.A(201721);
		chkHighlighterGreen.CheckOnClick = true;
		chkHighlighterGreen.Name = VH.A(202132);
		chkHighlighterGreen.Size = new System.Drawing.Size(113, 22);
		chkHighlighterGreen.Text = VH.A(201687);
		base.AutoScaleDimensions = new SizeF(6f, 13f);
		base.AutoScaleMode = AutoScaleMode.Font;
		base.Controls.Add(SplitContainer1);
		DoubleBuffered = true;
		base.Name = VH.A(202171);
		base.Size = new System.Drawing.Size(410, 961);
		SplitContainer1.Panel1.ResumeLayout(performLayout: false);
		SplitContainer1.Panel2.ResumeLayout(performLayout: false);
		((ISupportInitialize)SplitContainer1).EndInit();
		SplitContainer1.ResumeLayout(performLayout: false);
		TableLayoutPanel2.ResumeLayout(performLayout: false);
		TableLayoutPanel2.PerformLayout();
		flpControls.ResumeLayout(performLayout: false);
		pnlControls.ResumeLayout(performLayout: false);
		pnlControls.PerformLayout();
		TableLayoutPanel1.ResumeLayout(performLayout: false);
		pnlRichTextBox.ResumeLayout(performLayout: false);
		pnlRichTextBox.PerformLayout();
		Panel2.ResumeLayout(performLayout: false);
		cmsFile.ResumeLayout(performLayout: false);
		cmsLink.ResumeLayout(performLayout: false);
		cmsComment.ResumeLayout(performLayout: false);
		cmsPicture.ResumeLayout(performLayout: false);
		cmsPens.ResumeLayout(performLayout: false);
		cmsHighlighters.ResumeLayout(performLayout: false);
		ResumeLayout(performLayout: false);
	}

	private void A(object A, EventArgs B)
	{
		List<ListViewItem> list = new List<ListViewItem>();
		List<CustomXMLPart> list2 = new List<CustomXMLPart>();
		List<Name> list3 = new List<Name>();
		Microsoft.Office.Interop.Excel.Workbook activeWorkbook = this.m_A.ActiveWorkbook;
		Names names = activeWorkbook.Names;
		this.m_A = this.m_A.UserName;
		IEnumerator enumerator = default(IEnumerator);
		Name name;
		try
		{
			enumerator = activeWorkbook.CustomXMLParts.GetEnumerator();
			string index = default(string);
			string text2 = default(string);
			while (enumerator.MoveNext())
			{
				CustomXMLPart customXMLPart = (CustomXMLPart)enumerator.Current;
				if (!customXMLPart.XML.Contains(this.m_C))
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				int num = 0;
				string d = "";
				bool f = false;
				try
				{
					string text = this.A(customXMLPart);
					index = customXMLPart.SelectSingleNode(text + this.m_E).Text;
					text2 = customXMLPart.SelectSingleNode(text + this.m_F).Text;
					CustomXMLNodes customXMLNodes = customXMLPart.SelectNodes(text + this.m_K);
					_ = null;
					num = customXMLNodes.Count;
					if (num > 0)
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
						d = customXMLNodes[num].Attributes[this.m_D].Text;
						text2 = customXMLNodes[num].Attributes[this.m_C].Text;
						if (Operators.CompareString(text2, this.m_A, TextCompare: false) != 0)
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
							if (this.A(customXMLPart).SelectSingleNode(text + this.m_I + VH.A(202192) + JH.B(this.m_A) + VH.A(43340)) == null)
							{
								f = true;
							}
						}
					}
					customXMLNodes = null;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					clsReporting.LogException(ex2);
					ProjectData.ClearProjectError();
				}
				name = null;
				Range range = null;
				try
				{
					name = names.Item(index, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
					if (name != null)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								continue;
							}
							if (name.RefersToRange != null)
							{
								range = name.RefersToRange;
							}
							else
							{
								list3.Add(name);
							}
							break;
						}
					}
					else
					{
						list3.Add(name);
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
				if (range != null)
				{
					list.Add(this.A(range, customXMLPart, text2, d, num, f));
					range = null;
				}
				else
				{
					list2.Add(customXMLPart);
				}
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_0299;
				}
				continue;
				end_IL_0299:
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
		ListView listView = lvDiscussions;
		listView.BeginUpdate();
		listView.Items.AddRange(list.ToArray());
		listView.EndUpdate();
		_ = null;
		if (list2.Count == 0)
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
			if (!list3.Any())
			{
				goto IL_03e7;
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
		if (System.Windows.Forms.MessageBox.Show(VH.A(202209), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
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
				using (List<CustomXMLPart>.Enumerator enumerator2 = list2.GetEnumerator())
				{
					while (enumerator2.MoveNext())
					{
						enumerator2.Current.Delete();
					}
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							goto end_IL_0376;
						}
						continue;
						end_IL_0376:
						break;
					}
				}
				using List<Name>.Enumerator enumerator3 = list3.GetEnumerator();
				while (enumerator3.MoveNext())
				{
					name = enumerator3.Current;
					name.Delete();
				}
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						goto end_IL_03b5;
					}
					continue;
					end_IL_03b5:
					break;
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				clsReporting.LogException(ex6);
				ProjectData.ClearProjectError();
			}
		}
		goto IL_03e7;
		IL_03e7:
		activeWorkbook = null;
		names = null;
		name = null;
		list = null;
		list3 = null;
		list2 = null;
		chkEmbed.Checked = global::A.K.Settings.DiscussEmbedFiles;
		base.SizeChanged += this.B;
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).AddEventHandler(this.m_A, new AppEvents_SheetSelectionChangeEventHandler(this.A));
		chkEmbed.CheckedChanged += K;
		lvDiscussions.SelectedIndexChanged += D;
	}

	private void B(object A, EventArgs B)
	{
		ListView listView = lvDiscussions;
		listView.BeginUpdate();
		this.B();
		if (listView.CanFocus)
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
			if (listView.Items.Count > 0)
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
				C();
			}
		}
		listView.EndUpdate();
		listView = null;
		chkControls.Width = flpMessages.Width;
		pnlControls.Width = flpMessages.Width;
	}

	private void B()
	{
		lvDiscussions.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
	}

	private void C()
	{
		lvDiscussions.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
		ColumnHeader columnHeader = colDate;
		columnHeader.Width = Math.Max(columnHeader.Width, checked(lvDiscussions.Width - colUser.Width - colSheet.Width - colCell.Width - colValue.Width - colMessages.Width - 4));
		_ = null;
	}

	private void C(object A, EventArgs B)
	{
		new ComAwareEventInfo(typeof(AppEvents_Event), VH.A(1700)).RemoveEventHandler(this.m_A, new AppEvents_SheetSelectionChangeEventHandler(this.A));
		this.m_A = null;
		this.m_A = null;
	}

	private void A(object A, Range B)
	{
		try
		{
			lvDiscussions.SelectedItems.Clear();
			string right = ((Range)this.m_A.Selection).get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = lvDiscussions.Items.GetEnumerator();
				while (enumerator.MoveNext())
				{
					ListViewItem listViewItem = (ListViewItem)enumerator.Current;
					if (Operators.CompareString(((DH)listViewItem.Tag).A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value)), right, TextCompare: false) != 0)
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
						listViewItem.Selected = true;
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void D(object A, EventArgs B)
	{
		if (lvDiscussions.SelectedItems.Count > 0)
		{
			try
			{
				this.m_A.Stop();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			DH dH = this.A();
			CustomXMLPart a = dH.A;
			this.A(a, this.A(a));
			Range a2 = dH.A;
			try
			{
				a2.Worksheet.Activate();
				Ranges.ScrollIntoView(a2);
				a2.Select();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
			H();
			a2 = null;
			a = null;
			this.A(A: true);
			btnDelete.Enabled = true;
		}
		else
		{
			D();
			btnDelete.Enabled = false;
		}
	}

	private void A(object A, ColumnClickEventArgs B)
	{
		ListView listView = lvDiscussions;
		listView.SelectedIndexChanged -= D;
		this.A(lvDiscussions, B, ref this.m_A);
		listView.SelectedIndexChanged += D;
		_ = null;
	}

	private void D()
	{
		flpMessages.Controls.Clear();
		A(A: false);
	}

	private void A(bool A)
	{
		btnMessage.Enabled = A;
		btnFiles.Enabled = A;
		btnLink.Enabled = A;
		btnScreenShot.Enabled = A;
	}

	private void A(CustomXMLPart A, string B)
	{
		List<FileLinkButton> B2 = new List<FileLinkButton>();
		flpMessages.Controls.Clear();
		CustomXMLNodes customXMLNodes = A.SelectNodes(B + this.m_K);
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = customXMLNodes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				CustomXMLNode a = (CustomXMLNode)enumerator.Current;
				this.A(a, ref B2);
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
		customXMLNodes = null;
		E();
		this.A(B2);
		B2 = null;
	}

	private void E()
	{
		if (!B())
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
			N();
			flpMessages.ScrollControlIntoView(flpMessages.Controls[checked(flpMessages.Controls.Count - 1)]);
			return;
		}
	}

	private void A(CustomXMLNode A, [Optional][DefaultParameterValue(null)] ref List<FileLinkButton> B)
	{
		int num = 0;
		string userName = this.m_A.UserName;
		string tag = A.Attributes[this.m_E].Text;
		string text = A.Attributes[this.m_C].Text;
		bool flag = Operators.CompareString(text, userName, TextCompare: false) == 0;
		Balloon balloon = new Balloon();
		balloon.AuthorIsMe = flag;
		checked
		{
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				num = flpMessages.Width - 24 - balloon.Margin.Left - balloon.Margin.Right;
			}
			string text3;
			int index;
			FileLinkButton fileLinkButton3;
			FileLinkButton fileLinkButton;
			PictureBox pictureBox;
			switch (this.A(A))
			{
			case MessageType.Text:
			{
				balloon.BalloonContent = Balloon.BalloonContentEnum.Text;
				RichTextBox richTextBox = new RichTextBox();
				RichTextBox richTextBox2 = richTextBox;
				richTextBox2.Location = this.A(flag);
				richTextBox2.Width = num;
				richTextBox2.Margin = this.A(flag);
				richTextBox2.BorderStyle = BorderStyle.None;
				Color selectionColor;
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
					richTextBox2.BackColor = Color.FromKnownColor(KnownColor.MenuHighlight);
					selectionColor = Color.White;
				}
				else
				{
					richTextBox2.BackColor = Color.Gainsboro;
					selectionColor = Color.FromKnownColor(KnownColor.ControlText);
				}
				richTextBox2.ScrollBars = RichTextBoxScrollBars.None;
				richTextBox2.ContextMenuStrip = cmsComment;
				richTextBox2.Tag = tag;
				richTextBox2.ReadOnly = true;
				richTextBox2.TabStop = false;
				richTextBox2.ContentsResized += this.A;
				richTextBox2 = null;
				balloon.Controls.Add(richTextBox);
				flpMessages.Controls.Add(balloon);
				RichTextBox richTextBox3 = richTextBox;
				richTextBox3.Rtf = A.Text;
				richTextBox3.SelectAll();
				richTextBox3.SelectionColor = selectionColor;
				richTextBox3.DeselectAll();
				_ = null;
				richTextBox = null;
				break;
			}
			case MessageType.File:
			{
				text3 = A.Text;
				balloon.BalloonContent = Balloon.BalloonContentEnum.File;
				ToolTip1.SetToolTip(balloon, text3);
				fileLinkButton = this.A(num, flag);
				fileLinkButton3 = fileLinkButton;
				fileLinkButton3.Text = Path.GetFileName(text3);
				string extension = Path.GetExtension(text3);
				uint num2 = TH.A(extension);
				if (num2 <= 1944878404)
				{
					if (num2 <= 1027080323)
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
						if (num2 <= 285406141)
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
							if (num2 <= 71852739)
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
								if (num2 != 18585163)
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
									if (num2 != 71852739)
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
										if (Operators.CompareString(extension, VH.A(202486), TextCompare: false) == 0)
										{
											goto IL_0c47;
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
								else
								{
									if (Operators.CompareString(extension, VH.A(202551), TextCompare: false) == 0)
									{
										goto IL_0c51;
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
							}
							else if (num2 != 175576948)
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
								if (num2 != 285406141)
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
								}
								else
								{
									if (Operators.CompareString(extension, VH.A(202330), TextCompare: false) == 0)
									{
										goto IL_0c33;
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
							else if (Operators.CompareString(extension, VH.A(202439), TextCompare: false) == 0)
							{
								goto IL_0c42;
							}
						}
						else if (num2 <= 469959950)
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
							if (num2 != 402849474)
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
								if (num2 != 469959950)
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
									if (Operators.CompareString(extension, VH.A(98730), TextCompare: false) == 0)
									{
										goto IL_0c33;
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
							}
							else
							{
								if (Operators.CompareString(extension, VH.A(202341), TextCompare: false) == 0)
								{
									goto IL_0c33;
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
						}
						else if (num2 != 592705037)
						{
							if (num2 != 754654932)
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
								if (num2 == 1027080323)
								{
									if (Operators.CompareString(extension, VH.A(202495), TextCompare: false) == 0)
									{
										goto IL_0c47;
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
							}
							else if (Operators.CompareString(extension, VH.A(202504), TextCompare: false) == 0)
							{
								goto IL_0c4c;
							}
						}
						else if (Operators.CompareString(extension, VH.A(202542), TextCompare: false) == 0)
						{
							goto IL_0c51;
						}
					}
					else if (num2 <= 1616086803)
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
						if (num2 <= 1384894805)
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
							if (num2 != 1128223456)
							{
								if (num2 != 1384894805)
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
									if (Operators.CompareString(extension, VH.A(202410), TextCompare: false) == 0)
									{
										goto IL_0c42;
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
							else if (Operators.CompareString(extension, VH.A(63217), TextCompare: false) == 0)
							{
								goto IL_0c42;
							}
						}
						else if (num2 != 1388056268)
						{
							if (num2 != 1464784447)
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
								if (num2 != 1616086803)
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
								}
								else if (Operators.CompareString(extension, VH.A(202390), TextCompare: false) == 0)
								{
									goto IL_0c3d;
								}
							}
							else
							{
								if (Operators.CompareString(extension, VH.A(202459), TextCompare: false) == 0)
								{
									goto IL_0c42;
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
						}
						else if (Operators.CompareString(extension, VH.A(202419), TextCompare: false) == 0)
						{
							goto IL_0c42;
						}
					}
					else if (num2 <= 1680899145)
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
						if (num2 != 1644092503)
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
							if (num2 != 1680899145)
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
							}
							else
							{
								if (Operators.CompareString(extension, VH.A(97198), TextCompare: false) == 0)
								{
									index = 0;
									goto IL_0c64;
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
						else
						{
							if (Operators.CompareString(extension, VH.A(202381), TextCompare: false) == 0)
							{
								goto IL_0c38;
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
					}
					else if (num2 != 1928100785)
					{
						if (num2 != 1932828157)
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
							if (num2 != 1944878404)
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
							else if (Operators.CompareString(extension, VH.A(202533), TextCompare: false) == 0)
							{
								goto IL_0c51;
							}
						}
						else if (Operators.CompareString(extension, VH.A(202370), TextCompare: false) == 0)
						{
							goto IL_0c38;
						}
					}
					else
					{
						if (Operators.CompareString(extension, VH.A(202468), TextCompare: false) == 0)
						{
							goto IL_0c42;
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
				}
				else if (num2 <= 3182675714u)
				{
					if (num2 <= 2685385760u)
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
						if (num2 <= 2196542689u)
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
							if (num2 != 2194571213u)
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
								if (num2 != 2196542689u)
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
									if (Operators.CompareString(extension, VH.A(202625), TextCompare: false) == 0)
									{
										goto IL_0c56;
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
							else if (Operators.CompareString(extension, VH.A(202524), TextCompare: false) == 0)
							{
								goto IL_0c51;
							}
						}
						else if (num2 != 2641099312u)
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
							if (num2 != 2651941483u)
							{
								if (num2 != 2685385760u)
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
									if (Operators.CompareString(extension, VH.A(202562), TextCompare: false) == 0)
									{
										goto IL_0c51;
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
							else
							{
								if (Operators.CompareString(extension, VH.A(202352), TextCompare: false) == 0)
								{
									goto IL_0c33;
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
						}
						else
						{
							if (Operators.CompareString(extension, VH.A(202515), TextCompare: false) == 0)
							{
								goto IL_0c4c;
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
					}
					else if (num2 <= 3031676203u)
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
						if (num2 != 2928405759u)
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
							if (num2 == 3031676203u)
							{
								if (Operators.CompareString(extension, VH.A(202361), TextCompare: false) == 0)
								{
									goto IL_0c33;
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
						else
						{
							if (Operators.CompareString(extension, VH.A(202643), TextCompare: false) == 0)
							{
								goto IL_0c56;
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
					}
					else if (num2 != 3098210579u)
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
						if (num2 != 3114988198u)
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
							if (num2 != 3182675714u)
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
							}
							else if (Operators.CompareString(extension, VH.A(6144), TextCompare: false) == 0)
							{
								goto IL_0c33;
							}
						}
						else
						{
							if (Operators.CompareString(extension, VH.A(202589), TextCompare: false) == 0)
							{
								goto IL_0c51;
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
					}
					else if (Operators.CompareString(extension, VH.A(202580), TextCompare: false) == 0)
					{
						goto IL_0c51;
					}
				}
				else if (num2 <= 3388221377u)
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
					if (num2 <= 3211133207u)
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
						if (num2 != 3210454886u)
						{
							if (num2 != 3211133207u)
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
							}
							else
							{
								if (Operators.CompareString(extension, VH.A(202634), TextCompare: false) == 0)
								{
									goto IL_0c56;
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
						}
						else if (Operators.CompareString(extension, VH.A(202652), TextCompare: false) == 0)
						{
							index = 9;
							goto IL_0c64;
						}
					}
					else if (num2 != 3238515961u)
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
						if (num2 != 3349874864u)
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
							if (num2 != 3388221377u)
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
								if (Operators.CompareString(extension, VH.A(202616), TextCompare: false) == 0)
								{
									goto IL_0c56;
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
						else
						{
							if (Operators.CompareString(extension, VH.A(202571), TextCompare: false) == 0)
							{
								goto IL_0c51;
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
					}
					else if (Operators.CompareString(extension, VH.A(202401), TextCompare: false) == 0)
					{
						goto IL_0c3d;
					}
				}
				else if (num2 <= 3511907575u)
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
					if (num2 != 3498925813u)
					{
						if (num2 == 3511907575u)
						{
							if (Operators.CompareString(extension, VH.A(202598), TextCompare: false) == 0)
							{
								goto IL_0c51;
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
					else if (Operators.CompareString(extension, VH.A(202607), TextCompare: false) == 0)
					{
						goto IL_0c56;
					}
				}
				else if (num2 != 3534858618u)
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
					if (num2 != 3560597182u)
					{
						if (num2 != 4178554255u)
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
							if (Operators.CompareString(extension, VH.A(202428), TextCompare: false) == 0)
							{
								goto IL_0c42;
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
					else if (Operators.CompareString(extension, VH.A(202448), TextCompare: false) == 0)
					{
						goto IL_0c42;
					}
				}
				else
				{
					if (Operators.CompareString(extension, VH.A(202477), TextCompare: false) == 0)
					{
						goto IL_0c47;
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
				index = 5;
				goto IL_0c64;
			}
			case MessageType.Link:
			{
				string caption = A.Text;
				balloon.BalloonContent = Balloon.BalloonContentEnum.Link;
				ToolTip1.SetToolTip(balloon, caption);
				fileLinkButton = this.A(num, flag);
				FileLinkButton fileLinkButton2 = fileLinkButton;
				fileLinkButton2.Text = caption;
				fileLinkButton2.Icon = Icons.Images[10];
				fileLinkButton2.Tag = tag;
				fileLinkButton2.ContextMenuStrip = cmsLink;
				fileLinkButton2.Storage = clsDiscuss.Storage.WebLink;
				_ = null;
				ToolTip1.SetToolTip(fileLinkButton, caption);
				if (B != null)
				{
					B.Add(fileLinkButton);
				}
				balloon.Controls.Add(fileLinkButton);
				flpMessages.Controls.Add(balloon);
				fileLinkButton = null;
				pictureBox = null;
				break;
			}
			case MessageType.ScreenShot:
				{
					string text2 = A.Text;
					balloon.BalloonContent = Balloon.BalloonContentEnum.ScreenShot;
					pictureBox = this.A(num, flag);
					this.A(pictureBox, text2);
					pictureBox.Height = pictureBox.Image.Height;
					pictureBox.Tag = text2;
					balloon.Controls.Add(pictureBox);
					flpMessages.Controls.Add(balloon);
					pictureBox = null;
					break;
				}
				IL_0c56:
				index = 8;
				goto IL_0c64;
				IL_0c4c:
				index = 6;
				goto IL_0c64;
				IL_0c3d:
				index = 3;
				goto IL_0c64;
				IL_0c47:
				index = 5;
				goto IL_0c64;
				IL_0c42:
				index = 4;
				goto IL_0c64;
				IL_0c38:
				index = 2;
				goto IL_0c64;
				IL_0c33:
				index = 1;
				goto IL_0c64;
				IL_0c51:
				index = 7;
				goto IL_0c64;
				IL_0c64:
				fileLinkButton3.Icon = Icons.Images[index];
				fileLinkButton3.Tag = tag;
				fileLinkButton3.ContextMenuStrip = cmsFile;
				if (!clsFile.IsPathUrl(text3))
				{
					if (this.A(text3))
					{
						fileLinkButton3.Storage = clsDiscuss.Storage.Embedded;
					}
					else if (clsFile.IsNetworkPath(text3))
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
						fileLinkButton3.Storage = clsDiscuss.Storage.Remote;
					}
					else
					{
						fileLinkButton3.Storage = clsDiscuss.Storage.Local;
					}
				}
				else
				{
					fileLinkButton3.Storage = clsDiscuss.Storage.Remote;
				}
				fileLinkButton3 = null;
				ToolTip1.SetToolTip(fileLinkButton, text3);
				balloon.Controls.Add(fileLinkButton);
				flpMessages.Controls.Add(balloon);
				fileLinkButton = null;
				pictureBox = null;
				break;
			}
			string text4 = this.A(A);
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
				text4 = text + VH.A(25588) + text4;
			}
			Label label = new Label();
			Label label2 = label;
			label2.AutoSize = false;
			label2.Width = flpMessages.Width;
			label2.Height = label2.PreferredHeight;
			label2.ForeColor = Color.FromKnownColor(KnownColor.Gray);
			label2.Font = new System.Drawing.Font(Font.FontFamily, 7f, System.Drawing.FontStyle.Bold);
			label2.Text = text4;
			label2.Margin = new Padding(0, 0, 0, 0);
			label2.Padding = new Padding(0, 1, 0, 0);
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
				label2.TextAlign = ContentAlignment.MiddleLeft;
			}
			else
			{
				label2.TextAlign = ContentAlignment.MiddleRight;
			}
			label2 = null;
			flpMessages.Controls.Add(label);
			balloon.Tag = label;
			ControlCollection controls = flpMessages.Controls;
			int num3 = controls.OfType<Label>().Count() - 2;
			while (true)
			{
				if (num3 >= 0)
				{
					label = controls.OfType<Label>().ElementAt(num3);
					if (Operators.CompareString(label.Text, text4, TextCompare: false) == 0)
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
						controls.Remove(label);
						break;
					}
					num3 += -1;
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
			controls = null;
			label = null;
			balloon = null;
		}
	}

	private void A(List<FileLinkButton> A)
	{
		GH a = default(GH);
		GH CS_0024_003C_003E8__locals8 = new GH(a);
		CS_0024_003C_003E8__locals8.A = this;
		CS_0024_003C_003E8__locals8.A = A;
		new Task([SpecialName] () =>
		{
			foreach (FileLinkButton item in CS_0024_003C_003E8__locals8.A)
			{
				bool flag = false;
				string toolTip = CS_0024_003C_003E8__locals8.A.ToolTip1.GetToolTip(item);
				string a2;
				Image image;
				if (!CS_0024_003C_003E8__locals8.A.m_A.ContainsKey(toolTip))
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
					Uri uri = new Uri(toolTip);
					image = CS_0024_003C_003E8__locals8.A.Icons.Images[10];
					try
					{
						if (uri.HostNameType == UriHostNameType.Dns)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
								{
									HttpWebResponse httpWebResponse = (HttpWebResponse)WebRequest.Create(VH.A(212142) + uri.Host + VH.A(212157)).GetResponse();
									Stream responseStream = httpWebResponse.GetResponseStream();
									image = new Bitmap(Image.FromStream(responseStream), 16, 16);
									httpWebResponse.Close();
									responseStream.Close();
									httpWebResponse = null;
									goto end_IL_008a;
								}
								}
								continue;
								end_IL_008a:
								break;
							}
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						flag = true;
						ProjectData.ClearProjectError();
					}
					a2 = toolTip;
					try
					{
						a2 = Regex.Match(new WebClient().DownloadString(uri), VH.A(212182)).Groups[1].ToString();
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						flag = true;
						ProjectData.ClearProjectError();
					}
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
						FH value = new FH
						{
							A = image,
							A = a2
						};
						CS_0024_003C_003E8__locals8.A.m_A.Add(toolTip, value);
					}
				}
				else
				{
					FH fH = CS_0024_003C_003E8__locals8.A.m_A[toolTip];
					image = fH.A;
					a2 = fH.A;
				}
				item.Icon = image;
				item.Text = a2;
				image = null;
			}
		}).Start();
	}

	private ListViewItem A(Range A, CustomXMLPart B, string C, string D, int E, bool F)
	{
		int imageIndex;
		System.Drawing.Font font;
		if (F)
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
			imageIndex = this.m_H;
			font = this.B();
		}
		else
		{
			imageIndex = this.m_G;
			font = this.A();
		}
		ListViewItem listViewItem = new ListViewItem(C, imageIndex);
		listViewItem.Font = font;
		ListViewItem.ListViewSubItemCollection subItems = listViewItem.SubItems;
		if (D.Length > 0)
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
			subItems.Add(this.B(Conversions.ToString(this.A(D))));
		}
		else
		{
			subItems.Add("");
		}
		subItems.Add(E.ToString());
		subItems.Add(A.Worksheet.Name);
		subItems.Add(A.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)));
		if (Operators.ConditionalCompareObjectEqual(A.Cells.CountLarge, 1, TextCompare: false))
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
				subItems.Add(A.Value2.ToString());
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				subItems.Add("");
				ProjectData.ClearProjectError();
			}
		}
		else
		{
			subItems.Add("");
		}
		subItems = null;
		listViewItem.Tag = new DH
		{
			A = A,
			A = B
		};
		return listViewItem;
	}

	private void E(object A, EventArgs B)
	{
		if (!(this.m_A.Selection is Range))
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
			Range a = (Range)this.m_A.Selection;
			CustomXMLPart b;
			try
			{
				b = this.A(a);
				ListView listView = lvDiscussions;
				listView.BeginUpdate();
				listView.Items.Add(this.A(a, b, this.m_A.UserName, "", 0, F: false));
				C();
				listView.Items[checked(listView.Items.Count - 1)].Selected = true;
				listView.EndUpdate();
				listView = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			b = null;
			a = null;
			return;
		}
	}

	private CustomXMLPart A(Range A)
	{
		Microsoft.Office.Interop.Excel.Workbook workbook = (Microsoft.Office.Interop.Excel.Workbook)A.Worksheet.Parent;
		CustomXMLPart result = null;
		if (this.A(workbook))
		{
			result = workbook.CustomXMLParts.Add(this.A(this.A(A).Name), RuntimeHelpers.GetObjectValue(Missing.Value));
		}
		else
		{
			System.Windows.Forms.MessageBox.Show(VH.A(202661), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}
		workbook = null;
		return result;
	}

	private string A(string A)
	{
		XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
		StringBuilder stringBuilder = new StringBuilder();
		xmlWriterSettings.OmitXmlDeclaration = true;
		xmlWriterSettings.Indent = false;
		xmlWriterSettings.NewLineHandling = NewLineHandling.None;
		xmlWriterSettings.CloseOutput = true;
		XmlWriter xmlWriter = XmlWriter.Create(stringBuilder, xmlWriterSettings);
		string result;
		try
		{
			xmlWriter.WriteStartDocument();
			xmlWriter.WriteStartElement(this.m_C, this.m_L);
			xmlWriter.WriteStartAttribute(this.m_D);
			xmlWriter.WriteValue(this.m_B.ToString());
			xmlWriter.WriteEndAttribute();
			xmlWriter.WriteStartElement(this.m_E);
			xmlWriter.WriteValue(A);
			xmlWriter.WriteEndElement();
			xmlWriter.WriteStartElement(this.m_F);
			xmlWriter.WriteValue(this.m_A.UserName);
			xmlWriter.WriteEndElement();
			xmlWriter.WriteStartElement(this.m_G);
			xmlWriter.WriteValue(DateTime.UtcNow.ToString());
			xmlWriter.WriteEndElement();
			xmlWriter.WriteStartElement(this.m_H);
			xmlWriter.WriteEndElement();
			xmlWriter.WriteStartElement(this.m_J);
			xmlWriter.WriteEndElement();
			xmlWriter.WriteEndElement();
			xmlWriter.WriteEndDocument();
			xmlWriter.Flush();
			_ = null;
			result = stringBuilder.ToString();
		}
		finally
		{
			if (xmlWriter != null)
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
					((IDisposable)xmlWriter).Dispose();
					break;
				}
			}
		}
		xmlWriterSettings = null;
		stringBuilder = null;
		return result;
	}

	private Name A(Range A)
	{
		return ((Microsoft.Office.Interop.Excel.Workbook)A.Worksheet.Parent).Names.Add(this.A(), A, false, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	private string A()
	{
		return this.m_B + Guid.NewGuid().ToString().Replace(VH.A(13778), "");
	}

	private void F(object A, EventArgs B)
	{
		if (System.Windows.Forms.MessageBox.Show(VH.A(202817), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			ListViewItem listViewItem;
			try
			{
				listViewItem = lvDiscussions.SelectedItems[0];
				CustomXMLPart a = ((DH)listViewItem.Tag).A;
				try
				{
					this.m_A.ActiveWorkbook.Names.Item(C(a), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				a.Delete();
				a = null;
				lvDiscussions.BeginUpdate();
				lvDiscussions.Items.Remove(listViewItem);
				if (lvDiscussions.Items.Count == 0)
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
					this.B();
					D();
					F();
				}
				else
				{
					C();
				}
				lvDiscussions.EndUpdate();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
			}
			listViewItem = null;
			return;
		}
	}

	private void G(object A, EventArgs B)
	{
		Worksheet worksheet = this.A();
		if (System.Windows.Forms.MessageBox.Show(VH.A(202914), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			if (worksheet != null)
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
				if (Workbooks.IsShared(this.m_A.ActiveWorkbook, true, (System.Windows.Window)null))
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
					break;
				}
			}
			List<string> list;
			try
			{
				list = new List<string>();
				enumerator = lvDiscussions.Items.GetEnumerator();
				try
				{
					while (enumerator.MoveNext())
					{
						CustomXMLPart a = ((DH)((ListViewItem)enumerator.Current).Tag).A;
						list.Add(C(a));
						a.Delete();
						a = null;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							goto end_IL_00d0;
						}
						continue;
						end_IL_00d0:
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
				lvDiscussions.BeginUpdate();
				lvDiscussions.Items.Clear();
				this.B();
				lvDiscussions.EndUpdate();
				D();
				foreach (string item in list)
				{
					try
					{
						this.m_A.ActiveWorkbook.Names.Item(item, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Delete();
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				F();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
			}
			list = null;
			return;
		}
	}

	private void F()
	{
		Worksheet worksheet = A();
		if (worksheet == null)
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
			this.m_A.DisplayAlerts = false;
			try
			{
				worksheet.Visible = XlSheetVisibility.xlSheetHidden;
				worksheet.Delete();
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
			this.m_A.DisplayAlerts = true;
			worksheet = null;
			return;
		}
	}

	private XmlDocument A(Microsoft.Office.Interop.Excel.Workbook A, string B)
	{
		XmlDocument xmlDocument = new XmlDocument();
		try
		{
			xmlDocument.LoadXml(this.A(A, B).XML);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			xmlDocument = null;
			ProjectData.ClearProjectError();
		}
		return xmlDocument;
	}

	private CustomXMLPart A(Microsoft.Office.Interop.Excel.Workbook A, string B)
	{
		CustomXMLPart result = null;
		try
		{
			result = A.CustomXMLParts.SelectByID(B);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private string A(CustomXMLPart A)
	{
		return VH.A(197945) + B(A) + VH.A(2826);
	}

	private string B(CustomXMLPart A)
	{
		return A.NamespaceManager.LookupPrefix(A.NamespaceURI);
	}

	private bool A(Microsoft.Office.Interop.Excel.Workbook A)
	{
		if (A.Path.Length > 0)
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
			if (Path.GetExtension(A.Name).Length == 5)
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
		if (A.Path.Length == 0)
		{
			return true;
		}
		return false;
	}

	private void H(object A, EventArgs B)
	{
		Panel panel = (Panel)((PictureBox)A).Parent;
		RichTextBox richTextBox = panel.Controls.OfType<Panel>().ElementAt(0).Controls.OfType<RichTextBox>().ElementAt(0);
		if (System.Windows.Forms.MessageBox.Show(VH.A(115862), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			CustomXMLPart a;
			try
			{
				a = this.A().A;
				string text = this.A(a);
				a.SelectSingleNode(text + this.m_K + VH.A(203046) + this.B(a) + VH.A(203051) + Conversions.ToString(richTextBox.Tag) + VH.A(38059)).Delete();
				flpMessages.Controls.Remove(panel);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
				ProjectData.ClearProjectError();
			}
			a = null;
			richTextBox = null;
			panel = null;
			return;
		}
	}

	private void A(RichTextBox A, Color B)
	{
		A.BackColor = B;
		((Panel)A.Parent).BackColor = B;
		_ = null;
	}

	private void A(object A, ContentsResizedEventArgs B)
	{
		((RichTextBox)A).Height = checked(B.NewRectangle.Height + 3);
	}

	private void A(RichTextBox A)
	{
		CustomXMLPart a = this.A().A;
		string text = this.A(a);
		try
		{
			a.SelectSingleNode(text + this.m_K + VH.A(203046) + B(a) + VH.A(203051) + Conversions.ToString(A.Tag) + VH.A(38059)).FirstChild.Text = A.Rtf;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		a = null;
		A.ClearUndo();
	}

	private void A(RichTextBox A, ref KeyEventArgs B)
	{
		if (A.SelectionFont == null)
		{
			return;
		}
		checked
		{
			System.Drawing.FontStyle style2 = default(System.Drawing.FontStyle);
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
				System.Drawing.Font selectionFont = A.SelectionFont;
				System.Drawing.FontStyle style = selectionFont.Style;
				Keys keyCode = B.KeyCode;
				if (keyCode == Keys.B)
				{
					style2 = ((!selectionFont.Bold) ? (style + 1) : (style - 1));
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
					if (keyCode != Keys.I)
					{
						if (keyCode == Keys.U)
						{
							style2 = ((!selectionFont.Underline) ? (style + 4) : (style - 4));
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
						}
					}
					else
					{
						B.SuppressKeyPress = true;
						if (selectionFont.Italic)
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
							style2 = style - 2;
						}
						else
						{
							style2 = style + 2;
						}
					}
				}
				A.SelectionFont = new System.Drawing.Font(selectionFont.FontFamily, selectionFont.Size, style2);
				selectionFont = null;
				style2 = System.Drawing.FontStyle.Regular;
				style = System.Drawing.FontStyle.Regular;
				return;
			}
		}
	}

	private bool A()
	{
		return (Control.ModifierKeys & (Keys.Modifiers | Keys.KeyCode)) != 0;
	}

	private void I(object A, EventArgs B)
	{
		SuspendLayout();
		Panel panel = pnlRichTextBox;
		panel.Top = btnMessage.Top;
		panel.Left = btnMessage.Left;
		panel.Width = btnMessage.Width;
		panel.Visible = true;
		_ = null;
		btnMessage.Visible = false;
		rtbComment.Focus();
		ResumeLayout();
	}

	private void A(object A, KeyEventArgs B)
	{
		if (!this.A())
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
			Keys keyCode = B.KeyCode;
			if (keyCode <= Keys.I)
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
				if (keyCode != Keys.B)
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
					if (keyCode != Keys.I)
					{
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
			else
			{
				if (keyCode == Keys.S)
				{
					B.Handled = true;
					this.A(rtbComment.Rtf);
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
				if (keyCode != Keys.U)
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
			B.Handled = true;
			this.A(rtbComment, ref B);
			return;
		}
	}

	private void B(object A, KeyEventArgs B)
	{
		bool flag = false;
		if (B.KeyCode != Keys.Escape)
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
			if (rtbComment.TextLength > 0)
			{
				flag = System.Windows.Forms.MessageBox.Show(VH.A(203062), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel;
			}
			if (flag)
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
				btnMessage.Visible = true;
				rtbComment.Clear();
				pnlRichTextBox.Visible = false;
				return;
			}
		}
	}

	private void A(string A)
	{
		ListViewItem listViewItem = lvDiscussions.SelectedItems[0];
		DH dH = (DH)listViewItem.Tag;
		CustomXMLPart a = dH.A;
		string text = this.A(a);
		try
		{
			CustomXMLNode customXMLNode = a.SelectSingleNode(text + this.m_J);
			customXMLNode.AppendChildNode(this.m_K, this.m_L);
			this.A(customXMLNode.LastChild, A);
			this.A(customXMLNode.LastChild, MessageType.Text);
			CustomXMLNode lastChild = customXMLNode.LastChild;
			customXMLNode = null;
			this.A(a);
			CustomXMLNode a2 = lastChild;
			List<FileLinkButton> B = null;
			this.A(a2, ref B);
			E();
			this.A(lastChild);
			lastChild = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		rtbComment.Clear();
		btnMessage.Visible = true;
		pnlRichTextBox.Visible = false;
		dH.A = a;
		listViewItem.Tag = dH;
		listViewItem = null;
		a = null;
		C(VH.A(203155));
	}

	private void J(object A, EventArgs B)
	{
		Microsoft.Office.Core.FileDialog fileDialog = ((_Application)this.m_A).get_FileDialog(MsoFileDialogType.msoFileDialogFilePicker);
		fileDialog.Title = VH.A(203178);
		fileDialog.Filters.Clear();
		fileDialog.Show();
		fileDialog.AllowMultiSelect = true;
		FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
		if (selectedItems.Count > 0)
		{
			List<string> list = new List<string>();
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = selectedItems.GetEnumerator();
				while (enumerator.MoveNext())
				{
					string item = Conversions.ToString(enumerator.Current);
					list.Add(item);
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
			this.A(list);
			list = null;
		}
		_ = null;
		btnFiles.Parent.Focus();
	}

	private void A(object A, System.Windows.Forms.DragEventArgs B)
	{
		if (!B.Data.GetDataPresent(System.Windows.Forms.DataFormats.FileDrop))
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
			string[] source = (string[])B.Data.GetData(System.Windows.Forms.DataFormats.FileDrop);
			this.A(source.ToList());
			return;
		}
	}

	private void B(object A, System.Windows.Forms.DragEventArgs B)
	{
		if (B.Data.GetDataPresent(System.Windows.Forms.DataFormats.FileDrop))
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
					B.Effect = System.Windows.Forms.DragDropEffects.Copy;
					return;
				}
			}
		}
		B.Effect = System.Windows.Forms.DragDropEffects.None;
	}

	private void A(List<string> A)
	{
		_ = this.m_A.ActiveWorkbook.Path + Conversions.ToString(Path.DirectorySeparatorChar);
		ListViewItem listViewItem = lvDiscussions.SelectedItems[0];
		DH dH = (DH)listViewItem.Tag;
		CustomXMLPart a = dH.A;
		string text = this.A(a);
		bool flag = chkEmbed.Checked;
		CustomXMLNode customXMLNode = null;
		if (flag)
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
			if (Workbooks.IsShared(this.m_A.ActiveWorkbook, true, (System.Windows.Window)null))
			{
				goto IL_0382;
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
			using List<string>.Enumerator enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				string current = enumerator.Current;
				string text2 = ((!flag) ? current : Path.GetFileName(current));
				if (a.SelectSingleNode(text + this.m_K + VH.A(202192) + JH.B(text2) + VH.A(43340)) != null)
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
				CustomXMLNode customXMLNode2 = a.SelectSingleNode(text + this.m_J);
				customXMLNode2.AppendChildNode(this.m_K, this.m_L);
				this.A(customXMLNode2.LastChild, text2);
				this.A(customXMLNode2.LastChild, MessageType.File);
				customXMLNode = customXMLNode2.LastChild;
				customXMLNode2 = null;
				if (flag)
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
					object instance = this.B().OLEObjects(RuntimeHelpers.GetObjectValue(Missing.Value));
					string memberName = VH.A(60813);
					object[] obj = new object[7] { current, false, true, 0, 0, 100, 100 };
					object[] array = obj;
					string[] argumentNames = new string[7]
					{
						VH.A(203215),
						VH.A(203232),
						VH.A(203241),
						VH.A(56582),
						VH.A(57409),
						VH.A(109766),
						VH.A(109434)
					};
					bool[] obj2 = new bool[7] { true, false, false, false, false, false, false };
					bool[] array2 = obj2;
					object obj3 = NewLateBinding.LateGet(instance, null, memberName, obj, argumentNames, null, obj2);
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
						current = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
					}
					OLEObject obj4 = (OLEObject)obj3;
					obj4.AutoLoad = false;
					obj4.Name = text2;
				}
				CustomXMLNode a2 = customXMLNode;
				List<FileLinkButton> B = null;
				this.A(a2, ref B);
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_02e6;
				}
				continue;
				end_IL_02e6:
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
			clsReporting.LogException(ex2);
			ProjectData.ClearProjectError();
		}
		E();
		if (customXMLNode != null)
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
			this.A(a);
			this.A(customXMLNode);
			customXMLNode = null;
		}
		dH.A = a;
		listViewItem.Tag = dH;
		C(VH.A(203268));
		goto IL_0382;
		IL_0382:
		listViewItem = null;
		a = null;
	}

	private void K(object A, EventArgs B)
	{
		global::A.K.Settings.DiscussEmbedFiles = chkEmbed.Checked;
	}

	private void L(object A, EventArgs B)
	{
		if (System.Windows.Forms.Clipboard.ContainsText(System.Windows.Forms.TextDataFormat.Text))
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
					string text = System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.Text);
					if (text.StartsWith(VH.A(203291)))
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								break;
							default:
								this.B(text);
								return;
							}
						}
					}
					System.Windows.Forms.MessageBox.Show(VH.A(203300), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					return;
				}
				}
			}
		}
		System.Windows.Forms.MessageBox.Show(VH.A(93593), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
	}

	private void B(string A)
	{
		ListViewItem listViewItem = lvDiscussions.SelectedItems[0];
		DH dH = (DH)listViewItem.Tag;
		CustomXMLPart a = dH.A;
		string text = this.A(a);
		List<FileLinkButton> B = new List<FileLinkButton>();
		if (a.SelectSingleNode(text + this.m_K + VH.A(202192) + JH.B(A) + VH.A(43340)) == null)
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
			try
			{
				CustomXMLNode customXMLNode = a.SelectSingleNode(text + this.m_J);
				customXMLNode.AppendChildNode(this.m_K, this.m_L);
				this.A(customXMLNode.LastChild, A);
				this.A(customXMLNode.LastChild, MessageType.Link);
				CustomXMLNode lastChild = customXMLNode.LastChild;
				customXMLNode = null;
				this.A(a);
				this.A(lastChild, ref B);
				E();
				this.A(lastChild);
				this.A(B);
				B = null;
				lastChild = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			dH.A = a;
			listViewItem.Tag = dH;
		}
		listViewItem = null;
		a = null;
		C(VH.A(203454));
	}

	private void M(object A, EventArgs B)
	{
		Image image = null;
		string text = "";
		bool flag = false;
		CustomXMLPart a = ((DH)lvDiscussions.SelectedItems[0].Tag).A;
		string text2 = this.A(a);
		if (System.Windows.Forms.Clipboard.ContainsImage())
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
			image = System.Windows.Forms.Clipboard.GetImage();
			text = global::A.I.A.FileSystem.GetTempFileName();
			image.Save(text, ImageFormat.Png);
			flag = true;
		}
		else
		{
			Microsoft.Office.Core.FileDialog fileDialog = ((_Application)this.m_A).get_FileDialog(MsoFileDialogType.msoFileDialogFilePicker);
			fileDialog.Title = VH.A(203471);
			fileDialog.Filters.Clear();
			fileDialog.Filters.Add(VH.A(203512), VH.A(203523), RuntimeHelpers.GetObjectValue(Missing.Value));
			fileDialog.Show();
			fileDialog.AllowMultiSelect = false;
			FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
			if (selectedItems.Count > 0)
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
				text = selectedItems.Item(0);
				image = new Bitmap(text);
			}
			selectedItems = null;
			_ = null;
		}
		if (image == null)
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
			string text3 = Guid.NewGuid().ToString();
			try
			{
				CustomXMLNode customXMLNode = a.SelectSingleNode(text2 + this.m_J);
				customXMLNode.AppendChildNode(this.m_K, this.m_L, MsoCustomXMLNodeType.msoCustomXMLNodeElement, text3);
				this.A(customXMLNode.LastChild, MessageType.ScreenShot);
				CustomXMLNode lastChild = customXMLNode.LastChild;
				customXMLNode = null;
				this.A(a);
				G(text, text3);
				CustomXMLNode a2 = lastChild;
				List<FileLinkButton> B2 = null;
				this.A(a2, ref B2);
				E();
				this.A(lastChild);
				lastChild = null;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
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
				try
				{
					File.Delete(text);
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					ProjectData.ClearProjectError();
				}
			}
			C(VH.A(203622));
			return;
		}
	}

	private void G(string A, string B)
	{
		Worksheet worksheet = this.B();
		if (!Workbooks.IsShared((Microsoft.Office.Interop.Excel.Workbook)worksheet.Parent, true, (System.Windows.Window)null))
		{
			worksheet.Shapes.AddPicture2(A, MsoTriState.msoFalse, MsoTriState.msoTrue, 0f, 0f, -1f, -1f, MsoPictureCompress.msoPictureCompressTrue).Name = B;
			_ = null;
		}
		worksheet = null;
	}

	private void N(object A, EventArgs B)
	{
		RichTextBox richTextBox = this.A();
		if (System.Windows.Forms.MessageBox.Show(VH.A(115862), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK)
		{
			CustomXMLPart a;
			try
			{
				a = this.A().A;
				string text = this.A(a);
				a.SelectSingleNode(text + this.m_K + VH.A(203046) + this.B(a) + VH.A(203051) + Conversions.ToString(richTextBox.Tag) + VH.A(38059)).Delete();
				this.A((Balloon)richTextBox.Parent);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
				ProjectData.ClearProjectError();
			}
			a = null;
			richTextBox = null;
		}
	}

	private void O(object A, EventArgs B)
	{
		FileLinkButton fileLinkButton = this.A();
		string toolTip = ToolTip1.GetToolTip(fileLinkButton);
		if (!this.A(toolTip))
		{
			DialogResult dialogResult = System.Windows.Forms.MessageBox.Show(VH.A(203653), VH.A(40448), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
			if (dialogResult == DialogResult.Cancel)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (dialogResult == DialogResult.Yes)
			{
				try
				{
					File.Delete(toolTip);
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					if (File.Exists(toolTip))
					{
						System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
					}
					ProjectData.ClearProjectError();
				}
			}
		}
		else
		{
			if (System.Windows.Forms.MessageBox.Show(VH.A(203809), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.Cancel)
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
				break;
			}
			this.A(toolTip).Delete();
		}
		CustomXMLPart a;
		try
		{
			a = this.A().A;
			a.SelectSingleNode(this.A(a) + this.m_K + VH.A(202192) + JH.B(toolTip) + VH.A(43340)).Delete();
			this.A(fileLinkButton);
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			System.Windows.Forms.MessageBox.Show(ex4.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
			ProjectData.ClearProjectError();
		}
		a = null;
		fileLinkButton = null;
	}

	private void P(object A, EventArgs B)
	{
		FileLinkButton fileLinkButton = this.B();
		string toolTip = ToolTip1.GetToolTip(fileLinkButton);
		if (System.Windows.Forms.MessageBox.Show(VH.A(203894), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			CustomXMLPart a;
			try
			{
				a = this.A().A;
				a.SelectSingleNode(this.A(a) + this.m_K + VH.A(202192) + JH.B(toolTip) + VH.A(43340)).Delete();
				this.A(fileLinkButton);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
				ProjectData.ClearProjectError();
			}
			a = null;
			fileLinkButton = null;
			return;
		}
	}

	private void A(object A, CancelEventArgs B)
	{
		Balloon balloon = (Balloon)this.A().Parent;
		ToolStripMenuItem toolStripMenuItem = btnCommentDelete;
		int enabled;
		if (balloon.AuthorIsMe)
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
			enabled = ((balloon == flpMessages.Controls.OfType<Balloon>().Last()) ? 1 : 0);
		}
		else
		{
			enabled = 0;
		}
		toolStripMenuItem.Enabled = (byte)enabled != 0;
		balloon = null;
	}

	private RichTextBox A()
	{
		return (RichTextBox)cmsComment.SourceControl;
	}

	private void B(object A, CancelEventArgs B)
	{
		bool flag = this.A(this.B());
		btnFileShow.Enabled = !flag;
		btnFileEmbed.Enabled = !flag;
	}

	private void Q(object A, EventArgs B)
	{
		string text = this.B();
		if (!this.A(text))
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
			try
			{
				Process.Start(text);
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				if (!clsFile.FileExists(text, 3000))
				{
					System.Windows.Forms.MessageBox.Show(VH.A(203979), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}
				else
				{
					System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
					clsReporting.LogException(ex2);
				}
				ProjectData.ClearProjectError();
			}
		}
		else
		{
			try
			{
				this.A(text).Activate();
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				clsReporting.LogException(ex4);
				ProjectData.ClearProjectError();
			}
		}
		C(VH.A(204165));
	}

	private void R(object A, EventArgs B)
	{
	}

	private void S(object A, EventArgs B)
	{
		clsFile.OpenExplorerToFile(this.B());
		C(VH.A(204184));
	}

	private string B()
	{
		return ToolTip1.GetToolTip(A());
	}

	private FileLinkButton A()
	{
		return (FileLinkButton)cmsFile.SourceControl;
	}

	private OLEObject A(string A)
	{
		OLEObject result;
		try
		{
			object instance = this.A().OLEObjects(RuntimeHelpers.GetObjectValue(Missing.Value));
			string memberName = VH.A(140662);
			object[] obj = new object[1] { A };
			object[] array = obj;
			bool[] obj2 = new bool[1] { true };
			bool[] array2 = obj2;
			object obj3 = NewLateBinding.LateGet(instance, null, memberName, obj, null, null, obj2);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				A = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
			}
			result = (OLEObject)obj3;
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

	private bool A(string A)
	{
		return !Path.IsPathRooted(A);
	}

	private void T(object A, EventArgs B)
	{
		clsUtilities.GoToUrl(ToolTip1.GetToolTip(this.B()));
		C(VH.A(204223));
	}

	private FileLinkButton B()
	{
		return (FileLinkButton)cmsLink.SourceControl;
	}

	private void H()
	{
		try
		{
			if (lvDiscussions.SelectedItems[0].Font.Bold)
			{
				this.m_A = new System.Timers.Timer(this.m_A);
				this.m_A.Elapsed += [SpecialName] (object A, ElapsedEventArgs B) =>
				{
					CustomXMLPart a = this.A().A;
					CustomXMLNode customXMLNode = this.A(a);
					customXMLNode.AppendChildNode(this.m_I, this.m_L);
					this.A(customXMLNode.LastChild, this.m_A.UserName);
					customXMLNode = null;
					a = null;
					ListViewItem listViewItem = lvDiscussions.SelectedItems[0];
					listViewItem.Font = this.A();
					listViewItem.ImageIndex = this.m_G;
					_ = null;
				};
				this.m_A.AutoReset = false;
				this.m_A.Start();
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void A(CustomXMLPart A)
	{
		CustomXMLNode customXMLNode = this.A(A);
		while (customXMLNode.HasChildNodes())
		{
			customXMLNode.LastChild.Delete();
		}
		customXMLNode = null;
	}

	private CustomXMLNode A(CustomXMLPart A)
	{
		return A.SelectSingleNode(this.A(A) + this.m_H);
	}

	private System.Drawing.Font A()
	{
		return new System.Drawing.Font(Font, System.Drawing.FontStyle.Regular);
	}

	private System.Drawing.Font B()
	{
		return new System.Drawing.Font(Font, System.Drawing.FontStyle.Bold);
	}

	private Padding A(bool A)
	{
		return (!A) ? new Padding(9, 6, 15, 6) : new Padding(15, 6, 9, 6);
	}

	private System.Drawing.Point A(bool A)
	{
		return (!A) ? new System.Drawing.Point(9, 6) : new System.Drawing.Point(15, 6);
	}

	private FileLinkButton A(int A, bool B)
	{
		FileLinkButton fileLinkButton = new FileLinkButton();
		fileLinkButton.Location = this.A(B);
		fileLinkButton.Width = A;
		fileLinkButton.Height = 25;
		fileLinkButton.Margin = this.A(B);
		fileLinkButton.BackColor = Color.White;
		fileLinkButton.ForeColor = Color.FromKnownColor(KnownColor.ControlText);
		_ = null;
		return fileLinkButton;
	}

	private PictureBox A(int A, bool B)
	{
		PictureBox pictureBox;
		PictureBox result = (pictureBox = new PictureBox());
		if (B)
		{
			pictureBox.Location = new System.Drawing.Point(15, 9);
			pictureBox.Margin = new Padding(15, 9, 9, 9);
			pictureBox.MouseDown += this.A;
			pictureBox.MouseMove += this.B;
			pictureBox.MouseUp += C;
			pictureBox.Paint += this.A;
		}
		else
		{
			pictureBox.Location = new System.Drawing.Point(9, 9);
			pictureBox.Margin = new Padding(9, 9, 15, 9);
		}
		pictureBox.Width = A;
		pictureBox.BackColor = Color.White;
		pictureBox.SizeMode = PictureBoxSizeMode.Normal;
		pictureBox.Cursor = Cursors.Cross;
		pictureBox.ContextMenuStrip = cmsPicture;
		pictureBox.DoubleClick += V;
		pictureBox = null;
		return result;
	}

	private void A(PictureBox A, string B)
	{
		try
		{
			Shape shape = this.A(B);
			if (shape != null)
			{
				shape.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);
				if (System.Windows.Forms.Clipboard.ContainsImage())
				{
					A.Image = System.Windows.Forms.Clipboard.GetImage();
					System.Windows.Forms.Clipboard.Clear();
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

	private void U(object A, EventArgs B)
	{
		System.Windows.Forms.Clipboard.SetImage(this.A().Image);
	}

	private void V(object A, EventArgs B)
	{
		this.A((PictureBox)A);
	}

	private void W(object A, EventArgs B)
	{
		this.A(this.A());
	}

	private void A(PictureBox A)
	{
		string a = Conversions.ToString(A.Tag);
		Shape shape = this.A(a);
		if (shape == null)
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
			shape.Copy();
			Microsoft.Office.Interop.Excel.Workbook activeWorkbook = this.m_A.ActiveWorkbook;
			activeWorkbook.DisplayDrawingObjects = XlDisplayDrawingObjects.xlDisplayShapes;
			if (activeWorkbook.ActiveSheet is Worksheet)
			{
				((Worksheet)activeWorkbook.ActiveSheet).Paste(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			}
			else
			{
				Worksheet obj = (Worksheet)activeWorkbook.Worksheets[1];
				obj.Paste(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
				obj.Activate();
				_ = null;
			}
			shape = null;
			activeWorkbook = null;
			return;
		}
	}

	private void X(object A, EventArgs B)
	{
		if (System.Windows.Forms.MessageBox.Show(VH.A(204246), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			PictureBox pictureBox = this.A();
			string text = Conversions.ToString(pictureBox.Tag);
			Shape shape = this.A(text);
			if (shape != null)
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
				shape.Delete();
				shape = null;
			}
			CustomXMLPart a = this.A().A;
			a.SelectSingleNode(this.A(a) + this.m_K + VH.A(204333) + text + VH.A(38059)).Delete();
			a = null;
			this.A((Balloon)pictureBox.Parent);
			pictureBox = null;
			return;
		}
	}

	private PictureBox A()
	{
		return (PictureBox)cmsPicture.SourceControl;
	}

	private Shape A(string A)
	{
		Shape result;
		try
		{
			result = this.A().Shapes.Item(A);
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

	private void A(object A, MouseEventArgs B)
	{
		if (B.Button != MouseButtons.Left)
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
			this.m_A = B.Location;
			this.m_B = this.m_A;
			this.m_A = true;
			return;
		}
	}

	private void B(object A, MouseEventArgs B)
	{
		if (B.Button != MouseButtons.Left)
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
			this.m_B = B.Location;
			if (!this.m_A)
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
				((PictureBox)A).Invalidate();
				return;
			}
		}
	}

	private void C(object A, MouseEventArgs B)
	{
		if (B.Button != MouseButtons.Left)
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
			if (!this.m_A)
			{
				return;
			}
			this.m_A = false;
			PictureBox pictureBox = (PictureBox)A;
			if (pictureBox.Image.Width <= ((Panel)pictureBox.Parent).Width)
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
				Rectangle item = this.A();
				if (item.Width > 0)
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
					if (item.Height > 0)
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
						this.m_A.Add(item);
					}
				}
				pictureBox.Invalidate();
				this.B(pictureBox);
			}
			else
			{
				pictureBox.Invalidate();
				System.Windows.Forms.MessageBox.Show(VH.A(204352), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
			pictureBox = null;
			return;
		}
	}

	private void A(object A, PaintEventArgs B)
	{
		bool flag = chkPen.Checked;
		Graphics graphics = B.Graphics;
		if (this.m_A.Any())
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
			if (flag)
			{
				graphics.DrawRectangles(this.m_A, this.m_A.ToArray());
			}
			else
			{
				graphics.FillRectangles(this.m_A, this.m_A.ToArray());
			}
		}
		if (this.m_A)
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
			if (flag)
			{
				graphics.DrawRectangle(this.m_A, this.A());
			}
			else
			{
				graphics.FillRectangle(this.m_A, this.A());
			}
		}
		graphics = null;
	}

	private Rectangle A()
	{
		return checked(new Rectangle(Math.Min(this.m_A.X, this.m_B.X), Math.Min(this.m_A.Y, this.m_B.Y), Math.Abs(this.m_A.X - this.m_B.X), Math.Abs(this.m_A.Y - this.m_B.Y)));
	}

	private void B(PictureBox A)
	{
		int num = A.Image.Width;
		int num2 = A.Image.Height;
		Bitmap bitmap = new Bitmap(num, num2);
		A.DrawToBitmap(bitmap, new Rectangle(0, 0, num, num2));
		string tempFileName = global::A.I.A.FileSystem.GetTempFileName();
		bitmap.Save(tempFileName, ImageFormat.Png);
		A.Paint -= this.A;
		A.Image = (Image)bitmap.Clone();
		A.Paint += this.A;
		bitmap.Dispose();
		bitmap = null;
		string text = Conversions.ToString(A.Tag);
		A = null;
		this.m_A = new List<Rectangle>();
		Shape shape = this.A(text);
		if (shape != null)
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
			shape.Delete();
			shape = null;
		}
		G(tempFileName, text);
		try
		{
			File.Delete(tempFileName);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		C(VH.A(204514));
	}

	private void I()
	{
		//IL_0271: Unknown result type (might be due to invalid IL or missing references)
		//IL_0277: Expected O, but got Unknown
		this.m_E = global::A.K.Settings.DiscussPenColor;
		Color e = this.m_E;
		if (e == this.m_G)
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
			chkPenRed.Checked = true;
		}
		else if (e == this.m_H)
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
			chkPenGreen.Checked = true;
		}
		else if (e == this.m_I)
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
			chkPenBlue.Checked = true;
		}
		else if (e == this.m_J)
		{
			chkPenYellow.Checked = true;
		}
		this.m_I = global::A.K.Settings.DiscussPenThickness;
		int i = this.m_I;
		if (i == this.m_J)
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
			chkPenThin.Checked = true;
		}
		else if (i == this.m_K)
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
			chkPenMedium.Checked = true;
		}
		else if (i == this.m_L)
		{
			chkPenThick.Checked = true;
		}
		this.m_F = global::A.K.Settings.DiscussHighlighterColor;
		Color f = this.m_F;
		if (f == this.m_K)
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
			chkHighlighterYellow.Checked = true;
		}
		else if (f == this.m_L)
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
			chkHighlighterOrange.Checked = true;
		}
		else if (f == this.m_M)
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
			chkHighlighterPink.Checked = true;
		}
		else if (f == this.m_N)
		{
			chkHighlighterBlue.Checked = true;
		}
		else if (f == this.m_O)
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
			chkHighlighterGreen.Checked = true;
		}
		if (global::A.K.Settings.DiscussUsePen)
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
			chkPen.Checked = true;
			chkPenMenu.Checked = true;
		}
		else
		{
			chkHighlighter.Checked = true;
			chkHighlighterMenu.Checked = true;
		}
		L();
		M();
		clsDisplay val = new clsDisplay();
		checked
		{
			int num = (int)Math.Round(14.0 * val.X);
			int num2 = (int)Math.Round(14.0 * val.Y);
			val = null;
			chkPenRed.Image = clsColors.ColorSquare(this.m_G, num, num2);
			chkPenGreen.Image = clsColors.ColorSquare(this.m_H, num, num2);
			chkPenBlue.Image = clsColors.ColorSquare(this.m_I, num, num2);
			chkPenYellow.Image = clsColors.ColorSquare(this.m_J, num, num2);
			chkHighlighterYellow.Image = clsColors.ColorSquare(this.m_K, num, num2);
			chkHighlighterOrange.Image = clsColors.ColorSquare(this.m_L, num, num2);
			chkHighlighterPink.Image = clsColors.ColorSquare(this.m_M, num, num2);
			chkHighlighterBlue.Image = clsColors.ColorSquare(this.m_N, num, num2);
			chkHighlighterGreen.Image = clsColors.ColorSquare(this.m_O, num, num2);
		}
	}

	private void J()
	{
		chkHighlighter.Checked = false;
		chkHighlighterMenu.Checked = false;
		chkPen.Checked = true;
		chkPenMenu.Checked = true;
		global::A.K.Settings.DiscussUsePen = true;
	}

	private void K()
	{
		chkPen.Checked = false;
		chkPenMenu.Checked = false;
		chkHighlighter.Checked = true;
		chkHighlighterMenu.Checked = true;
		global::A.K.Settings.DiscussUsePen = false;
	}

	private void Y(object A, EventArgs B)
	{
		J();
	}

	private void Z(object A, EventArgs B)
	{
		K();
	}

	private void AB(object A, EventArgs B)
	{
		J();
		ContextMenuStrip contextMenuStrip = cmsPens;
		if (this.m_B)
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
			contextMenuStrip.Hide();
		}
		else
		{
			contextMenuStrip.Visible = true;
			contextMenuStrip.Show(chkPenMenu, new System.Drawing.Point(contextMenuStrip.Width, -1), ToolStripDropDownDirection.AboveLeft);
			contextMenuStrip.Focus();
		}
		contextMenuStrip = null;
	}

	private void BB(object A, EventArgs B)
	{
		K();
		ContextMenuStrip contextMenuStrip = cmsHighlighters;
		if (this.m_C)
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
			contextMenuStrip.Hide();
		}
		else
		{
			contextMenuStrip.Visible = true;
			contextMenuStrip.Show(chkHighlighterMenu, new System.Drawing.Point(contextMenuStrip.Width, -1), ToolStripDropDownDirection.AboveLeft);
			contextMenuStrip.Focus();
		}
		contextMenuStrip = null;
	}

	private void CB(object A, EventArgs B)
	{
		ToolStripMenuItem[] array = new ToolStripMenuItem[4] { chkPenRed, chkPenGreen, chkPenBlue, chkPenYellow };
		for (int i = 0; i < array.Length; i = checked(i + 1))
		{
			array[i].Checked = false;
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
			((ToolStripMenuItem)A).Checked = true;
			if (chkPenRed.Checked)
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
				this.m_E = this.m_G;
			}
			else if (chkPenGreen.Checked)
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
				this.m_E = this.m_H;
			}
			else if (chkPenBlue.Checked)
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
				this.m_E = this.m_I;
			}
			else if (chkPenYellow.Checked)
			{
				this.m_E = this.m_J;
			}
			L();
			global::A.K.Settings.DiscussPenColor = this.m_E;
			return;
		}
	}

	private void DB(object A, EventArgs B)
	{
		ToolStripMenuItem[] array = new ToolStripMenuItem[3] { chkPenThin, chkPenThick, chkPenMedium };
		for (int i = 0; i < array.Length; i = checked(i + 1))
		{
			array[i].Checked = false;
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
			((ToolStripMenuItem)A).Checked = true;
			if (chkPenThin.Checked)
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
				this.m_I = this.m_J;
			}
			else if (chkPenMedium.Checked)
			{
				this.m_I = this.m_K;
			}
			else if (chkPenThick.Checked)
			{
				this.m_I = this.m_L;
			}
			L();
			global::A.K.Settings.DiscussPenThickness = this.m_I;
			return;
		}
	}

	private void EB(object A, EventArgs B)
	{
		ToolStripMenuItem[] array = new ToolStripMenuItem[5] { chkHighlighterYellow, chkHighlighterOrange, chkHighlighterPink, chkHighlighterBlue, chkHighlighterGreen };
		for (int i = 0; i < array.Length; i = checked(i + 1))
		{
			array[i].Checked = false;
		}
		((ToolStripMenuItem)A).Checked = true;
		if (chkHighlighterYellow.Checked)
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
			this.m_F = this.m_K;
		}
		else if (chkHighlighterOrange.Checked)
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
			this.m_F = this.m_L;
		}
		else if (chkHighlighterPink.Checked)
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
			this.m_F = this.m_M;
		}
		else if (chkHighlighterBlue.Checked)
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
			this.m_F = this.m_N;
		}
		else if (chkHighlighterGreen.Checked)
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
			this.m_F = this.m_O;
		}
		M();
		global::A.K.Settings.DiscussHighlighterColor = this.m_F;
	}

	private void A(object A, ToolStripDropDownClosingEventArgs B)
	{
		this.m_B = false;
	}

	private void B(object A, ToolStripDropDownClosingEventArgs B)
	{
		this.m_C = false;
	}

	private void C(object A, CancelEventArgs B)
	{
		this.m_B = true;
	}

	private void D(object A, CancelEventArgs B)
	{
		this.m_C = true;
	}

	private void L()
	{
		this.m_A = new Pen(this.m_E, this.m_I);
	}

	private void M()
	{
		this.m_A = new SolidBrush(Color.FromArgb(50, this.m_F));
	}

	private void FB(object A, EventArgs B)
	{
		bool flag = chkControls.Checked;
		pnlControls.Visible = flag;
		if (flag)
		{
			chkControls.Image = global::A.J.TreeNodeExpanded;
		}
		else
		{
			chkControls.Image = global::A.J.TreeNodeCollapsed;
		}
		N();
	}

	private void GB(object A, EventArgs B)
	{
		if (!rtbComment.Focused)
		{
			flpMessages.Focus();
		}
	}

	private void HB(object A, EventArgs B)
	{
		N();
	}

	private void N()
	{
		int num = flpMessages.Width;
		flpMessages.SuspendLayout();
		checked
		{
			if (B())
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
				num -= SystemInformation.VerticalScrollBarWidth;
			}
			IEnumerator<Balloon> enumerator = default(IEnumerator<Balloon>);
			try
			{
				enumerator = flpMessages.Controls.OfType<Balloon>().GetEnumerator();
				while (enumerator.MoveNext())
				{
					Balloon current = enumerator.Current;
					int num2 = current.Margin.Left + current.Margin.Right;
					Control control = current.Controls[0];
					control.Width = num - control.Margin.Left - control.Margin.Right - num2;
					control = null;
					current = null;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						goto end_IL_00e6;
					}
					continue;
					end_IL_00e6:
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
			IEnumerator<Label> enumerator2 = default(IEnumerator<Label>);
			try
			{
				enumerator2 = flpMessages.Controls.OfType<Label>().GetEnumerator();
				while (enumerator2.MoveNext())
				{
					enumerator2.Current.Width = num;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_0140;
					}
					continue;
					end_IL_0140:
					break;
				}
			}
			finally
			{
				if (enumerator2 != null)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						enumerator2.Dispose();
						break;
					}
				}
			}
			flpMessages.ResumeLayout();
		}
	}

	private bool B()
	{
		return flpMessages.PreferredSize.Height > flpMessages.Height;
	}

	private void A(Balloon A)
	{
		try
		{
			flpMessages.Controls.Remove((Label)A.Tag);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		flpMessages.Controls.Remove(A);
	}

	private void A(FileLinkButton A)
	{
		this.A((Balloon)A.Parent);
	}

	private Worksheet A()
	{
		Worksheet worksheet = null;
		try
		{
			worksheet = (Worksheet)this.m_A.ActiveWorkbook.Worksheets[clsDiscuss.HIDDEN_SHEET_NAME];
			worksheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return worksheet;
	}

	private Worksheet B()
	{
		Worksheet worksheet = A();
		if (worksheet == null)
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
			if (this.m_A.ActiveWindow.SelectedSheets.Count > 1)
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
				NewLateBinding.LateCall(this.m_A.ActiveSheet, null, VH.A(51162), new object[0], null, null, null, IgnoreReturn: true);
			}
			worksheet = (Worksheet)this.m_A.ActiveWorkbook.Worksheets.Add(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
			worksheet.Name = clsDiscuss.HIDDEN_SHEET_NAME;
			worksheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
		}
		return worksheet;
	}

	private void A(CustomXMLNode A, string B)
	{
		A.AppendChildNode("", this.m_L, MsoCustomXMLNodeType.msoCustomXMLNodeCData, B);
	}

	private void A(CustomXMLNode A, MessageType B)
	{
		A.AppendChildNode(VH.A(204553), this.m_L, MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, this.m_A.UserName);
		A.AppendChildNode(VH.A(204562), this.m_L, MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, DateTime.UtcNow.ToString());
		A.AppendChildNode(VH.A(204571), this.m_L, MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, Guid.NewGuid().ToString());
		string name = VH.A(144960);
		string l = this.m_L;
		int num = (int)B;
		A.AppendChildNode(name, l, MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, num.ToString());
		_ = null;
	}

	private MessageType A(CustomXMLNode A)
	{
		return (MessageType)Conversions.ToInteger(A.Attributes[this.m_F].Text);
	}

	private string A(CustomXMLNode A)
	{
		return B(Conversions.ToString(this.A(A.Attributes[this.m_D].Text)));
	}

	private string C(CustomXMLPart A)
	{
		return A.SelectSingleNode(this.A(A) + this.m_E).Text;
	}

	private DH A()
	{
		return (DH)lvDiscussions.SelectedItems[0].Tag;
	}

	private void A(CustomXMLNode A)
	{
		lvDiscussions.BeginUpdate();
		ListViewItem listViewItem = lvDiscussions.SelectedItems[0];
		listViewItem.SubItems[colMessages.Index].Text = Conversions.ToString(flpMessages.Controls.OfType<Balloon>().Count());
		listViewItem.SubItems[colDate.Index].Text = this.A(A);
		_ = null;
		C();
		lvDiscussions.EndUpdate();
	}

	private DateTime A(string A)
	{
		return DateTime.Parse(A).ToLocalTime();
	}

	private string B(string A)
	{
		return Base.FormatTime(A);
	}

	private void A(ListView A, ColumnClickEventArgs B, ref int C)
	{
		try
		{
			ListView listView = A;
			listView.Groups.Clear();
			if (B.Column != C)
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
				C = B.Column;
				listView.Sorting = SortOrder.Ascending;
			}
			else
			{
				listView.Sorting = ((listView.Sorting != SortOrder.Ascending) ? SortOrder.Ascending : SortOrder.Descending);
			}
			listView.BeginUpdate();
			listView.Sort();
			listView.ListViewItemSorter = new HH(B.Column, listView.Sorting);
			clsApis.ConfigureListView(A);
			listView.EndUpdate();
			listView = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void C(string A)
	{
		clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)15, A);
	}

	[SpecialName]
	[CompilerGenerated]
	private void A(object A, ElapsedEventArgs B)
	{
		CustomXMLPart a = this.A().A;
		CustomXMLNode customXMLNode = this.A(a);
		customXMLNode.AppendChildNode(this.m_I, this.m_L);
		this.A(customXMLNode.LastChild, this.m_A.UserName);
		customXMLNode = null;
		a = null;
		ListViewItem listViewItem = lvDiscussions.SelectedItems[0];
		listViewItem.Font = this.A();
		listViewItem.ImageIndex = this.m_G;
		_ = null;
	}
}
