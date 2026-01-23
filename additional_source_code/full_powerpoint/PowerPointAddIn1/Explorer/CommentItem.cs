using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Windows.Media;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Explorer;

public sealed class CommentItem : ContentItem
{
	private new bool m_A;

	[CompilerGenerated]
	private new Comment m_A;

	public override bool IsSelected
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(AH.A(62846));
			A();
		}
	}

	public Comment Comment
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

	public CommentItem(SlideItem si, Comment cmt, SolidColorBrush brush)
		: base(si, "", brush, Pane.CachedObjects.GeoComment)
	{
		Comment = cmt;
		((BaseItem)this).Label = A(cmt);
		base.Tooltip = cmt.Text;
		SearchOnInstantiate();
	}

	public override void Refresh()
	{
		A();
		base.PreviewImage = null;
	}

	public override void Delete()
	{
		if (MessageBox.Show(AH.A(113483), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			Comment.Delete();
			base.Parent.RemoveChild(this);
			return;
		}
	}

	public override void Search(string strQuery)
	{
		int isHighlighted;
		if (!((BaseItem)this).Label.ToLower().Contains(strQuery))
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
			if (!Comment.Text.ToLower().Contains(strQuery))
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
				isHighlighted = ((Operators.CompareString(strQuery, AH.A(113574), TextCompare: false) == 0) ? 1 : 0);
				goto IL_006c;
			}
		}
		isHighlighted = 1;
		goto IL_006c;
		IL_006c:
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}

	private void A()
	{
		((BaseItem)this).Label = A(Comment);
	}

	private string A(Comment A)
	{
		return A.Author + AH.A(15084) + A.Text.Replace(AH.A(47334), AH.A(14625));
	}
}
