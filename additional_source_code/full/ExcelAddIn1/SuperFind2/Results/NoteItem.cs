using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Xml;
using A;
using ExcelAddIn1.Comments;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class NoteItem : ExploreItem
{
	private bool m_A;

	[CompilerGenerated]
	private Comment m_A;

	public override bool IsSelected
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(21693));
			D();
		}
	}

	internal Comment Note
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

	public NoteItem(WorksheetItem wsi, Range rng)
		: base(wsi, Constants.ColorPalette.Amber.Clone(), Props.Icons.GeoNote, 21)
	{
		base.Range = rng;
		Note = rng.Comment;
		Refresh();
	}

	public override void Refresh()
	{
		D();
		base.Tooltip = Note.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		base.PreviewImage = null;
	}

	public override void Delete()
	{
		if (MessageBox.Show(VH.A(117201), VH.A(40448), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) != DialogResult.OK)
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
			Note.Delete();
			base.Parent.A(this);
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
			if (!Note.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).ToLower().Contains(strQuery))
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
				isHighlighted = ((Operators.CompareString(strQuery, VH.A(117286), TextCompare: false) == 0) ? 1 : 0);
				goto IL_0090;
			}
		}
		isHighlighted = 1;
		goto IL_0090;
		IL_0090:
		((BaseItem)this).IsHighlighted = (byte)isHighlighted != 0;
	}

	internal void A()
	{
		try
		{
			XmlNode xmlNode = KH.A.SettingsXml.DocumentElement.SelectSingleNode(VH.A(117297));
			Fix.Note(Note, Conversions.ToBoolean(xmlNode.Attributes[VH.A(117320)].Value), Conversions.ToBoolean(xmlNode.Attributes[VH.A(117333)].Value));
			xmlNode = null;
			base.PreviewImage = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	internal void B()
	{
		string text = Author.RemoveFromText(Note);
		Range obj = (Range)Note.Parent;
		Note.Delete();
		Comment comment = obj.AddComment(text);
		comment.Shape.TextFrame.Characters(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Font.Bold = false;
		((BaseItem)this).Label = A(comment);
		base.Tooltip = comment.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		Note = comment;
		comment = null;
		base.PreviewImage = null;
	}

	internal void C()
	{
		string text = Author.RemoveFromText(Note);
		Range obj = (Range)Note.Parent;
		obj.Comment.Delete();
		string memberName = VH.A(117350);
		object[] obj2 = new object[1] { text };
		object[] array = obj2;
		bool[] obj3 = new bool[1] { true };
		bool[] array2 = obj3;
		NewLateBinding.LateCall(obj, null, memberName, obj2, null, null, obj3, IgnoreReturn: true);
		if (!array2[0])
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
			text = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
			return;
		}
	}

	private void D()
	{
		((BaseItem)this).Label = A(Note);
	}

	private string A(Comment A)
	{
		return A.Text(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value)).Replace(VH.A(41382), VH.A(41385));
	}
}
