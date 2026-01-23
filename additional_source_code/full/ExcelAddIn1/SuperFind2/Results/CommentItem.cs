using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using ExcelAddIn1.SuperFind2.UI;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class CommentItem : ExploreItem
{
	private bool m_A;

	[CompilerGenerated]
	private object m_A;

	private Visibility m_A;

	private Visibility m_B;

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
			C();
		}
	}

	internal object Comment
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = RuntimeHelpers.GetObjectValue(value);
		}
	}

	public Visibility ResolveVisibility
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(115788));
			C();
		}
	}

	public Visibility UnresolveVisibility
	{
		get
		{
			return this.m_B;
		}
		set
		{
			this.m_B = value;
			((BaseItem)this).NotifyPropertyChanged(VH.A(115823));
			C();
		}
	}

	public CommentItem(WorksheetItem wsi, Range rng)
		: base(wsi, Constants.ColorPalette.Lilac.Clone(), Props.Icons.GeoCommentSolid, 20)
	{
		base.Range = rng;
		Comment = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(rng, null, VH.A(103833), new object[0], null, null, null));
		C();
		if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(NewLateBinding.LateGet(Comment, null, VH.A(102647), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null), 0, TextCompare: false))
		{
			base.Tooltip = A(RuntimeHelpers.GetObjectValue(Comment));
		}
		else
		{
			object comment = Comment;
			string memberName = VH.A(102647);
			object instance;
			object[] obj = new object[1] { NewLateBinding.LateGet(instance = NewLateBinding.LateGet(Comment, null, VH.A(102647), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null) };
			object[] array = obj;
			bool[] obj2 = new bool[1] { true };
			bool[] array2 = obj2;
			object obj3 = NewLateBinding.LateGet(comment, null, memberName, obj, null, null, obj2);
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				NewLateBinding.LateSetComplex(instance, null, VH.A(52690), new object[1] { array[0] }, null, null, OptimisticSet: true, RValueBase: true);
			}
			base.Tooltip = A(RuntimeHelpers.GetObjectValue(obj3));
		}
		D();
	}

	public override void Refresh()
	{
		C();
		base.PreviewImage = null;
	}

	public override void Delete()
	{
		if (MessageBox.Show(VH.A(115862), VH.A(40448), MessageBoxButton.OKCancel, MessageBoxImage.Exclamation) != MessageBoxResult.OK)
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
			NewLateBinding.LateCall(Comment, null, VH.A(60691), new object[0], null, null, null, IgnoreReturn: true);
			base.Parent.A(this);
			return;
		}
	}

	public override void Search(string strQuery)
	{
		object instance = NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(Comment, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null), null, VH.A(103864), new object[0], null, null, null);
		string memberName = VH.A(104857);
		object[] obj = new object[1] { strQuery };
		object[] array = obj;
		bool[] obj2 = new bool[1] { true };
		bool[] array2 = obj2;
		object value = NewLateBinding.LateGet(instance, null, memberName, obj, null, null, obj2);
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
			strQuery = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
		}
		int num;
		if (!Conversions.ToBoolean(value))
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
			object instance2 = NewLateBinding.LateGet(NewLateBinding.LateGet(Comment, null, VH.A(96399), new object[0], null, null, null), null, VH.A(103864), new object[0], null, null, null);
			string memberName2 = VH.A(104857);
			object[] obj3 = new object[1] { strQuery };
			array = obj3;
			bool[] obj4 = new bool[1] { true };
			array2 = obj4;
			object value2 = NewLateBinding.LateGet(instance2, null, memberName2, obj3, null, null, obj4);
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
				strQuery = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
			}
			if (!Conversions.ToBoolean(value2))
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
				num = ((Operators.CompareString(strQuery, VH.A(115953), TextCompare: false) == 0) ? 1 : 0);
				goto IL_01aa;
			}
		}
		num = 1;
		goto IL_01aa;
		IL_01aa:
		if (Conversions.ToBoolean((byte)num != 0))
		{
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					((BaseItem)this).IsHighlighted = true;
					return;
				}
			}
		}
		bool isHighlighted = false;
		int num2 = Conversions.ToInteger(NewLateBinding.LateGet(NewLateBinding.LateGet(Comment, null, VH.A(102647), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null));
		for (int i = 1; i <= num2; i = checked(i + 1))
		{
			object obj5 = NewLateBinding.LateGet(Comment, null, VH.A(102647), array = new object[1] { i }, null, null, array2 = new bool[1] { true });
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
				i = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
			}
			object instance3 = obj5;
			object value3 = NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(instance3, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null), null, VH.A(103864), new object[0], null, null, null), null, VH.A(104857), array = new object[1] { strQuery }, null, null, array2 = new bool[1] { true });
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
				strQuery = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
			}
			int num3;
			if (!Conversions.ToBoolean(value3))
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
				object instance4 = NewLateBinding.LateGet(NewLateBinding.LateGet(instance3, null, VH.A(96399), new object[0], null, null, null), null, VH.A(103864), new object[0], null, null, null);
				string memberName3 = VH.A(104857);
				object[] obj6 = new object[1] { strQuery };
				array = obj6;
				bool[] obj7 = new bool[1] { true };
				array2 = obj7;
				object value4 = NewLateBinding.LateGet(instance4, null, memberName3, obj6, null, null, obj7);
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
					strQuery = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
				}
				num3 = (Conversions.ToBoolean(value4) ? 1 : 0);
			}
			else
			{
				num3 = 1;
			}
			if (Conversions.ToBoolean((byte)num3 != 0))
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
				isHighlighted = true;
			}
			instance3 = null;
		}
		((BaseItem)this).IsHighlighted = isHighlighted;
	}

	internal void A()
	{
		try
		{
			NewLateBinding.LateSetComplex(Comment, null, VH.A(102617), new object[1] { true }, null, null, OptimisticSet: false, RValueBase: true);
			D();
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
		try
		{
			NewLateBinding.LateSetComplex(Comment, null, VH.A(102617), new object[1] { false }, null, null, OptimisticSet: false, RValueBase: true);
			D();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private void C()
	{
		((BaseItem)this).Label = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateGet(NewLateBinding.LateGet(Comment, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null), VH.A(115970)), NewLateBinding.LateGet(NewLateBinding.LateGet(Comment, null, VH.A(102647), new object[0], null, null, null), null, VH.A(52690), new object[0], null, null, null).ToString()), VH.A(115975)), base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value))));
	}

	private string A(object A)
	{
		string text = Conversions.ToString(NewLateBinding.LateGet(A, null, VH.A(96399), new object[0], null, null, null));
		text = Conversions.ToString(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(NewLateBinding.LateGet(A, null, VH.A(96399), new object[0], null, null, null), null, VH.A(65312), new object[2]
		{
			VH.A(7803),
			VH.A(41385)
		}, null, null, null), null, VH.A(65312), new object[2]
		{
			'\r',
			VH.A(41385)
		}, null, null, null), null, VH.A(65312), new object[2]
		{
			'\n',
			VH.A(41385)
		}, null, null, null));
		if (text.Length > 100)
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
			text = Strings.Left(text, 97) + VH.A(116000);
		}
		return Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateGet(NewLateBinding.LateGet(A, null, VH.A(102634), new object[0], null, null, null), null, VH.A(19019), new object[0], null, null, null), VH.A(7803)), NewLateBinding.LateGet(A, null, VH.A(102662), new object[0], null, null, null).ToString()), VH.A(7803)), text));
	}

	private void D()
	{
		if (Conversions.ToBoolean(NewLateBinding.LateGet(Comment, null, VH.A(102617), new object[0], null, null, null)))
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
					((BaseItem)this).Icon = Props.Icons.GeoCommentOutline;
					ResolveVisibility = Visibility.Collapsed;
					UnresolveVisibility = Visibility.Visible;
					return;
				}
			}
		}
		((BaseItem)this).Icon = Props.Icons.GeoCommentSolid;
		ResolveVisibility = Visibility.Visible;
		UnresolveVisibility = Visibility.Collapsed;
	}
}
