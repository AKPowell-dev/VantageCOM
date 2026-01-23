using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Explorer;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.SuperFind2.Results;

public sealed class ValueItem : SearchItem
{
	public ValueItem(WorksheetItem wsi, Range rng)
		: base(wsi, rng)
	{
		Refresh();
	}

	public override void Refresh()
	{
		string text = base.Range.get_Address((object)0, (object)0, XlReferenceStyle.xlA1, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
		if (Operators.ConditionalCompareObjectEqual(base.Range.Cells.CountLarge, 1, TextCompare: false))
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
			object objectValue = RuntimeHelpers.GetObjectValue(base.Range.Value2);
			if (Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(objectValue)))
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
				string left = text;
				string left2 = VH.A(116343);
				Type typeFromHandle = typeof(Math);
				string memberName = VH.A(118502);
				object[] obj = new object[2] { objectValue, 4 };
				object[] array = obj;
				bool[] obj2 = new bool[2] { true, false };
				bool[] array2 = obj2;
				object right = NewLateBinding.LateGet(null, typeFromHandle, memberName, obj, null, null, obj2);
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
				text = Conversions.ToString(Operators.ConcatenateObject(left, Operators.ConcatenateObject(left2, right)));
			}
			else
			{
				text = Conversions.ToString(Operators.ConcatenateObject(text, Operators.ConcatenateObject(VH.A(116343), objectValue)));
			}
		}
		((BaseItem)this).Label = text;
	}
}
