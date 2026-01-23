using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using A;

namespace ExcelAddIn1.Audit.TraceDialogs;

public sealed class RangeObservableCollection<T> : ObservableCollection<T>
{
	private bool A;

	public RangeObservableCollection()
	{
		A = false;
	}

	public RangeObservableCollection(IEnumerable<T> list)
	{
		A = false;
		AddRange(list);
	}

	protected override void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
	{
		if (A)
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
			base.OnCollectionChanged(e);
			return;
		}
	}

	public void AddRange(IEnumerable<T> list)
	{
		if (list == null)
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
					throw new ArgumentNullException(VH.A(49981));
				}
			}
		}
		A = true;
		using (IEnumerator<T> enumerator = list.GetEnumerator())
		{
			while (enumerator.MoveNext())
			{
				T current = enumerator.Current;
				Add(current);
			}
			while (true)
			{
				switch (1)
				{
				case 0:
					break;
				default:
					goto end_IL_0052;
				}
				continue;
				end_IL_0052:
				break;
			}
		}
		A = false;
		OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
	}
}
