using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace PowerPointAddIn1.Explorer;

public sealed class LeftMarginMultiplierConverter : IValueConverter
{
	public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
	{
		TreeViewItem treeViewItem = (TreeViewItem)value;
		int num;
		if (treeViewItem.DataContext is PresentationItem)
		{
			num = 0;
		}
		else if (treeViewItem.DataContext is SlideItem)
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
			num = 1;
		}
		else
		{
			num = 2;
		}
		treeViewItem = null;
		return new Thickness(checked(num * 20), 0.0, 0.0, 0.0);
	}

	object IValueConverter.Convert(object value, Type targetType, object parameter, CultureInfo culture)
	{
		//ILSpy generated this explicit interface implementation from .override directive in Convert
		return this.Convert(value, targetType, parameter, culture);
	}

	public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
	{
		object result = default(object);
		return result;
	}

	object IValueConverter.ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
	{
		//ILSpy generated this explicit interface implementation from .override directive in ConvertBack
		return this.ConvertBack(value, targetType, parameter, culture);
	}
}
