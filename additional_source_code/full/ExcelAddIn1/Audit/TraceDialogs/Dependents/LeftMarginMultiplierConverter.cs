using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ExcelAddIn1.Audit.TraceDialogs.Dependents;

public sealed class LeftMarginMultiplierConverter : IValueConverter
{
	public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
	{
		return new Thickness(checked(((BaseItem)((TreeViewItem)value).DataContext).Level * 20), 0.0, 0.0, 0.0);
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
