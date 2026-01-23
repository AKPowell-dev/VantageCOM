using System;
using System.Globalization;
using System.Windows.Data;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Audit.Check.UI;

public sealed class PopupBooleanConverter : IValueConverter
{
	public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
	{
		return !Conversions.ToBoolean(value);
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
