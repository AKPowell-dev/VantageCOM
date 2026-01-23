using System;
using System.Drawing;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using A;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.DeckCheck;

namespace PowerPointAddIn1.Utilities;

public sealed class Dialogs
{
	public enum DialogPosition
	{
		Above,
		Below,
		Left,
		Right
	}

	internal static readonly int A = 28;

	private static void A(double A, double B, ref int C, ref int D)
	{
		HwndSource hwndSource = new HwndSource(default(HwndSourceParameters));
		Matrix transformToDevice;
		try
		{
			transformToDevice = hwndSource.CompositionTarget.TransformToDevice;
		}
		finally
		{
			if (hwndSource != null)
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
					((IDisposable)hwndSource).Dispose();
					break;
				}
			}
		}
		checked
		{
			C = (int)Math.Round(transformToDevice.M11 * A);
			D = (int)Math.Round(transformToDevice.M22 * B);
		}
	}

	internal static void B(ref double A, ref double B, int C, int D)
	{
		HwndSource hwndSource = new HwndSource(default(HwndSourceParameters));
		Matrix transformToDevice;
		try
		{
			transformToDevice = hwndSource.CompositionTarget.TransformToDevice;
		}
		finally
		{
			if (hwndSource != null)
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
					((IDisposable)hwndSource).Dispose();
					break;
				}
			}
		}
		A = (double)C / transformToDevice.M11;
		B = (double)D / transformToDevice.M22;
	}

	internal static Rectangle A(Shape A)
	{
		Main.EnsureSlidePaneActive();
		float top = A.Top;
		float left = A.Left;
		float height = A.Height;
		float width = A.Width;
		_ = null;
		DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
		int num = activeWindow.PointsToScreenPixelsX(left);
		int num2 = activeWindow.PointsToScreenPixelsY(top);
		Rectangle result = checked(new Rectangle(num, num2, activeWindow.PointsToScreenPixelsX(left + width) - num, activeWindow.PointsToScreenPixelsY(top + height) - num2));
		activeWindow = null;
		return result;
	}

	internal static void A(Shape A, DialogPosition B, ref double C, ref double D)
	{
		Rectangle rectangle = Dialogs.A(A);
		int num = rectangle.Left;
		int num2 = rectangle.Top;
		checked
		{
			switch (B)
			{
			case DialogPosition.Above:
				num2 -= 10;
				break;
			case DialogPosition.Below:
				num2 += rectangle.Height + 10;
				break;
			case DialogPosition.Left:
				num -= 10;
				break;
			case DialogPosition.Right:
				num += rectangle.Width + 10;
				break;
			}
			Dialogs.B(ref C, ref D, num, num2);
			rectangle = default(Rectangle);
		}
	}

	internal static HwndSource GetHwndSource(Window win)
	{
		return HwndSource.FromHwnd(new WindowInteropHelper(win).Handle);
	}
}
