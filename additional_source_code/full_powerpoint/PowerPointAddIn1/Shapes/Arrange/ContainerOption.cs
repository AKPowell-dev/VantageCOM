using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Windows.Media.Imaging;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddIn1.Template;

namespace PowerPointAddIn1.Shapes.Arrange;

public sealed class ContainerOption : INotifyPropertyChanged
{
	[CompilerGenerated]
	private PropertyChangedEventHandler m_A;

	[CompilerGenerated]
	private string m_A;

	[CompilerGenerated]
	private BitmapSource m_A;

	[CompilerGenerated]
	private RectangleF m_A;

	private bool m_A;

	public string Label
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

	public BitmapSource Image
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

	internal RectangleF Rectangle
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

	public bool IsChecked
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = value;
			A(AH.A(12198));
		}
	}

	public event PropertyChangedEventHandler PropertyChanged
	{
		[CompilerGenerated]
		add
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Combine(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
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
				return;
			}
		}
		[CompilerGenerated]
		remove
		{
			PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
			PropertyChangedEventHandler propertyChangedEventHandler2;
			do
			{
				propertyChangedEventHandler2 = propertyChangedEventHandler;
				PropertyChangedEventHandler value2 = (PropertyChangedEventHandler)Delegate.Remove(propertyChangedEventHandler2, value);
				propertyChangedEventHandler = Interlocked.CompareExchange(ref this.m_A, value2, propertyChangedEventHandler2);
			}
			while ((object)propertyChangedEventHandler != propertyChangedEventHandler2);
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
				return;
			}
		}
	}

	public ContainerOption(string strLabel)
	{
		Label = strLabel;
		SizeF sizeF = A();
		Rectangle = new RectangleF(0f, 0f, sizeF.Width, sizeF.Height);
	}

	public ContainerOption(string strLabel, ContainerCanvas canvas, Settings.Margins margins)
	{
		Label = strLabel;
		SizeF sizeF = A();
		Rectangle = new RectangleF(margins.Left, margins.Top, sizeF.Width - margins.Left - margins.Right, sizeF.Height - margins.Top - margins.Bottom);
		A(canvas);
	}

	public ContainerOption(string strLabel, ContainerCanvas canvas, RectangleF rect)
	{
		Label = strLabel;
		Rectangle = rect;
		A(canvas);
	}

	public ContainerOption(string strLabel, ContainerCanvas canvas, Shape shp)
	{
		Label = strLabel;
		Shape shape = shp;
		Rectangle = new RectangleF(shape.Left, shape.Top, shape.Width, shape.Height);
		shape = null;
		A(canvas);
	}

	private void A(string A)
	{
		PropertyChangedEventHandler propertyChangedEventHandler = this.m_A;
		if (propertyChangedEventHandler == null)
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
			propertyChangedEventHandler(this, new PropertyChangedEventArgs(A));
			return;
		}
	}

	private SizeF A()
	{
		Master slideMaster = NG.A.Application.ActivePresentation.SlideMaster;
		return new SizeF(slideMaster.Width, slideMaster.Height);
	}

	private void A(ContainerCanvas A)
	{
		Bitmap bitmap = new Bitmap(A.Width, A.Height);
		Rectangle destRect = new Rectangle(0, 0, A.Width, A.Height);
		Color color = ColorTranslator.FromHtml(AH.A(12183));
		Graphics graphics = Graphics.FromImage(bitmap);
		try
		{
			Pen pen = new Pen(color, 0.5f);
			try
			{
				pen.DashPattern = new float[2] { 2f, 1f };
				SolidBrush solidBrush = new SolidBrush(color);
				try
				{
					RectangleF rectangle = Rectangle;
					graphics.DrawRectangle(pen, rectangle.Left * A.Scale, rectangle.Top * A.Scale, rectangle.Width * A.Scale, rectangle.Height * A.Scale);
					graphics.DrawImage(bitmap, destRect, 0, 0, A.Width, A.Height, GraphicsUnit.Pixel);
				}
				finally
				{
					if (solidBrush != null)
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
							((IDisposable)solidBrush).Dispose();
							break;
						}
					}
				}
			}
			finally
			{
				if (pen != null)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						((IDisposable)pen).Dispose();
						break;
					}
				}
			}
		}
		finally
		{
			if (graphics != null)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						continue;
					}
					((IDisposable)graphics).Dispose();
					break;
				}
			}
		}
		destRect = default(Rectangle);
		Image = Forms.GetImageSource(bitmap);
	}
}
