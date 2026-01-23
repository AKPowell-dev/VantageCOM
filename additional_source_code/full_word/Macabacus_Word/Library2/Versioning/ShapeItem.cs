using System.Runtime.CompilerServices;
using System.Windows.Media;
using A;
using MacabacusMacros.Libraries.Versioning;
using Macabacus_Word.Shapes;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.Library2.Versioning;

public sealed class ShapeItem : ContentItem
{
	private object m_A;

	internal object Shape
	{
		get
		{
			return this.m_A;
		}
		set
		{
			this.m_A = RuntimeHelpers.GetObjectValue(value);
			A();
		}
	}

	internal ShapeItem(object A, ContentInfo B, ManifestInfo C)
		: base(B, C)
	{
		//IL_0001: Unknown result type (might be due to invalid IL or missing references)
		//IL_0002: Unknown result type (might be due to invalid IL or missing references)
		Shape = RuntimeHelpers.GetObjectValue(A);
	}

	private void A()
	{
		string source = default(string);
		InlineShape inlineShape;
		if (Shape is InlineShape)
		{
			inlineShape = (InlineShape)Shape;
			if (Images.A(inlineShape))
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
				source = XC.A(7434);
			}
			else
			{
				if (inlineShape.Type != WdInlineShapeType.wdInlineShapeChart)
				{
					if (inlineShape.HasChart != MsoTriState.msoTrue)
					{
						source = XC.A(7767);
						goto IL_0085;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				source = XC.A(7666);
			}
			goto IL_0085;
		}
		Microsoft.Office.Interop.Word.Shape shape;
		if (Shape is Microsoft.Office.Interop.Word.Shape)
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
			shape = (Microsoft.Office.Interop.Word.Shape)Shape;
			if (Images.A(shape))
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
				source = XC.A(7434);
			}
			else
			{
				if (shape.Type != MsoShapeType.msoChart)
				{
					if (shape.HasChart != MsoTriState.msoTrue)
					{
						if (shape.Type == MsoShapeType.msoTable)
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
							source = XC.A(8037);
						}
						else
						{
							source = XC.A(7767);
						}
						goto IL_0136;
					}
					while (true)
					{
						switch (2)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				source = XC.A(7666);
			}
			goto IL_0136;
		}
		goto IL_0138;
		IL_0138:
		((ContentItem)this).IconData = Geometry.Parse(source);
		return;
		IL_0136:
		shape = null;
		goto IL_0138;
		IL_0085:
		inlineShape = null;
		goto IL_0138;
	}
}
