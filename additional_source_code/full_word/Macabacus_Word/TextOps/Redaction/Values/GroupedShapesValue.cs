using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Macabacus_Word.TextOps.Redaction.Process;
using Macabacus_Word.Values;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.TextOps.Redaction.Values;

public sealed class GroupedShapesValue
{
	[CompilerGenerated]
	private List<IShape> m_A;

	[CompilerGenerated]
	private ShapeValue m_A;

	[CompilerGenerated]
	private bool m_A;

	public List<IShape> FloatingShapesList
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
	}

	public ShapeValue GroupedShape
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
	}

	public bool HasPictureOrChart
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
	}

	public GroupedShapesValue(Shape shp)
	{
		this.m_A = new List<IShape>();
		this.m_A = A(shp);
		this.m_A = new ShapeValue(shp);
	}

	private bool A(Shape A)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GroupItems.GetEnumerator();
			bool result = default(bool);
			while (enumerator.MoveNext())
			{
				Shape shp = (Shape)enumerator.Current;
				if (!RedactUtilities.IsPictureOrChart(shp))
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
					FloatingShapesList.Add(new ShapeValue(shp));
				}
				else
				{
					result = true;
				}
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					continue;
				}
				return result;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}
}
