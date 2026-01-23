using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Macabacus_Word.Values;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace Macabacus_Word.TextOps.Redaction.Values;

public sealed class RangeValue
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<IShape, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal int A(IShape A)
		{
			return A.RangeStart();
		}
	}

	private SelectionValue m_A;

	[CompilerGenerated]
	private Range m_A;

	[CompilerGenerated]
	private bool m_A;

	[CompilerGenerated]
	private bool m_B;

	[CompilerGenerated]
	private bool m_C;

	[CompilerGenerated]
	private bool m_D;

	[CompilerGenerated]
	private List<IShape> m_A;

	[CompilerGenerated]
	private List<IShape> m_B;

	[CompilerGenerated]
	private List<GroupedShapesValue> m_A;

	[CompilerGenerated]
	private List<GroupedShapesValue> m_B;

	public Range Range
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
	}

	public bool HasNonFloatingShapes
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
	}

	public bool HasFloatingShapes
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
	}

	public bool HasGroupedNonFloatingShapes
	{
		[CompilerGenerated]
		get
		{
			return this.m_C;
		}
	}

	public bool HasGroupedFloatingShapes
	{
		[CompilerGenerated]
		get
		{
			return this.m_D;
		}
	}

	public List<IShape> NonFloatingShapesList
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
	}

	public List<IShape> FloatingShapesList
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
	}

	public List<GroupedShapesValue> GroupedFloatingShapesList
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
	}

	public List<GroupedShapesValue> GroupedNonFloatingShapesList
	{
		[CompilerGenerated]
		get
		{
			return this.m_B;
		}
	}

	public RangeValue(Range rng, SelectionValue selectionValue)
	{
		this.m_A = selectionValue;
		this.m_A = rng;
		this.m_A = new List<IShape>();
		this.m_B = new List<IShape>();
		this.m_A = new List<GroupedShapesValue>();
		this.m_B = new List<GroupedShapesValue>();
		A();
		this.m_A = (List<IShape>)A();
		B();
		this.m_B = FloatingShapesList.Count > 0;
		this.m_D = GroupedFloatingShapesList.Count > 0;
		this.m_A = NonFloatingShapesList.Count > 0;
		this.m_C = GroupedNonFloatingShapesList.Count > 0;
	}

	private void A()
	{
		IEnumerator enumerator = Range.InlineShapes.GetEnumerator();
		try
		{
			while (enumerator.MoveNext())
			{
				InlineShape shp = (InlineShape)enumerator.Current;
				NonFloatingShapesList.Add(new InlineShapeValue(shp));
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
				break;
			}
		}
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
		if (this.m_A.ShapeRange.Count > 0)
		{
			C();
		}
		else
		{
			D();
		}
	}

	private void B()
	{
		if (this.m_A.ShapeRange.Count > 0)
		{
			E();
		}
		else
		{
			F();
		}
	}

	private void C()
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = this.m_A.ShapeRange.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
				if (shape.WrapFormat.Type != WdWrapType.wdWrapInline)
				{
					continue;
				}
				if (shape.Type != MsoShapeType.msoGroup)
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
					NonFloatingShapesList.Add(new ShapeValue(shape));
				}
				else
				{
					if (A())
					{
						continue;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					GroupedNonFloatingShapesList.Add(new GroupedShapesValue(shape));
				}
			}
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (5)
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

	private void D()
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = Range.ShapeRange.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
				if (shape.WrapFormat.Type != WdWrapType.wdWrapInline)
				{
					continue;
				}
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
				if (shape.Type != MsoShapeType.msoGroup)
				{
					NonFloatingShapesList.Add(new ShapeValue(shape));
				}
				else
				{
					GroupedNonFloatingShapesList.Add(new GroupedShapesValue(shape));
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (3)
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

	private void E()
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = this.m_A.ShapeRange.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
				if (shape.WrapFormat.Type == WdWrapType.wdWrapInline)
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (shape.Type != MsoShapeType.msoGroup)
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
					FloatingShapesList.Add(new ShapeValue(shape));
				}
				else
				{
					if (A())
					{
						continue;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					GroupedFloatingShapesList.Add(new GroupedShapesValue(shape));
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return;
				}
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

	private void F()
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = Range.ShapeRange.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Microsoft.Office.Interop.Word.Shape shape = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
				if (shape.WrapFormat.Type == WdWrapType.wdWrapInline)
				{
					continue;
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (shape.Type != MsoShapeType.msoGroup)
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
					FloatingShapesList.Add(new ShapeValue(shape));
				}
				else
				{
					GroupedFloatingShapesList.Add(new GroupedShapesValue(shape));
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					return;
				}
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (6)
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

	private bool A()
	{
		if (this.m_A.ChildShapeRange.Count > 0)
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
			Microsoft.Office.Interop.Word.ShapeRange childShapeRange = this.m_A.ChildShapeRange;
			object Index = 1;
			IEnumerator enumerator = default(IEnumerator);
			if (childShapeRange[ref Index].Type != MsoShapeType.msoGroup)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						try
						{
							enumerator = this.m_A.ChildShapeRange.GetEnumerator();
							while (enumerator.MoveNext())
							{
								Microsoft.Office.Interop.Word.Shape shp = (Microsoft.Office.Interop.Word.Shape)enumerator.Current;
								FloatingShapesList.Add(new ShapeValue(shp, isInsideGroup: true));
							}
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									goto end_IL_009a;
								}
								continue;
								end_IL_009a:
								break;
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
										break;
									default:
										(enumerator as IDisposable).Dispose();
										goto end_IL_00ae;
									}
									continue;
									end_IL_00ae:
									break;
								}
							}
						}
						return true;
					}
				}
			}
		}
		return false;
	}

	private object A()
	{
		List<IShape> nonFloatingShapesList = NonFloatingShapesList;
		Func<IShape, int> keySelector;
		if (_Closure_0024__.A == null)
		{
			keySelector = (_Closure_0024__.A = [SpecialName] (IShape A) => A.RangeStart());
		}
		else
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
			keySelector = _Closure_0024__.A;
		}
		return nonFloatingShapesList.OrderBy(keySelector).ToList();
	}
}
