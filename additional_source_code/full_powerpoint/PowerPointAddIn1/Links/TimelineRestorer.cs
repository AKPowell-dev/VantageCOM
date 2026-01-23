using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Links;

public sealed class TimelineRestorer
{
	private sealed class JE : IDisposable
	{
		[Serializable]
		[CompilerGenerated]
		internal sealed class _Closure_0024__
		{
			public static readonly _Closure_0024__ A;

			public static Func<List<List<PE>>> A;

			public static Func<List<int>> A;

			public static Func<int, int> A;

			public static Func<List<List<PE>>> B;

			public static Func<PE, bool> A;

			public static Func<PE, bool> B;

			static _Closure_0024__()
			{
				_Closure_0024__.A = new _Closure_0024__();
			}

			[SpecialName]
			internal List<List<PE>> A()
			{
				return new List<List<PE>>();
			}

			[SpecialName]
			internal List<int> A()
			{
				return new List<int>();
			}

			[SpecialName]
			internal int A(int A)
			{
				return A;
			}

			[SpecialName]
			internal List<List<PE>> B()
			{
				return new List<List<PE>>();
			}

			[SpecialName]
			internal bool A(PE A)
			{
				return A.ShapeId.HasValue;
			}

			[SpecialName]
			internal bool B(PE A)
			{
				return A.ShapeId.HasValue;
			}
		}

		[CompilerGenerated]
		internal sealed class GE
		{
			public int? A;

			public Func<PE, bool> A;

			public GE(GE A)
			{
				if (A == null)
				{
					return;
				}
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
					this.A = A.A;
					return;
				}
			}

			[SpecialName]
			internal bool A(PE A)
			{
				return object.Equals(A.TriggerShapeId, this.A.Value);
			}

			[SpecialName]
			internal bool B(PE A)
			{
				return object.Equals(A.TriggerShapeId, this.A.Value);
			}
		}

		[CompilerGenerated]
		internal sealed class HE
		{
			public int A;

			public Func<PE, bool> A;

			public HE(HE A)
			{
				if (A != null)
				{
					this.A = A.A;
				}
			}

			[SpecialName]
			internal IEnumerable<PE> A(List<PE> A)
			{
				Func<PE, bool> predicate;
				if (this.A != null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					predicate = this.A;
				}
				else
				{
					predicate = (this.A = [SpecialName] (PE pE) => object.Equals(pE.ShapeId, this.A));
				}
				return A.Where(predicate).ToList();
			}

			[SpecialName]
			internal bool A(PE A)
			{
				return object.Equals(A.ShapeId, this.A);
			}

			[SpecialName]
			internal bool A(int A)
			{
				return A == this.A;
			}
		}

		[CompilerGenerated]
		internal sealed class IE
		{
			public Effect A;

			public IE(IE A)
			{
				if (A == null)
				{
					return;
				}
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
					this.A = A.A;
					return;
				}
			}

			[SpecialName]
			internal bool A(PE A)
			{
				return A.B(this.A);
			}
		}

		private Slide m_A;

		private TimelineRestorer m_A;

		private readonly List<int> m_A;

		private List<List<PE>> m_A;

		private Sequence m_A;

		private Sequences m_A;

		private Dictionary<int, Microsoft.Office.Interop.PowerPoint.Shape> m_A;

		private bool m_A;

		internal JE(Slide A, TimelineRestorer B)
		{
			this.m_A = null;
			this.m_A = A;
			this.m_A = B;
			int slideID = A.SlideID;
			Dictionary<int, List<List<PE>>> a = this.m_A.m_A;
			Func<List<List<PE>>> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = [SpecialName] () => new List<List<PE>>());
			}
			else
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				c = _Closure_0024__.A;
			}
			this.m_A = TimelineRestorer.A(a, slideID, c);
			if (this.m_A.Count == 0)
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
				this.m_A = TimelineRestorer.A(this.m_A.m_A, slideID, [SpecialName] () => new List<int>());
				try
				{
					TimeLine timeLine = this.m_A.TimeLine;
					this.m_A = timeLine.MainSequence;
					this.m_A = timeLine.InteractiveSequences;
					return;
				}
				finally
				{
					TimeLine timeLine = null;
				}
			}
		}

		internal void B()
		{
			try
			{
				if (this.m_A.Count == 0)
				{
					while (true)
					{
						switch (1)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							return;
						}
					}
				}
				if (this.m_A.Count == this.m_A.Count)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
				this.m_A = B(this.m_A);
				IEnumerator<IGrouping<int, int>> enumerator = default(IEnumerator<IGrouping<int, int>>);
				try
				{
					List<int> source = TimelineRestorer.A(this.m_A);
					Func<int, int> keySelector;
					if (_Closure_0024__.A == null)
					{
						keySelector = (_Closure_0024__.A = [SpecialName] (int A) => A);
					}
					else
					{
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						keySelector = _Closure_0024__.A;
					}
					enumerator = source.GroupBy(keySelector).GetEnumerator();
					while (enumerator.MoveNext())
					{
						IGrouping<int, int> current = enumerator.Current;
						C(current.Key, current.Count());
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_00da;
						}
						continue;
						end_IL_00da:
						break;
					}
				}
				finally
				{
					if (enumerator != null)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							enumerator.Dispose();
							break;
						}
					}
				}
				D();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			finally
			{
				Dictionary<int, Microsoft.Office.Interop.PowerPoint.Shape> a = this.m_A;
				if (a == null)
				{
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				else
				{
					a.Clear();
				}
				this.m_A = null;
			}
		}

		private void C(int A, int B)
		{
			HE a = default(HE);
			HE CS_0024_003C_003E8__locals10 = new HE(a);
			CS_0024_003C_003E8__locals10.A = A;
			if (CS_0024_003C_003E8__locals10.A == int.MinValue || !this.m_A.ContainsKey(CS_0024_003C_003E8__locals10.A))
			{
				return;
			}
			List<PE> list = null;
			checked
			{
				try
				{
					list = this.m_A.SelectMany([SpecialName] (List<PE> source) =>
					{
						Func<PE, bool> predicate;
						if (CS_0024_003C_003E8__locals10.A != null)
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
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							predicate = CS_0024_003C_003E8__locals10.A;
						}
						else
						{
							predicate = (CS_0024_003C_003E8__locals10.A = [SpecialName] (PE pE2) => object.Equals(pE2.ShapeId, CS_0024_003C_003E8__locals10.A));
						}
						return source.Where(predicate).ToList();
					}).ToList();
					int num = this.m_A.Where([SpecialName] (int num3) => num3 == CS_0024_003C_003E8__locals10.A).Count();
					IE iE = default(IE);
					for (int num2 = num + 1; num2 <= B; num2++)
					{
						iE = new IE(iE);
						try
						{
							iE.A = this.B(CS_0024_003C_003E8__locals10.A, num + 1);
							PE pE = list.FirstOrDefault(iE.A);
							if (pE == null)
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
									return;
								}
							}
							list.Remove(pE);
							if (!pE.A())
							{
								while (true)
								{
									switch (4)
									{
									case 0:
										break;
									default:
										goto end_IL_00f1;
									}
									continue;
									end_IL_00f1:
									break;
								}
							}
							else
							{
								Sequence a2 = this.B(pE);
								pE.A(a2, CS_0024_003C_003E8__locals10.A, this.m_A);
								iE.A.Delete();
							}
						}
						catch (Exception projectError)
						{
							ProjectData.SetProjectError(projectError);
							ProjectData.ClearProjectError();
						}
						finally
						{
							Sequence a2 = null;
							iE.A = null;
							PE pE = null;
						}
					}
				}
				finally
				{
					if (list != null)
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
						list.Clear();
					}
					list = null;
				}
			}
		}

		private static Dictionary<int, Microsoft.Office.Interop.PowerPoint.Shape> B(Slide A)
		{
			Dictionary<int, Microsoft.Office.Interop.PowerPoint.Shape> dictionary = new Dictionary<int, Microsoft.Office.Interop.PowerPoint.Shape>();
			try
			{
				Microsoft.Office.Interop.PowerPoint.Shapes shapes = A.Shapes;
				foreach (Microsoft.Office.Interop.PowerPoint.Shape item in shapes)
				{
					try
					{
						dictionary[item.Id] = item;
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
					finally
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape = null;
					}
				}
				return dictionary;
			}
			finally
			{
				Microsoft.Office.Interop.PowerPoint.Shapes shapes = null;
			}
		}

		private Sequence B(PE A)
		{
			Sequence sequence = null;
			try
			{
				IEnumerator enumerator = default(IEnumerator);
				try
				{
					enumerator = this.m_A.GetEnumerator();
					while (true)
					{
						if (enumerator.MoveNext())
						{
							Sequence sequence2 = (Sequence)enumerator.Current;
							try
							{
								if (sequence2.Count == 0)
								{
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
										break;
									}
									continue;
								}
								Effect a = sequence2[1];
								if (!A.A(a))
								{
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
									continue;
								}
								sequence = sequence2;
							}
							catch (Exception projectError)
							{
								ProjectData.SetProjectError(projectError);
								ProjectData.ClearProjectError();
								continue;
							}
							finally
							{
								Effect a = null;
								sequence2 = null;
							}
							break;
						}
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								goto end_IL_007e;
							}
							continue;
							end_IL_007e:
							break;
						}
						break;
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
				if (sequence == null)
				{
					sequence = this.m_A.Add();
				}
				return sequence;
			}
			finally
			{
				sequence = null;
			}
		}

		private Effect B(int A, int B)
		{
			int num = 0;
			int count = this.m_A.Count;
			checked
			{
				for (int i = 1; i <= count; i++)
				{
					try
					{
						Effect effect = this.m_A[i];
						if (effect.Shape.Id != A)
						{
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
								break;
							}
						}
						else
						{
							num++;
							if (num >= B)
							{
								return effect;
							}
						}
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
					finally
					{
						Effect effect = null;
					}
				}
				return null;
			}
		}

		private void D()
		{
			List<List<PE>> B = TimelineRestorer.A(this.m_A);
			if (B.Count == 0)
			{
				return;
			}
			List<List<PE>>.Enumerator enumerator = default(List<List<PE>>.Enumerator);
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
				try
				{
					Dictionary<int, List<List<PE>>> a = this.m_A.m_A;
					int slideID = this.m_A.SlideID;
					Func<List<List<PE>>> c;
					if (_Closure_0024__.B == null)
					{
						c = (_Closure_0024__.B = [SpecialName] () => new List<List<PE>>());
					}
					else
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
						c = _Closure_0024__.B;
					}
					enumerator = TimelineRestorer.A(a, slideID, c).GetEnumerator();
					while (enumerator.MoveNext())
					{
						List<PE> current = enumerator.Current;
						this.B(current, ref B);
					}
					while (true)
					{
						switch (1)
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
					((IDisposable)enumerator/*cast due to .constrained prefix*/).Dispose();
				}
			}
		}

		private void B(List<PE> A, ref List<List<PE>> B)
		{
			GE a = default(GE);
			GE CS_0024_003C_003E8__locals4 = new GE(a);
			if (A.Count < 2)
			{
				return;
			}
			checked
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					PE pE = A.First();
					CS_0024_003C_003E8__locals4.A = pE.TriggerShapeId;
					if (!CS_0024_003C_003E8__locals4.A.HasValue || !A.All([SpecialName] (PE pE3) => object.Equals(pE3.TriggerShapeId, CS_0024_003C_003E8__locals4.A.Value)))
					{
						return;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						Func<PE, bool> predicate;
						if (_Closure_0024__.A == null)
						{
							predicate = (_Closure_0024__.A = [SpecialName] (PE pE3) => pE3.ShapeId.HasValue);
						}
						else
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								break;
							}
							predicate = _Closure_0024__.A;
						}
						if (!A.All(predicate))
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
							List<PE> list = null;
							int num = 0;
							using (List<List<PE>>.Enumerator enumerator = B.GetEnumerator())
							{
								while (true)
								{
									if (enumerator.MoveNext())
									{
										List<PE> current = enumerator.Current;
										num++;
										try
										{
											if (current.Count != A.Count)
											{
												while (true)
												{
													switch (3)
													{
													case 0:
														break;
													default:
														goto end_IL_00e9;
													}
													continue;
													end_IL_00e9:
													break;
												}
												continue;
											}
											if (!current.All([SpecialName] (PE pE3) => object.Equals(pE3.TriggerShapeId, CS_0024_003C_003E8__locals4.A.Value)))
											{
												while (true)
												{
													switch (1)
													{
													case 0:
														break;
													default:
														goto end_IL_012a;
													}
													continue;
													end_IL_012a:
													break;
												}
												continue;
											}
											List<PE> source = current;
											Func<PE, bool> predicate2;
											if (_Closure_0024__.B == null)
											{
												predicate2 = (_Closure_0024__.B = [SpecialName] (PE pE3) => pE3.ShapeId.HasValue);
											}
											else
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
												predicate2 = _Closure_0024__.B;
											}
											if (!source.All(predicate2))
											{
												while (true)
												{
													switch (3)
													{
													case 0:
														break;
													default:
														goto end_IL_016d;
													}
													continue;
													end_IL_016d:
													break;
												}
												continue;
											}
											list = current;
										}
										finally
										{
											current = null;
										}
										break;
									}
									while (true)
									{
										switch (5)
										{
										case 0:
											break;
										default:
											goto end_IL_0190;
										}
										continue;
										end_IL_0190:
										break;
									}
									break;
								}
							}
							if (list == null)
							{
								return;
							}
							while (true)
							{
								switch (5)
								{
								case 0:
									continue;
								}
								int count = A.Count;
								int num2 = 1;
								while (true)
								{
									if (num2 <= count)
									{
										PE pE2 = A[num2 - 1];
										int? num3 = null;
										int num4 = num2;
										int count2 = list.Count;
										int num5 = num4;
										while (true)
										{
											if (num5 <= count2)
											{
												if (object.Equals(pE2.ShapeId.Value, list[num5 - 1].ShapeId.Value))
												{
													while (true)
													{
														switch (4)
														{
														case 0:
															continue;
														}
														break;
													}
													num3 = num5;
													break;
												}
												num5++;
												continue;
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
											break;
										}
										if (!num3.HasValue)
										{
											break;
										}
										while (true)
										{
											switch (6)
											{
											case 0:
												continue;
											}
											if (num3.Value != num2)
											{
												try
												{
													this.m_A[num][num3.Value].MoveTo(num2);
													PE item = list[num3.Value - 1];
													list.RemoveAt(num3.Value - 1);
													list.Insert(num2 - 1, item);
												}
												finally
												{
												}
											}
											num2++;
											break;
										}
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
									break;
								}
								B.Remove(list);
								return;
							}
						}
					}
				}
			}
		}

		protected virtual void A(bool A)
		{
			if (this.m_A)
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
				this.m_A = null;
				this.m_A = null;
				List<List<PE>> a = this.m_A;
				if (a == null)
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
				}
				else
				{
					a.Clear();
				}
				this.m_A = null;
				this.m_A = null;
				this.m_A = null;
				this.m_A = true;
				return;
			}
		}

		~JE()
		{
			A(A: false);
			base.Finalize();
		}

		public void Dispose()
		{
			A(A: true);
			GC.SuppressFinalize(this);
		}

		void IDisposable.Dispose()
		{
			//ILSpy generated this explicit interface implementation from .override directive in Dispose
			this.Dispose();
		}
	}

	private sealed class PE
	{
		[CompilerGenerated]
		internal sealed class KE
		{
			public Effect A;

			public Timing A;

			public PE A;

			[SpecialName]
			internal void A()
			{
				this.A.Exit = this.A.Exit;
			}

			[SpecialName]
			internal void B()
			{
				this.A.Paragraph = this.A.Paragraph;
			}

			[SpecialName]
			internal void C()
			{
				this.A = this.A.Timing;
			}
		}

		[CompilerGenerated]
		internal sealed class LE
		{
			public Effect A;

			public PE A;

			[SpecialName]
			internal void A()
			{
				if (!this.A.Exit.HasValue)
				{
					return;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					this.A.Exit = this.A.Exit.Value;
					return;
				}
			}

			[SpecialName]
			internal void B()
			{
				if (!this.A.Paragraph.HasValue)
				{
					return;
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					this.A.Paragraph = this.A.Paragraph.Value;
					return;
				}
			}
		}

		[CompilerGenerated]
		internal sealed class ME
		{
			public Timing A;

			public PE A;

			[SpecialName]
			internal void A()
			{
				this.A.Duration = this.A?.Duration;
			}

			[SpecialName]
			internal void B()
			{
				PE pE = this.A;
				Timing timing = this.A;
				float? triggerDelayTime;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					triggerDelayTime = null;
				}
				else
				{
					triggerDelayTime = timing.TriggerDelayTime;
				}
				pE.TriggerDelayTime = triggerDelayTime;
			}

			[SpecialName]
			internal void C()
			{
				PE pE = this.A;
				Timing timing = this.A;
				float? accelerate;
				if (timing == null)
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
					accelerate = null;
				}
				else
				{
					accelerate = timing.Accelerate;
				}
				pE.Accelerate = accelerate;
			}

			[SpecialName]
			internal void D()
			{
				PE pE = this.A;
				Timing timing = this.A;
				float? decelerate;
				if (timing == null)
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
					decelerate = null;
				}
				else
				{
					decelerate = timing.Decelerate;
				}
				pE.Decelerate = decelerate;
			}

			[SpecialName]
			internal void E()
			{
				PE pE = this.A;
				Timing timing = this.A;
				float? speed;
				if (timing == null)
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
					speed = null;
				}
				else
				{
					speed = timing.Speed;
				}
				pE.Speed = speed;
			}

			[SpecialName]
			internal void F()
			{
				PE pE = this.A;
				Timing timing = this.A;
				MsoTriState? autoReverse;
				if (timing == null)
				{
					while (true)
					{
						switch (4)
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
					autoReverse = null;
				}
				else
				{
					autoReverse = timing.AutoReverse;
				}
				pE.AutoReverse = autoReverse;
			}

			[SpecialName]
			internal void G()
			{
				PE pE = this.A;
				Timing timing = this.A;
				MsoTriState? bounceEnd;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					bounceEnd = null;
				}
				else
				{
					bounceEnd = timing.BounceEnd;
				}
				pE.BounceEnd = bounceEnd;
			}

			[SpecialName]
			internal void H()
			{
				PE pE = this.A;
				Timing timing = this.A;
				float? bounceEndIntensity;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					bounceEndIntensity = null;
				}
				else
				{
					bounceEndIntensity = timing.BounceEndIntensity;
				}
				pE.BounceEndIntensity = bounceEndIntensity;
			}

			[SpecialName]
			internal void I()
			{
				PE pE = this.A;
				Timing timing = this.A;
				int? repeatCount;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					repeatCount = null;
				}
				else
				{
					repeatCount = timing.RepeatCount;
				}
				pE.RepeatCount = repeatCount;
			}

			[SpecialName]
			internal void J()
			{
				this.A.RepeatDuration = this.A?.RepeatDuration;
			}

			[SpecialName]
			internal void K()
			{
				PE pE = this.A;
				Timing timing = this.A;
				MsoAnimEffectRestart? restart;
				if (timing == null)
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
					restart = null;
				}
				else
				{
					restart = timing.Restart;
				}
				pE.Restart = restart;
			}

			[SpecialName]
			internal void L()
			{
				PE pE = this.A;
				Timing timing = this.A;
				MsoTriState? rewindAtEnd;
				if (timing == null)
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
					rewindAtEnd = null;
				}
				else
				{
					rewindAtEnd = timing.RewindAtEnd;
				}
				pE.RewindAtEnd = rewindAtEnd;
			}

			[SpecialName]
			internal void M()
			{
				this.A.SmoothStart = this.A?.SmoothStart;
			}

			[SpecialName]
			internal void N()
			{
				PE pE = this.A;
				Timing timing = this.A;
				MsoTriState? smoothEnd;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					smoothEnd = null;
				}
				else
				{
					smoothEnd = timing.SmoothEnd;
				}
				pE.SmoothEnd = smoothEnd;
			}
		}

		[CompilerGenerated]
		internal sealed class NE
		{
			public Timing A;

			public PE A;

			[SpecialName]
			internal void A()
			{
				if (!this.A.Duration.HasValue)
				{
					return;
				}
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
					this.A.Duration = this.A.Duration.Value;
					return;
				}
			}

			[SpecialName]
			internal void B()
			{
				if (!this.A.TriggerDelayTime.HasValue)
				{
					return;
				}
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					this.A.TriggerDelayTime = this.A.TriggerDelayTime.Value;
					return;
				}
			}

			[SpecialName]
			internal void C()
			{
				if (this.A.Accelerate.HasValue)
				{
					this.A.Accelerate = this.A.Accelerate.Value;
				}
			}

			[SpecialName]
			internal void D()
			{
				if (!this.A.Decelerate.HasValue)
				{
					return;
				}
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
					this.A.Decelerate = this.A.Decelerate.Value;
					return;
				}
			}

			[SpecialName]
			internal void E()
			{
				if (!this.A.Speed.HasValue)
				{
					return;
				}
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
					this.A.Speed = this.A.Speed.Value;
					return;
				}
			}

			[SpecialName]
			internal void F()
			{
				if (!this.A.AutoReverse.HasValue)
				{
					return;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					this.A.AutoReverse = this.A.AutoReverse.Value;
					return;
				}
			}

			[SpecialName]
			internal void G()
			{
				if (this.A.BounceEnd.HasValue)
				{
					this.A.BounceEnd = this.A.BounceEnd.Value;
				}
			}

			[SpecialName]
			internal void H()
			{
				if (!this.A.BounceEndIntensity.HasValue)
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
					this.A.BounceEndIntensity = this.A.BounceEndIntensity.Value;
					return;
				}
			}

			[SpecialName]
			internal void I()
			{
				if (!this.A.RepeatCount.HasValue)
				{
					return;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					this.A.RepeatCount = this.A.RepeatCount.Value;
					return;
				}
			}

			[SpecialName]
			internal void J()
			{
				if (this.A.RepeatDuration.HasValue)
				{
					this.A.RepeatDuration = this.A.RepeatDuration.Value;
				}
			}

			[SpecialName]
			internal void K()
			{
				if (!this.A.Restart.HasValue)
				{
					return;
				}
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
					this.A.Restart = this.A.Restart.Value;
					return;
				}
			}

			[SpecialName]
			internal void L()
			{
				if (!this.A.RewindAtEnd.HasValue)
				{
					return;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					this.A.RewindAtEnd = this.A.RewindAtEnd.Value;
					return;
				}
			}

			[SpecialName]
			internal void M()
			{
				if (!this.A.SmoothStart.HasValue)
				{
					return;
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					this.A.SmoothStart = this.A.SmoothStart.Value;
					return;
				}
			}

			[SpecialName]
			internal void N()
			{
				if (!this.A.SmoothStart.HasValue)
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
					this.A.SmoothStart = this.A.SmoothEnd.Value;
					return;
				}
			}
		}

		[CompilerGenerated]
		internal sealed class OE
		{
			public Effect A;

			public PE A;

			[SpecialName]
			internal bool A()
			{
				return object.Equals(this.A.Exit, this.A.Exit);
			}

			[SpecialName]
			internal bool B()
			{
				return object.Equals(this.A.Paragraph, this.A.Paragraph);
			}
		}

		[CompilerGenerated]
		private int? m_A;

		[CompilerGenerated]
		private MsoAnimEffect? m_A;

		[CompilerGenerated]
		private MsoTriState? m_A;

		[CompilerGenerated]
		private int? m_B;

		[CompilerGenerated]
		private int? C;

		[CompilerGenerated]
		private MsoAnimTriggerType? m_A;

		[CompilerGenerated]
		private string m_A;

		[CompilerGenerated]
		private float? m_A;

		[CompilerGenerated]
		private float? m_B;

		[CompilerGenerated]
		private float? C;

		[CompilerGenerated]
		private float? D;

		[CompilerGenerated]
		private MsoTriState? m_B;

		[CompilerGenerated]
		private MsoTriState? C;

		[CompilerGenerated]
		private float? E;

		[CompilerGenerated]
		private int? D;

		[CompilerGenerated]
		private float? F;

		[CompilerGenerated]
		private MsoAnimEffectRestart? m_A;

		[CompilerGenerated]
		private MsoTriState? D;

		[CompilerGenerated]
		private MsoTriState? E;

		[CompilerGenerated]
		private MsoTriState? F;

		[CompilerGenerated]
		private float? G;

		internal int? ShapeId
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

		internal MsoAnimEffect? EffectType
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

		internal MsoTriState? Exit
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

		internal int? Paragraph
		{
			[CompilerGenerated]
			get
			{
				return this.m_B;
			}
			[CompilerGenerated]
			set
			{
				this.m_B = value;
			}
		}

		internal int? TriggerShapeId
		{
			[CompilerGenerated]
			get
			{
				return this.C;
			}
			[CompilerGenerated]
			set
			{
				this.C = value;
			}
		}

		internal MsoAnimTriggerType? TriggerType
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

		internal string TriggerBookmark
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

		internal float? Duration
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

		internal float? TriggerDelayTime
		{
			[CompilerGenerated]
			get
			{
				return this.m_B;
			}
			[CompilerGenerated]
			set
			{
				this.m_B = value;
			}
		}

		internal float? Accelerate
		{
			[CompilerGenerated]
			get
			{
				return this.C;
			}
			[CompilerGenerated]
			set
			{
				this.C = value;
			}
		}

		internal float? Decelerate
		{
			[CompilerGenerated]
			get
			{
				return this.D;
			}
			[CompilerGenerated]
			set
			{
				this.D = value;
			}
		}

		internal MsoTriState? AutoReverse
		{
			[CompilerGenerated]
			get
			{
				return this.m_B;
			}
			[CompilerGenerated]
			set
			{
				this.m_B = value;
			}
		}

		internal MsoTriState? BounceEnd
		{
			[CompilerGenerated]
			get
			{
				return C;
			}
			[CompilerGenerated]
			set
			{
				C = value;
			}
		}

		internal float? BounceEndIntensity
		{
			[CompilerGenerated]
			get
			{
				return this.E;
			}
			[CompilerGenerated]
			set
			{
				this.E = value;
			}
		}

		internal int? RepeatCount
		{
			[CompilerGenerated]
			get
			{
				return this.D;
			}
			[CompilerGenerated]
			set
			{
				this.D = value;
			}
		}

		internal float? RepeatDuration
		{
			[CompilerGenerated]
			get
			{
				return this.F;
			}
			[CompilerGenerated]
			set
			{
				this.F = value;
			}
		}

		internal MsoAnimEffectRestart? Restart
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

		internal MsoTriState? RewindAtEnd
		{
			[CompilerGenerated]
			get
			{
				return D;
			}
			[CompilerGenerated]
			set
			{
				D = value;
			}
		}

		internal MsoTriState? SmoothStart
		{
			[CompilerGenerated]
			get
			{
				return E;
			}
			[CompilerGenerated]
			set
			{
				E = value;
			}
		}

		internal MsoTriState? SmoothEnd
		{
			[CompilerGenerated]
			get
			{
				return F;
			}
			[CompilerGenerated]
			set
			{
				F = value;
			}
		}

		internal float? Speed
		{
			[CompilerGenerated]
			get
			{
				return G;
			}
			[CompilerGenerated]
			set
			{
				G = value;
			}
		}

		internal PE(Effect A)
		{
			PE A2 = this;
			Timing A3 = null;
			try
			{
				ShapeId = A.Shape.Id;
				EffectType = A.EffectType;
				PE.A([SpecialName] () =>
				{
					A2.Exit = A.Exit;
				});
				PE.A([SpecialName] () =>
				{
					A2.Paragraph = A.Paragraph;
				});
				PE.A([SpecialName] () =>
				{
					A3 = A.Timing;
				});
				Timing timing = A3;
				object obj;
				if (timing == null)
				{
					while (true)
					{
						switch (4)
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
					obj = null;
				}
				else
				{
					obj = timing.TriggerShape;
				}
				Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)obj;
				int? triggerShapeId;
				if (shape == null)
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
					triggerShapeId = null;
				}
				else
				{
					triggerShapeId = shape.Id;
				}
				TriggerShapeId = triggerShapeId;
				Timing timing2 = A3;
				MsoAnimTriggerType? triggerType;
				if (timing2 == null)
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
					triggerType = null;
				}
				else
				{
					triggerType = timing2.TriggerType;
				}
				TriggerType = triggerType;
				TriggerBookmark = PE.A(A3);
				this.A(A3);
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			finally
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = null;
				A3 = null;
			}
		}

		internal void A(Sequence A, int B, Dictionary<int, Microsoft.Office.Interop.PowerPoint.Shape> C)
		{
			Effect A2;
			try
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape = C[B];
				Microsoft.Office.Interop.PowerPoint.Shape shape2 = C[TriggerShapeId.Value];
				Microsoft.Office.Interop.PowerPoint.Shape pShape = shape;
				MsoAnimEffect value = EffectType.Value;
				MsoAnimTriggerType value2 = TriggerType.Value;
				Microsoft.Office.Interop.PowerPoint.Shape pTriggerShape = shape2;
				object obj = TriggerBookmark;
				if (obj == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					obj = "";
				}
				A2 = A.AddTriggerEffect(pShape, value, value2, pTriggerShape, (string)obj);
				PE.A([SpecialName] () =>
				{
					if (Exit.HasValue)
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
								A2.Exit = Exit.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (Paragraph.HasValue)
					{
						while (true)
						{
							switch (2)
							{
							case 0:
								break;
							default:
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								A2.Paragraph = Paragraph.Value;
								return;
							}
						}
					}
				});
				this.A(A2);
			}
			finally
			{
				Microsoft.Office.Interop.PowerPoint.Shape shape2 = null;
				Microsoft.Office.Interop.PowerPoint.Shape shape = null;
				A2 = null;
			}
		}

		private void A(Timing A)
		{
			PE.A([SpecialName] () =>
			{
				Duration = A?.Duration;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				float? triggerDelayTime;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					triggerDelayTime = null;
				}
				else
				{
					triggerDelayTime = timing.TriggerDelayTime;
				}
				pE.TriggerDelayTime = triggerDelayTime;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				float? accelerate;
				if (timing == null)
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
					accelerate = null;
				}
				else
				{
					accelerate = timing.Accelerate;
				}
				pE.Accelerate = accelerate;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				float? decelerate;
				if (timing == null)
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
					decelerate = null;
				}
				else
				{
					decelerate = timing.Decelerate;
				}
				pE.Decelerate = decelerate;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				float? speed;
				if (timing == null)
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
					speed = null;
				}
				else
				{
					speed = timing.Speed;
				}
				pE.Speed = speed;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				MsoTriState? autoReverse;
				if (timing == null)
				{
					while (true)
					{
						switch (4)
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
					autoReverse = null;
				}
				else
				{
					autoReverse = timing.AutoReverse;
				}
				pE.AutoReverse = autoReverse;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				MsoTriState? bounceEnd;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					bounceEnd = null;
				}
				else
				{
					bounceEnd = timing.BounceEnd;
				}
				pE.BounceEnd = bounceEnd;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				float? bounceEndIntensity;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					bounceEndIntensity = null;
				}
				else
				{
					bounceEndIntensity = timing.BounceEndIntensity;
				}
				pE.BounceEndIntensity = bounceEndIntensity;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				int? repeatCount;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					repeatCount = null;
				}
				else
				{
					repeatCount = timing.RepeatCount;
				}
				pE.RepeatCount = repeatCount;
			});
			PE.A([SpecialName] () =>
			{
				RepeatDuration = A?.RepeatDuration;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				MsoAnimEffectRestart? restart;
				if (timing == null)
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
					restart = null;
				}
				else
				{
					restart = timing.Restart;
				}
				pE.Restart = restart;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				MsoTriState? rewindAtEnd;
				if (timing == null)
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
					rewindAtEnd = null;
				}
				else
				{
					rewindAtEnd = timing.RewindAtEnd;
				}
				pE.RewindAtEnd = rewindAtEnd;
			});
			PE.A([SpecialName] () =>
			{
				SmoothStart = A?.SmoothStart;
			});
			PE.A([SpecialName] () =>
			{
				PE pE = this;
				Timing timing = A;
				MsoTriState? smoothEnd;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					smoothEnd = null;
				}
				else
				{
					smoothEnd = timing.SmoothEnd;
				}
				pE.SmoothEnd = smoothEnd;
			});
		}

		private void A(Effect A)
		{
			Timing A2;
			try
			{
				A2 = A.Timing;
				PE.A([SpecialName] () =>
				{
					if (Duration.HasValue)
					{
						while (true)
						{
							switch (1)
							{
							case 0:
								break;
							default:
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								A2.Duration = Duration.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (TriggerDelayTime.HasValue)
					{
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								A2.TriggerDelayTime = TriggerDelayTime.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (Accelerate.HasValue)
					{
						A2.Accelerate = Accelerate.Value;
					}
				});
				PE.A([SpecialName] () =>
				{
					if (Decelerate.HasValue)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								A2.Decelerate = Decelerate.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (Speed.HasValue)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								A2.Speed = Speed.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (AutoReverse.HasValue)
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
								A2.AutoReverse = AutoReverse.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (BounceEnd.HasValue)
					{
						A2.BounceEnd = BounceEnd.Value;
					}
				});
				PE.A([SpecialName] () =>
				{
					if (BounceEndIntensity.HasValue)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								break;
							default:
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								A2.BounceEndIntensity = BounceEndIntensity.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (RepeatCount.HasValue)
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
								A2.RepeatCount = RepeatCount.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (RepeatDuration.HasValue)
					{
						A2.RepeatDuration = RepeatDuration.Value;
					}
				});
				PE.A([SpecialName] () =>
				{
					if (Restart.HasValue)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								A2.Restart = Restart.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (RewindAtEnd.HasValue)
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
								A2.RewindAtEnd = RewindAtEnd.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (SmoothStart.HasValue)
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
								A2.SmoothStart = SmoothStart.Value;
								return;
							}
						}
					}
				});
				PE.A([SpecialName] () =>
				{
					if (SmoothStart.HasValue)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								if (1 == 0)
								{
									/*OpCode not supported: LdMemberToken*/;
								}
								A2.SmoothStart = SmoothEnd.Value;
								return;
							}
						}
					}
				});
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
			finally
			{
				A2 = null;
			}
		}

		internal bool A()
		{
			if (TriggerShapeId.HasValue)
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
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				if (TriggerType.HasValue)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							return EffectType.HasValue;
						}
					}
				}
			}
			return false;
		}

		internal bool A(Effect A)
		{
			try
			{
				Timing timing = A.Timing;
				object objA = TriggerType;
				MsoAnimTriggerType? obj;
				if (timing == null)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					obj = null;
				}
				else
				{
					obj = timing.TriggerType;
				}
				int result;
				if (object.Equals(objA, obj))
				{
					object objA2 = TriggerShapeId;
					int? obj2;
					if (timing == null)
					{
						obj2 = null;
					}
					else
					{
						Microsoft.Office.Interop.PowerPoint.Shape triggerShape = timing.TriggerShape;
						if (triggerShape == null)
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
							obj2 = null;
						}
						else
						{
							obj2 = triggerShape.Id;
						}
					}
					if (object.Equals(objA2, obj2))
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
						result = (object.Equals(TriggerBookmark, PE.A(timing)) ? 1 : 0);
						goto IL_00cb;
					}
				}
				result = 0;
				goto IL_00cb;
				IL_00cb:
				return (byte)result != 0;
			}
			finally
			{
				Timing timing = null;
			}
		}

		internal bool B(Effect A)
		{
			try
			{
				int result;
				if (object.Equals(EffectType, A.EffectType))
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
					object objA = ShapeId;
					Microsoft.Office.Interop.PowerPoint.Shape shape = A.Shape;
					int? obj;
					if (shape == null)
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
						obj = null;
					}
					else
					{
						obj = shape.Id;
					}
					if (object.Equals(objA, obj))
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
						if (PE.A([SpecialName] () => object.Equals(Exit, A.Exit)))
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
							result = (PE.A([SpecialName] () => object.Equals(Paragraph, A.Paragraph)) ? 1 : 0);
							goto IL_00d4;
						}
					}
				}
				result = 0;
				goto IL_00d4;
				IL_00d4:
				return (byte)result != 0;
			}
			finally
			{
			}
		}

		private static string A(Timing A)
		{
			string result;
			try
			{
				object obj;
				if (A == null)
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
					obj = null;
				}
				else
				{
					obj = A.TriggerBookmark;
				}
				result = (string)obj;
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				result = null;
				ProjectData.ClearProjectError();
			}
			return result;
		}

		private static void A(Action A)
		{
			try
			{
				A();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				ProjectData.ClearProjectError();
			}
		}

		private static bool A(Func<bool> A, bool B = true)
		{
			bool result;
			try
			{
				result = A();
			}
			catch (Exception projectError)
			{
				ProjectData.SetProjectError(projectError);
				result = B;
				ProjectData.ClearProjectError();
			}
			return result;
		}
	}

	private sealed class QE
	{
		[CompilerGenerated]
		private Effect A;

		[CompilerGenerated]
		private int A;

		[CompilerGenerated]
		private int B;

		internal Effect Effect
		{
			[CompilerGenerated]
			get
			{
				return this.A;
			}
			[CompilerGenerated]
			set
			{
				this.A = value;
			}
		}

		internal int CurPos
		{
			[CompilerGenerated]
			get
			{
				return A;
			}
			[CompilerGenerated]
			set
			{
				A = value;
			}
		}

		internal int NewPos
		{
			[CompilerGenerated]
			get
			{
				return B;
			}
			[CompilerGenerated]
			set
			{
				B = value;
			}
		}

		internal QE(Effect A, int B, int C)
		{
			Effect = A;
			CurPos = B;
			NewPos = C;
		}
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<List<int>> A;

		public static Func<List<List<PE>>> A;

		public static Func<QE, bool> A;

		public static Func<QE, int> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal List<int> A()
		{
			return new List<int>();
		}

		[SpecialName]
		internal List<List<PE>> A()
		{
			return new List<List<PE>>();
		}

		[SpecialName]
		internal bool A(QE A)
		{
			if (A.NewPos > 0)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return A.NewPos != A.CurPos;
					}
				}
			}
			return false;
		}

		[SpecialName]
		internal int A(QE A)
		{
			return A.NewPos;
		}
	}

	[CompilerGenerated]
	internal sealed class RE
	{
		public int A;

		public Predicate<int> A;

		public RE(RE A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}

		[SpecialName]
		internal bool A(int A)
		{
			return A == this.A;
		}
	}

	[CompilerGenerated]
	internal sealed class SE
	{
		public List<int> A;

		public List<int> B;

		public UE A;

		public SE(SE A)
		{
			if (A == null)
			{
				return;
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				this.A = A.A;
				B = A.B;
				return;
			}
		}

		[SpecialName]
		internal void A(Effect A, Microsoft.Office.Interop.PowerPoint.Shape B, int C)
		{
			TE tE = new TE(tE)
			{
				A = B
			};
			int num = -1;
			checked
			{
				while (true)
				{
					num = this.A.FindIndex(num + 1, tE.A);
					if (num < 0)
					{
						break;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						if (this.B.Contains(num))
						{
							break;
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
						this.B.Add(num);
						goto end_IL_0010;
					}
					continue;
					end_IL_0010:
					break;
				}
				this.A.A.Add(new QE(A, C, num + 1));
			}
		}
	}

	[CompilerGenerated]
	internal sealed class TE
	{
		public Microsoft.Office.Interop.PowerPoint.Shape A;

		public TE(TE A)
		{
			if (A == null)
			{
				return;
			}
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
				this.A = A.A;
				return;
			}
		}

		[SpecialName]
		internal bool A(int A)
		{
			return A == this.A.Id;
		}
	}

	[CompilerGenerated]
	internal sealed class UE
	{
		public List<QE> A;

		public UE(UE A)
		{
			if (A != null)
			{
				this.A = A.A;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class VE
	{
		public int A;

		public int B;

		public VE(VE A)
		{
			if (A == null)
			{
				return;
			}
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
				this.A = A.A;
				this.B = A.B;
				return;
			}
		}

		[SpecialName]
		internal void A(QE A)
		{
			if (A.CurPos <= this.A)
			{
				return;
			}
			checked
			{
				while (true)
				{
					switch (4)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (A.CurPos <= this.B)
					{
						A.CurPos--;
					}
					return;
				}
			}
		}

		[SpecialName]
		internal void B(QE A)
		{
			if (A.CurPos < this.B)
			{
				return;
			}
			checked
			{
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
					if (A.CurPos >= this.A)
					{
						return;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						A.CurPos++;
						return;
					}
				}
			}
		}
	}

	[CompilerGenerated]
	internal sealed class WE
	{
		public List<int> A;

		[SpecialName]
		internal void A(Effect A, Microsoft.Office.Interop.PowerPoint.Shape B, int C)
		{
			this.A.Add(B.Id);
		}

		[SpecialName]
		internal void A()
		{
			this.A.Add(int.MinValue);
		}
	}

	private readonly Dictionary<int, List<int>> m_A;

	private Dictionary<int, List<List<PE>>> m_A;

	private const int m_A = int.MinValue;

	public TimelineRestorer()
	{
		this.m_A = new Dictionary<int, List<int>>();
		this.m_A = new Dictionary<int, List<List<PE>>>();
	}

	internal void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		try
		{
			Slide slideFromShape = clsPowerPoint.GetSlideFromShape(A);
			if (slideFromShape == null)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return;
					}
				}
			}
			int slideID = slideFromShape.SlideID;
			if (!this.m_A.ContainsKey(slideID))
			{
				this.m_A[slideID] = TimelineRestorer.A(slideFromShape);
			}
			if (this.m_A.ContainsKey(slideID))
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
				this.m_A[slideID] = TimelineRestorer.A(slideFromShape);
				return;
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Slide slideFromShape = null;
		}
	}

	internal void A(Microsoft.Office.Interop.PowerPoint.Shape A, int B)
	{
		RE a = default(RE);
		RE CS_0024_003C_003E8__locals7 = new RE(a);
		CS_0024_003C_003E8__locals7.A = B;
		try
		{
			Slide slideFromShape = clsPowerPoint.GetSlideFromShape(A);
			if (slideFromShape == null)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return;
					}
				}
			}
			int slideID = slideFromShape.SlideID;
			Dictionary<int, List<int>> a2 = this.m_A;
			Func<List<int>> c;
			if (_Closure_0024__.A == null)
			{
				c = (_Closure_0024__.A = [SpecialName] () => new List<int>());
			}
			else
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
				c = _Closure_0024__.A;
			}
			List<int> list = TimelineRestorer.A(a2, slideID, c);
			int num = -1;
			while (true)
			{
				int startIndex = checked(num + 1);
				Predicate<int> match;
				if (CS_0024_003C_003E8__locals7.A != null)
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
					match = CS_0024_003C_003E8__locals7.A;
				}
				else
				{
					match = (CS_0024_003C_003E8__locals7.A = [SpecialName] (int num2) => num2 == CS_0024_003C_003E8__locals7.A);
				}
				num = list.FindIndex(startIndex, match);
				if (num < 0)
				{
					break;
				}
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						goto end_IL_00bd;
					}
					continue;
					end_IL_00bd:
					break;
				}
				list[num] = A.Id;
			}
			List<List<PE>>.Enumerator enumerator = default(List<List<PE>>.Enumerator);
			try
			{
				Dictionary<int, List<List<PE>>> a3 = this.m_A;
				Func<List<List<PE>>> c2;
				if (_Closure_0024__.A == null)
				{
					c2 = (_Closure_0024__.A = [SpecialName] () => new List<List<PE>>());
				}
				else
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
					c2 = _Closure_0024__.A;
				}
				enumerator = TimelineRestorer.A(a3, slideID, c2).GetEnumerator();
				while (enumerator.MoveNext())
				{
					List<PE> current = enumerator.Current;
					using List<PE>.Enumerator enumerator2 = current.GetEnumerator();
					while (enumerator2.MoveNext())
					{
						PE current2 = enumerator2.Current;
						if (object.Equals(current2.ShapeId, CS_0024_003C_003E8__locals7.A))
						{
							current2.ShapeId = A.Id;
						}
						if (!object.Equals(current2.TriggerShapeId, CS_0024_003C_003E8__locals7.A))
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
						current2.TriggerShapeId = A.Id;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							goto end_IL_01bf;
						}
						continue;
						end_IL_01bf:
						break;
					}
				}
				while (true)
				{
					switch (5)
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
				((IDisposable)enumerator/*cast due to .constrained prefix*/).Dispose();
			}
		}
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Slide slideFromShape = null;
		}
	}

	internal void A()
	{
		try
		{
			Microsoft.Office.Interop.PowerPoint.Slides slides = NG.A.Application.ActivePresentation.Slides;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = slides.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Slide a = (Slide)enumerator.Current;
					A(a);
				}
				while (true)
				{
					switch (3)
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
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (1)
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
		catch (Exception projectError)
		{
			ProjectData.SetProjectError(projectError);
			ProjectData.ClearProjectError();
		}
		finally
		{
			this.m_A.Clear();
			this.m_A = null;
			Microsoft.Office.Interop.PowerPoint.Slides slides = null;
		}
	}

	private void A(Slide A)
	{
		UE uE = new UE(uE);
		uE.A = new List<QE>();
		checked
		{
			try
			{
				SE a = default(SE);
				SE CS_0024_003C_003E8__locals13 = new SE(a);
				CS_0024_003C_003E8__locals13.A = uE;
				int slideID = A.SlideID;
				if (!this.m_A.ContainsKey(slideID))
				{
					while (true)
					{
						switch (6)
						{
						case 0:
							break;
						default:
							if (1 == 0)
							{
								/*OpCode not supported: LdMemberToken*/;
							}
							return;
						}
					}
				}
				JE jE = new JE(A, this);
				try
				{
					jE.B();
				}
				finally
				{
					if (jE != null)
					{
						while (true)
						{
							switch (6)
							{
							case 0:
								continue;
							}
							((IDisposable)jE).Dispose();
							break;
						}
					}
				}
				CS_0024_003C_003E8__locals13.A = this.m_A[slideID];
				if (CS_0024_003C_003E8__locals13.A.Count == 0)
				{
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
				CS_0024_003C_003E8__locals13.B = new List<int>();
				TimelineRestorer.A(A, [SpecialName] (Effect a2, Microsoft.Office.Interop.PowerPoint.Shape B, int C) =>
				{
					TE tE = new TE(tE);
					tE.A = B;
					int num3 = -1;
					while (true)
					{
						num3 = CS_0024_003C_003E8__locals13.A.FindIndex(num3 + 1, tE.A);
						if (num3 < 0)
						{
							break;
						}
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
								if (CS_0024_003C_003E8__locals13.B.Contains(num3))
								{
									goto end_IL_0031;
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
								CS_0024_003C_003E8__locals13.B.Add(num3);
								goto end_IL_0010;
							}
							continue;
							end_IL_0031:
							break;
						}
						continue;
						end_IL_0010:
						break;
					}
					CS_0024_003C_003E8__locals13.A.A.Add(new QE(a2, C, num3 + 1));
				});
				long num = CS_0024_003C_003E8__locals13.A.A.Count * (CS_0024_003C_003E8__locals13.A.A.Count - 1);
				long num2 = -1L;
				VE vE = default(VE);
				while (true)
				{
					vE = new VE(vE);
					num2++;
					if (num2 == num)
					{
						while (true)
						{
							switch (7)
							{
							case 0:
								break;
							default:
								return;
							}
						}
					}
					List<QE> source = CS_0024_003C_003E8__locals13.A.A;
					Func<QE, bool> predicate;
					if (_Closure_0024__.A == null)
					{
						predicate = (_Closure_0024__.A = [SpecialName] (QE qE2) =>
						{
							if (qE2.NewPos > 0)
							{
								while (true)
								{
									switch (3)
									{
									case 0:
										break;
									default:
										if (1 == 0)
										{
											/*OpCode not supported: LdMemberToken*/;
										}
										return qE2.NewPos != qE2.CurPos;
									}
								}
							}
							return false;
						});
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
						predicate = _Closure_0024__.A;
					}
					IEnumerable<QE> source2 = source.Where(predicate);
					Func<QE, int> keySelector;
					if (_Closure_0024__.A == null)
					{
						keySelector = (_Closure_0024__.A = [SpecialName] (QE qE2) => qE2.NewPos);
					}
					else
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
						keySelector = _Closure_0024__.A;
					}
					QE qE = source2.OrderByDescending(keySelector).FirstOrDefault();
					if (qE == null)
					{
						break;
					}
					qE.Effect.MoveTo(qE.NewPos);
					vE.A = qE.CurPos;
					vE.B = qE.NewPos;
					if (vE.A < vE.B)
					{
						CS_0024_003C_003E8__locals13.A.A.ForEach(vE.A);
					}
					else
					{
						CS_0024_003C_003E8__locals13.A.A.ForEach(vE.B);
					}
					qE.CurPos = vE.B;
				}
				while (true)
				{
					switch (3)
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
				uE.A.Clear();
				uE.A = null;
			}
		}
	}

	private static List<int> A(Slide A)
	{
		List<int> A2 = new List<int>();
		TimelineRestorer.A(A, [SpecialName] (Effect effect, Microsoft.Office.Interop.PowerPoint.Shape B, int C) =>
		{
			A2.Add(B.Id);
		}, [SpecialName] () =>
		{
			A2.Add(int.MinValue);
		});
		return A2;
	}

	private static List<List<PE>> A(Slide A)
	{
		List<List<PE>> list = new List<List<PE>>();
		try
		{
			Sequences interactiveSequences = A.TimeLine.InteractiveSequences;
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = interactiveSequences.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Sequence sequence = (Sequence)enumerator.Current;
					List<PE> list2 = new List<PE>();
					try
					{
						int count = sequence.Count;
						for (int i = 1; i <= count; i = checked(i + 1))
						{
							try
							{
								Effect a = sequence[i];
								list2.Add(new PE(a));
							}
							finally
							{
								Effect a = null;
							}
						}
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
							list.Add(list2);
							break;
						}
					}
					catch (Exception projectError)
					{
						ProjectData.SetProjectError(projectError);
						ProjectData.ClearProjectError();
					}
					finally
					{
						sequence = null;
					}
				}
				return list;
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
		finally
		{
			Sequences interactiveSequences = null;
		}
	}

	private static void A(Slide A, Action<Effect, Microsoft.Office.Interop.PowerPoint.Shape, int> B, Action C = null)
	{
		try
		{
			Sequence mainSequence = A.TimeLine.MainSequence;
			int count = mainSequence.Count;
			for (int i = 1; i <= count; i = checked(i + 1))
			{
				try
				{
					Effect effect = mainSequence[i];
					Microsoft.Office.Interop.PowerPoint.Shape shape = effect.Shape;
					B(effect, shape, i);
				}
				catch (Exception projectError)
				{
					ProjectData.SetProjectError(projectError);
					C?.Invoke();
					ProjectData.ClearProjectError();
				}
				finally
				{
					Microsoft.Office.Interop.PowerPoint.Shape shape = null;
					Effect effect = null;
				}
			}
			while (true)
			{
				switch (3)
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
		catch (Exception projectError2)
		{
			ProjectData.SetProjectError(projectError2);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Sequence mainSequence = null;
		}
	}

	private static B A<A, B>(Dictionary<A, B> A, A B, Func<B> C)
	{
		if (!A.ContainsKey(B))
		{
			return C();
		}
		return A[B];
	}
}
