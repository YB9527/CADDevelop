using System;
using System.Reflection;

namespace YanBo_CG.ReFlect
{
	public class Clazz
	{
		private string fullClassName;

		private string function;

		private Type paramterType;

		private Type resultType;

		public MethodInfo GetMethodInfo;

		public MethodInfo SetMethodInfo;

		public string getFullClassName()
		{
			return this.fullClassName;
		}

		public void setFullClassName(string fullClassName)
		{
			this.fullClassName = fullClassName;
		}

		public string getFunction()
		{
			return this.function;
		}

		public void setFunction(string function)
		{
			this.function = function;
		}

		public Type getParamterType()
		{
			return this.paramterType;
		}

		public void setParamterType(Type paramterType)
		{
			this.paramterType = paramterType;
		}

		public Type getResultType()
		{
			return this.resultType;
		}

		public void setResultType(Type resultType)
		{
			this.resultType = resultType;
		}
	}
}
