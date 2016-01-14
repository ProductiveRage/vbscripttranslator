using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace VBScriptTranslator.RuntimeSupport.Compat
{
    [ComVisible(true)]
    public class BasicIReflectImplementation : IReflect
	{
		public virtual FieldInfo GetField(string name, BindingFlags bindingAttr)
		{
			return this.GetType().GetField(name, bindingAttr);
		}

		public virtual FieldInfo[] GetFields(BindingFlags bindingAttr)
		{
			return this.GetType().GetFields(bindingAttr);
		}

		public virtual MemberInfo[] GetMember(string name, BindingFlags bindingAttr)
		{
			return this.GetType().GetMember(name, bindingAttr);
		}

		public virtual MemberInfo[] GetMembers(BindingFlags bindingAttr)
		{
			return this.GetType().GetMembers(bindingAttr);
		}

		public virtual MethodInfo GetMethod(string name, BindingFlags bindingAttr)
		{
			return this.GetType().GetMethod(name, bindingAttr);
		}

		public virtual MethodInfo GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers)
		{
			return this.GetType().GetMethod(name, bindingAttr, binder, types, modifiers);
		}

		public virtual MethodInfo[] GetMethods(BindingFlags bindingAttr)
		{
			return this.GetType().GetMethods(bindingAttr);
		}

		public virtual PropertyInfo[] GetProperties(BindingFlags bindingAttr)
		{
			return this.GetType().GetProperties(bindingAttr);
		}

		public virtual PropertyInfo GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
		{
			return this.GetType().GetProperty(name, bindingAttr, binder, returnType, types, modifiers);
		}

		public virtual PropertyInfo GetProperty(string name, BindingFlags bindingAttr)
		{
			return this.GetType().GetProperty(name, bindingAttr);
		}

		public virtual object InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target, object[] args, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters)
		{
			return this.GetType().InvokeMember(name, invokeAttr, binder, target, args, modifiers, culture, namedParameters);
		}

		public virtual Type UnderlyingSystemType
		{
			get { return this.GetType().UnderlyingSystemType; }
		}
	}
}
