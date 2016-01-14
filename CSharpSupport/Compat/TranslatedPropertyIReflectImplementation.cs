using VBScriptTranslator.RuntimeSupport.Attributes;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace VBScriptTranslator.RuntimeSupport.Compat
{
    /// <summary>
    /// C# doesn't support named indexed-properties so indexed VBScript properties will be represented by separate getter and setter methods (unless the property
    /// is read-only or write-only, in which case there will be only a getter or a setter). Generated classes that use these property representations will inherit
    /// from this which route requests from COM components that rquire the property access to the appropriate methods. These methods are identified by the presence
    /// of the TranslatedProperty attribute.
    /// </summary>
    [ComVisible(true)]
    public class TranslatedPropertyIReflectImplementation : BasicIReflectImplementation
	{
        private readonly IEnumerable<PropertyInfo> _representedPublicProperties;
        public TranslatedPropertyIReflectImplementation()
        {
            _representedPublicProperties = GetPublicTranslatedPropertyRepresentations(this.GetType());
        }

        public override PropertyInfo[] GetProperties(BindingFlags bindingAttr)
        {
            var baseProperties = base.GetProperties(bindingAttr);
            if (!IsBindingRequestFor(bindingAttr, BindingFlags.Public | BindingFlags.Instance))
                return baseProperties;
            return _representedPublicProperties.Concat(baseProperties).ToArray();
        }

        public override PropertyInfo GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
        {
            // For our purposes, we should be able to get away with ignoring all of the possible binder options and return and parameter types since
            // the only valid, operationl code is supported for translation. So we'll assume that if a property is being called that the return type
            // and parameter types are matching those requested.
            return GetProperty(name, bindingAttr);
        }

        public override PropertyInfo GetProperty(string name, BindingFlags bindingAttr)
        {
            return TryToGetRepresentedProperty(name, bindingAttr) ?? base.GetProperty(name, bindingAttr);
        }

        public override object InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target, object[] args, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters)
        {
            // See notes in GetProperty method above whose signature includes binder, args, modifiers, etc. as to why we are ignoring them here
            var representedProperty = TryToGetRepresentedProperty(name, invokeAttr);
            if (representedProperty != null)
            {
                MethodInfo method;
                if (IsBindingRequestFor(invokeAttr, BindingFlags.GetProperty))
                    method = representedProperty.GetGetMethod();
                else
                    method = representedProperty.GetSetMethod();
                return method.Invoke(target, args);
            }
            return base.InvokeMember(name, invokeAttr, binder, target, args, modifiers, culture, namedParameters);
        }

        private PropertyInfo TryToGetRepresentedProperty(string name, BindingFlags bindingAttr)
        {
            var stringComparison = IsBindingRequestFor(bindingAttr, BindingFlags.IgnoreCase)
                ? StringComparison.InvariantCultureIgnoreCase
                : StringComparison.InvariantCulture;
            var representedProperty = _representedPublicProperties.FirstOrDefault(p => p.Name.Equals(name, stringComparison));
            if (representedProperty == null)
                return null;

            if (IsBindingRequestFor(bindingAttr, BindingFlags.GetProperty) && !representedProperty.CanRead)
                return null;
            if (IsBindingRequestFor(bindingAttr, BindingFlags.PutDispProperty) && !representedProperty.CanWrite)
                return null;
            if (IsBindingRequestFor(bindingAttr, BindingFlags.PutRefDispProperty))
                return null;

            return representedProperty;
        }

        private bool IsBindingRequestFor(BindingFlags value, BindingFlags valueToSearchFor)
        {
            return (value & valueToSearchFor) == valueToSearchFor;
        }

        private static IEnumerable<PropertyInfo> GetPublicTranslatedPropertyRepresentations(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            var propertyTranslatedMethodGroups = type.GetMethods()
                .Select(m => new { Method = m, TranslatedPropertyAttribute = GetTranslatedPropertyAttributeIfAny(m) })
                .Where(m => m.TranslatedPropertyAttribute != null)
                .GroupBy(m => m.TranslatedPropertyAttribute.Name);
            var representedProperties = new List<PropertyInfo>();
            foreach (var methodGroup in propertyTranslatedMethodGroups)
            {
                if (methodGroup.Count() > 2)
                    throw new ArgumentException("There may not be more than two methods with the same name specified by a TranslatedProperty attribute (" + methodGroup.First().TranslatedPropertyAttribute.Name + ")");
                var getters = methodGroup.Where(m => m.Method.ReturnType != typeof(void));
                if (getters.Count() > 1)
                    throw new ArgumentException("There may not be more than two methods with the same name specified by a TranslatedProperty attribute (" + methodGroup.First().TranslatedPropertyAttribute.Name + ") that have a non-void return type (the getters)");
                var setters = methodGroup.Where(m => m.Method.ReturnType == typeof(void));
                if (setters.Count() > 1)
                    throw new ArgumentException("There may not be more than two methods with the same name specified by a TranslatedProperty attribute (" + methodGroup.First().TranslatedPropertyAttribute.Name + ") that have a void return type (the setters)");
                var getterMethod = (getters.Any() ? getters.Single().Method : null);
                var setterMethod = (setters.Any() ? setters.Single().Method : null);
                representedProperties.Add(
                    new TranslatedPropertyInfo(
                        methodGroup.Key,
                        type,
                        getterMethod,
                        setterMethod
                    )
                );
            }
            return representedProperties;
        }

        private static TranslatedProperty GetTranslatedPropertyAttributeIfAny(MethodInfo method)
        {
            if (method == null)
                throw new ArgumentNullException("method");

            var translatedPropertyAttributes = method.GetCustomAttributes(typeof(TranslatedProperty), true).Cast<TranslatedProperty>();
            if (translatedPropertyAttributes.Count() > 1)
                throw new ArgumentException("Method " + method.Name + " has multiple TranslatedProperty attributes - invalid");
            return translatedPropertyAttributes.SingleOrDefault();
        }

        [ComVisible(true)]
        private class TranslatedPropertyInfo : PropertyInfo
        {
            private readonly string _name;
            private readonly Type _owningType;
            private readonly MethodInfo _getter, _setter;
            public TranslatedPropertyInfo(string name, Type owningType, MethodInfo getter, MethodInfo setter)
            {
                if (name == null)
                    throw new ArgumentNullException("name");
                if (owningType == null)
                    throw new ArgumentNullException("owningType");
                if ((getter == null) && (setter == null))
                    throw new ArgumentException("At least one of getter and setter must be non-null");

                ParameterInfo[] setterParameters;
                if (setter == null)
                    setterParameters = null;
                else
                {
                    setterParameters = setter.GetParameters();
                    if (!setterParameters.Any())
                        throw new ArgumentException("The setter (if non-null) must have at least one parameter)");
                }
                if ((getter != null) && (getter.ReturnType == typeof(void)))
                    throw new ArgumentException("The getter (if non-null) must not have a return type of void");
                if ((getter != null) && (setter != null))
                {
                    var getterParameters = getter.GetParameters();
                    if (getterParameters.Length != (setterParameters.Length - 1))
                        throw new ArgumentException("Where both getter and setter are non-null, the setter must have one more parameter than the gettter");
                    if (setterParameters.Last().ParameterType != getter.ReturnType)
                        throw new ArgumentException("Where both getter and setter are non-null, the setter's last parameter's type must match the getter's return type");
                    for (var index = 0; index < getterParameters.Length; index++)
                    {
                        if (getterParameters[index].ParameterType != setterParameters[index].ParameterType)
                            throw new ArgumentException("Where both getter and setter are non-null, the parameter types must be consistent");
                    }
                }

                _name = name;
                _owningType = owningType;
                _getter = getter;
                _setter = setter;
            }

            public override PropertyAttributes Attributes { get { return PropertyAttributes.None; } }

            public override bool CanRead
            {
                get { return _getter != null; }
            }

            public override bool CanWrite
            {
                get { return _setter != null; }
            }

            public override MethodInfo[] GetAccessors(bool nonPublic)
            {
                return new[] { _getter, _setter }.Where(m => m != null).ToArray();
            }

            public override MethodInfo GetGetMethod(bool nonPublic)
            {
                return _getter;
            }

            public override ParameterInfo[] GetIndexParameters()
            {
                return _getter.GetParameters();
            }

            public override MethodInfo GetSetMethod(bool nonPublic)
            {
                return _setter;
            }

            public override object GetValue(object obj, BindingFlags invokeAttr, Binder binder, object[] index, CultureInfo culture)
            {
                return _getter.Invoke(obj, invokeAttr | BindingFlags.InvokeMethod, binder, index, culture);
            }

            public override Type PropertyType
            {
                get
                {
                    if (_getter != null)
                        return _getter.ReturnType;
                    return _setter.GetParameters().Last().ParameterType;
                }
            }

            public override void SetValue(object obj, object value, BindingFlags invokeAttr, Binder binder, object[] index, CultureInfo culture)
            {
                _setter.Invoke(obj, invokeAttr | BindingFlags.InvokeMethod, binder, index.Concat(new[] { value }).ToArray(), culture);
            }

            public override Type DeclaringType
            {
                get { return _owningType; }
            }

            public override object[] GetCustomAttributes(Type attributeType, bool inherit)
            {
                return new object[0];
            }

            public override object[] GetCustomAttributes(bool inherit)
            {
                return new object[0];
            }

            public override bool IsDefined(Type attributeType, bool inherit)
            {
                return false;
            }

            public override string Name
            {
                get { return _name; }
            }

            public override Type ReflectedType
            {
                get { return _owningType; }
            }
        }
	}
}