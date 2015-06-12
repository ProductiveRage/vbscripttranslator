using System;
using System.Reflection;

namespace CSharpSupport.Implementations
{
    public static class MissingMemberException_Extensions
    {
        public static bool RelatesTo(this MissingMemberException source, Type type, string memberNameIfAny)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (type == null)
                throw new ArgumentNullException("type");

            // If a default member is requested, then a number of things may happen. If the request comes from VBScript then it will likely be requested as "[DISPID=0]",
            // in which case that string will appear in the exception message. If a request is made through an IReflect.InvokeMember call then the member may appear blank.
            // If a request is made through a Type.InvokeMember call then the blank string may be replaced with the member identified by the DefaultMemberAttribute that
            // the type has (if it has one) - eg. typeof(string) will specify "Chars" as the target member (since that is what the DefaultMemberAttribute specifies).
            // - So first, try the simplest match case, where there is no funny business
            if (source.Message.Contains("'" + type.FullName + "." + memberNameIfAny + "'"))
                return true;
            
            // If that doesn't succeed, and it looks like the request was for the default member, then try the various default member options
            if (string.IsNullOrWhiteSpace(memberNameIfAny) || (memberNameIfAny == "[DISPID=0]"))
            {
                var defaultMemberNameOfTargetType = type.GetCustomAttribute<DefaultMemberAttribute>(inherit: true);
                if (defaultMemberNameOfTargetType != null)
                {
                    // TODO: I don't even know if this is correct any more
                    return
                        source.Message.Contains("'" + type.FullName + "." + defaultMemberNameOfTargetType.MemberName + "'") ||
                        source.Message.Contains("'" + type.FullName + ".[DISPID=0]'") ||
                        source.Message.Contains("'" + type.FullName + ".'");
                }
            }
            return false;
        }
    }
}
