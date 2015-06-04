/* Undeclared variable: "i" (line 3) */
using System;
using System.Collections;
using System.Runtime.InteropServices;
using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpSupport.Exceptions;

namespace TranslatedProgram
{
    public class Runner
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        public Runner(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer)
        {
            if (compatLayer == null)
                throw new ArgumentNullException("compatLayer");
            _ = compatLayer;
        }

        public void Go()
        {
            Go(new EnvironmentReferences());
        }
        public void Go(EnvironmentReferences env)
        {
            if (env == null)
                throw new ArgumentNullException("env");

            var _env = env;
            var _outer = new GlobalReferences(_, _env);

            // Test
            for (_outer.i = (Int16)1; _.StrictLTE(_outer.i, 10); _outer.i = _.ADD(_outer.i, (Int16)1))
            {
                _.CALL(_env.wscript, "Echo", _.ARGS.Val(_.CONCAT("Item", _outer.i)));
            }
        }

        public class GlobalReferences
        {
            private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
            private readonly GlobalReferences _outer;
            private readonly EnvironmentReferences _env;
            public GlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env)
            {
                if (compatLayer == null)
                    throw new ArgumentNullException("compatLayer");
                if (env == null)
                    throw new ArgumentNullException("env");
                _ = compatLayer;
                _env = env;
                _outer = this;
                i = null;
            }

            public object i { get; set; }
        }

        public class EnvironmentReferences
        {
            public object wscript { get; set; }
            public object i { get; set; }
        }
    }
}
