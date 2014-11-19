using CSharpWriter.Lists;
using System;
using VBScriptTranslator.LegacyParser.CodeBlocks;
using VBScriptTranslator.LegacyParser.CodeBlocks.Basic;

namespace CSharpWriter.CodeTranslation.Extensions
{
    public static class ScopeAccessInformation_Error_Extensions
    {
        public static ScopeAccessInformation SetErrorRegistrationToken(this ScopeAccessInformation scopeAccessInformation, CSharpName errorRegistrationTokenIfAny)
        {
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");

            return new ScopeAccessInformation(
                scopeAccessInformation.Parent,
                scopeAccessInformation.ScopeDefiningParent,
                scopeAccessInformation.ParentReturnValueNameIfAny,
                errorRegistrationTokenIfAny,
                scopeAccessInformation.DirectedWithReferenceIfAny,
                scopeAccessInformation.ExternalDependencies,
                scopeAccessInformation.Classes,
                scopeAccessInformation.Functions,
                scopeAccessInformation.Properties,
                scopeAccessInformation.Variables,
                scopeAccessInformation.StructureExitPoints
            );
        }

        /// <summary>
        /// TODO:
        /// Note that this does not analyse the type of code block, other than looking through the nested code blocks if the IHaveNestedContent or ILoopOverNestedContent
        /// interfaces or implemented. Whether error handling applies to the particular code block is left to the caller to determine, this only checks whether it MAY.
        /// For example, a static DIM statement does not need error handling, nor does an OPTION EXPLICIT, ON ERROR RESUME NEXT or ON ERROR GOTO 0. How the error
        /// handling is implemented is also left to the caller - it will be different for an IF statement than a function assignment, for example.
        /// </summary>
        public static bool MayRequireErrorWrapping(this ScopeAccessInformation scopeAccessInformation, ICodeBlock codeBlock)
        {
            if (scopeAccessInformation == null)
                throw new ArgumentNullException("scopeAccessInformation");
            if (codeBlock == null)
                throw new ArgumentNullException("codeBlock");

            // If the scopeAccessInformation is not configured to allow error-trapping then always return false
            if (scopeAccessInformation.ErrorRegistrationTokenIfAny == null)
                return false;

            // If there's no OnErrorResumeNext present at all within the scope-defining parent's code blocks, then don't both with any of the work below,
            // it isn't necessary (scope-defining blocks are the outermost scope, classes and functions / properties - so if error trapping is enabled
            // outside of the current scope then they don't affect what happens within this scope).
            var codeBlocksWithinScopeDefiningParent = scopeAccessInformation.ScopeDefiningParent.AllExecutableBlocks.ToNonNullImmutableList();
            if (!codeBlocksWithinScopeDefiningParent.DoesScopeContainOnErrorResumeNext())
                return false;

            // Locate where the specified code block exists within its scope
            var codeBlockWithinContent = TryToLocateCodeBlock(codeBlock, codeBlocksWithinScopeDefiningParent, new NonNullImmutableList<IHaveNestedContent>());
            if (codeBlockWithinContent == null)
                throw new ArgumentException("The codeBlock does not exist within the ScopeDefiningParent of the specified scopeAccessInformation");

            // Now try to determine whether the code in the current scope means that error-trapping should be enabled for the specified code block. It
            // might identify false positives in some scenarios but should never identify false negatives.
            var currentCodeBlockInTrackBack = codeBlockWithinContent.CodeBlock;
            var remainingParentsToTrackBackThrough = codeBlockWithinContent.ParentsWithinScope.Add(scopeAccessInformation.ScopeDefiningParent);
            while (remainingParentsToTrackBackThrough.Any())
            {
                var parent = remainingParentsToTrackBackThrough.Last();
                
                // If the parent loops and there is an OnErrorResumeNext ANYWHERE within it, then there's a chance that error handling will be required
                // (there may be conditions that never get met but it would require much deeper analysis to determine this than this implementation will
                // consider)
                if ((parent is ILoopOverNestedContent) && parent.AllExecutableBlocks.DoesScopeContainOnErrorResumeNext())
                    return true;

                // If any of the other code blocks within the same parent that appear before the current block in an OnErrorResumeNext or contains an
                // OnErrorResumeNext then we have to assume that the target codeBlock may be affected by error-trapping. Again, there may be conditions
                // that mean that the OnErrorResumeNext statement does not apply but the analysis required to determine this are out of scope of what I
                // have in mind here. With the exception of the most simple case where a non-nested OnErrorGoto0 appears between where error-trapping is
                // enabled and where the original code block was identified - explicitly disabling the error-trapping. This addresses the common forms
                // of OnErrorResumeNext / One-Statement / OnErrorGoto0 or OnErrorResumeNext / One-Statement / If (Err) Then Statement / OnErrorGoto0
                var errorTrappingMustBeEnabled = false;
                foreach (var siblingCodeBlock in parent.AllExecutableBlocks)
                {
                    if (siblingCodeBlock == codeBlock)
                        break;

                    if (siblingCodeBlock is OnErrorResumeNext)
                    {
                        errorTrappingMustBeEnabled = true;
                        continue;
                    }
                    else if (siblingCodeBlock is OnErrorGoto0)
                    {
                        errorTrappingMustBeEnabled = false;
                        continue;
                    }

                    var siblingCodeBlockAsNestedContent = siblingCodeBlock as IHaveNestedContent;
                    if ((siblingCodeBlockAsNestedContent != null) && siblingCodeBlockAsNestedContent.AllExecutableBlocks.DoesScopeContainOnErrorResumeNext())
                    {
                        errorTrappingMustBeEnabled = true;
                        continue;
                    }
                }
                if (errorTrappingMustBeEnabled)
                    return true;

                // We can't be sure of a negative result since there could be a parent that contains an OnErrorResumeNext, so all the parents must
                // be analysed before returning
                currentCodeBlockInTrackBack = parent;
                remainingParentsToTrackBackThrough = remainingParentsToTrackBackThrough.RemoveLast();
            }

            // If we haven't positively identified a case where error-trapping must be enabled, then we can declare it not required!
            return false;
        }

        private static CodeBlockWithinScopeDefiningParent TryToLocateCodeBlock(
            ICodeBlock codeBlock,
            NonNullImmutableList<ICodeBlock> potentialAncesterCodeBlocks,
            NonNullImmutableList<IHaveNestedContent> parentCodeBlocksAroundPotentialAncestors)
        {
            if (codeBlock == null)
                throw new ArgumentNullException("codeBlock");
            if (potentialAncesterCodeBlocks == null)
                throw new ArgumentNullException("potentialAncesterCodeBlocks");
            if (parentCodeBlocksAroundPotentialAncestors == null)
                throw new ArgumentNullException("parentCodeBlocksAroundPotentialAncestors");

            foreach (var potentialAncesterCodeBlock in potentialAncesterCodeBlocks)
            {
                if (potentialAncesterCodeBlock == codeBlock)
                    return new CodeBlockWithinScopeDefiningParent(codeBlock, parentCodeBlocksAroundPotentialAncestors);

                var potentialAncesterNestedContentCodeBlock = potentialAncesterCodeBlock as IHaveNestedContent;
                if (potentialAncesterNestedContentCodeBlock == null)
                    continue;

                var nestedMatch = TryToLocateCodeBlock(
                    codeBlock,
                    potentialAncesterNestedContentCodeBlock.AllExecutableBlocks.ToNonNullImmutableList(),
                    parentCodeBlocksAroundPotentialAncestors.Add(potentialAncesterNestedContentCodeBlock)
                );
                if (nestedMatch != null)
                    return nestedMatch;
            }
            return null;
        }

        private class CodeBlockWithinScopeDefiningParent
        {
            public CodeBlockWithinScopeDefiningParent(ICodeBlock codeBlock, NonNullImmutableList<IHaveNestedContent> parentsWithinScope)
            {
                if (codeBlock == null)
                    throw new ArgumentNullException("codeBlock");
                if (parentsWithinScope == null)
                    throw new ArgumentNullException("parentsWithinScope");

                CodeBlock = codeBlock;
                ParentsWithinScope = parentsWithinScope;
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public ICodeBlock CodeBlock { get; private set; }
            
            /// <summary>
            /// This will never be null (though it may be empty)
            /// </summary>
            public NonNullImmutableList<IHaveNestedContent> ParentsWithinScope { get; private set; }
        }
    }
}