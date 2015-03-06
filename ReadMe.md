# The VBScript (to C#) Translator

I think that this is probably one of the stupidest projects I've attempted. And one of the most challenging. And kind of one of the most fun in a perverse way.

Years ago, I wrote a JavaScript parser that could be used to combine and minify JavaScript content on-the-fly. There wasn't really much around at that time that did it "on demand", Google's Closure Compiler was just making an appearance but it intended to be part of a "build step", really - but then it does do a lot of clever things. My parser didn't, it was basic but functional - parse JavaScript, remove unnecessary whitespace, spit it out. At work at the time we were writing a lot of VBScript (Classic ASP) and I wondered how hard it would be parse that instead. I don't really know why, maybe I wanted to write some sort of static analysis tools (ReSharper4VBScript, anyone?? :) Whatever. I did much of the work to parse VBScript into some sort of tree and then put the project away and didn't think about it again.

For several years, at least.

At this point were *still* writing a lot of VBScript. How hard, I wondered, could it be to take this parsed VBScript and twist and turn and manipulate it such that C# came out the other end that was functionally equivalent(ish)? I mean, VBScript is *simple*, right, so how bad could it be??

If I did this then we could translate the old VBScript code into C# and it would run lovely and quick. And the interesting stuff that we'd maintain and keep working on could be rewritten into "native C#" and any old stuff that was still required but rarely touched could remain as "from-VBScript C#".

It's just classes, functions and if-blocks with some dynamic madness thrown in!

## Reality check

Let me cut to the chase: This doesn't work. Yet. It parses nearly all VBScript structures into a format that it can then use to emit C# (as of March 2015 I still haven't finished the parsing of VBScript's "SELECT CASE" statements but it's not far off, plus I ran it against thousands of line of legacy code and there was one particular line that confused it - but it's fixed). And it does indeed emit compilable C# for many of the structures that it can parse.

The other thing is that it only tries to translate raw VBScript files - if it was going to translate ASP files (with mixed markup and script) then some sort of transformation would need to be applied first. If it was going to translate WSCs files (does anyone else even remember these) then it would need to parse the extra meta data in those files and adjust the interface on the compiled class accordingly.

But what it *does* do (rather heroically, thank you very much) is deal with VBScript's ideas around scoping (including scoping of undeclared variables), insanely "helpful" (confusing) calling conventions and error-handling. And it makes a sort of wild stab at approximating its deterministic garbage collection (with a real emphasis on *approximating*).

I've tried to think of one example that can encapsulate as much as possible of the *essence* of this work. And this seems like as good of a place to start as any:

    On Error Resume Next
    Set o = new C1
    a = 1
    o.F1(a)
    If o.F2(a) Then
      Response.Write "Hurrah! (a = " & a & ")<br/>""
    Else
      Response.Write "Ohhhh.. sad face (a = " & a & ")<br/>""
    End If

    Class C1
      Function F1(b)
        Response.Write "b is " & b & " (a = " & a & ")<br/>""
        b = 2
        Response.Write "b is " & b & " (a = " & a & ")<br/>""
      End Function

      Function F2(c)
        Response.Write "c is " & c & " (a = " & a & ")<br/>""
        c = 3
        Response.Write "c is " & c & " (a = " & a & ")<br/>""
        Response.Write "Time to die: " & (1/0)
      End Function
    End Class
    
Before diving in, let's just keep in mind: This is a fun project to sort of work out some masochistic tendencies. I sort of have an idea where I could realistically use this in real life (as a temporary measure) but there's definitely something in me that thinks that it's just madness for madness sake. Surely there's a reason why there was never a go-to solution for this when everyone was going mad about migrating from "Classic" ASP to .net, a decade ago or whenever it was! Here be dragons, basically. You've been warned.

Where to start! The code that is executed is just sort floating in space. That won't do for C#. It will need wrapping up in a class that is instantiated and has some sort of entry method called each the script is executed. That's not such a big deal. But the variable "a" is accessible from the functions "F1" and "F2" within the class "C1" - if "a" is to be accessible by the main "TranslatedProgram" class *and* by the explicit VBScript classes then "a" will have to be static, which means that the entry method must be static. But then state would be shared between concurrent calls to the script! And if this is to translate web scripts then concurrent calls must be considered to be possible. So, instead of a static "TranslatedProgram" class, how about a non-static class that stashes its global variables in a "GlobalReferences" class that it then passed into each translated class - so the C# version of "C1" will have a constructor that takes a "GlobalReferences" instance, so that the functions "F1" and "F2" can monkey around with "a". That sounds ok. Any functions that are not within classes will have to go into this "GlobalReferences" class too, since they must also be accessible from functions within classes, in the same way that "a" is.

Fine. Plain sailing now. Well, obviously there will have to be code in the translated C# that knows to translate "new C1" into an instantiation via a constructor that takes a "GlobalReferences" instance, but that's no big deal. Concurrency *solved*. Now just some simple function calls to deal with.

You may or may not be surprised to know that the first four lines printed out by this script when executed as VBScript are:

    b is 1 (a = 1)
    b is 2 (a = 1)
    c is 1 (a = 1)
    c is 3 (a = 3)

Yes. The call to "F1" does not affect the value of "a" while the call to "F2" *does*. In VBScript, function arguments may be ByRef or ByVal. ByRef meaning that any changes to them within the function are reflected in the caller's value - like using "ref" in C#. And the default is "ByRef". Not what you might expect coming from any other programming language (but then, this is the case with *many* of VBScript's language decisions!).

So that explains the "F2" call - it changes the value of "a" since "a" is passed ByRef as the argument "c" and so the change to "c" means that "a" is changed as well. But what about "F1"? Well! Brackets around the argument set are compulsory when a function is called and its return value considered (as is the case when "F2" is called and its return value would be used as the "If" condition.. if it didn't throw an error). But when the return value is not considered then brackets may *not* be used to wrap an argument set. In fact if "F1" took two arguments and we tried to call it with

    o.F1(a, False)
    
Then we would be treated to a

> VBScript compilation error: Cannot use parenteses when calling a Sub

But the brackets in our first example are not wrapping an argument *set*, they are wrapping a *single argument*. And in VBScript, this has special meaning - it means pass this value "ByVal", even if the function declares the argument as ByRef. And this is why "a" is not changed by the call to "F1".

To continue exploring the parallel universe where "F1" had multiple arguments, the equivalent calling code would be:

    o.F1(a), (False)
    
or

    o.F1 (a), (False)
    
In these cases, both "a" and "False" would be passed ByRef. Which, of course, would not affect "False", which is a constant and can't be changed through a function call. Which highlights a suitable difference between VBScript's ByRef and C#'s ref - in C#, a ref argument *must* be a mutatable value (*not*, for example, a constant value such as False or 42). There are other limitations in C# - you may not pass an int variable as a ref argument if the method signature is "ref object", for example.

But let's push on. There is a fifth line printed by this code when executed as VBScript. Kudos (or maybe commiserations would be more appropriate) to anyone that realises that the fifth line is, in fact

    Hurrah! (a = 3)
    
(If you guessed that this would be the case solely due to a belief that VBScript is intentionally perverse then you don't get the pat on the back.. but it's a fair enough assumption).

VBScript's error handling model is to proceed to the next statement when an error is raised and "trapped". In the case of an "If" statement, the next statement is the first one within the conditional block. This is *not* the same as C#, where the entire "if..else" block would bypassed if it was wrapped in a "try..catch" and the first condition raised an exception. This was explained by Eric Lippert back in 2004 in his article "[Error Handling in VBScript, Part One](http://blogs.msdn.com/b/ericlippert/archive/2004/08/19/error-handling-in-vbscript-part-one.aspx)". Incidentally, [Part 3](http://blogs.msdn.com/b/ericlippert/archive/2004/08/25/error-handling-in-vbscript-part-three.aspx) of that mini-series includes a really insightful comment (bear in mind that this someone who was deeply involved in writing the VBScript Interpreter and not just some random hater):

> I'm working in C# and C++, languages specifically designed for implementing complex software written by large teams. VBScript is not such a language -- it was designed for simple administration and web scripts, where often "muddle on through" is exactly what you want it to do.

A lot of the decisions made regarding how VBScript should work may seem insane to a C# Developer (include me in this group!) but they (nearly) all have some sort of grounding in this logic - VBScript did not have the same people in mind that other languages do.

## Tip of the iceberg

That was a fairly simple example. It's not hard to find many *many* more cases like it. There are other If evaluation oddities such as

    a = 1
    b = "1"
    If (1 = "1") Then Response.Write "Condition 1 is True<br/>" ' True
    If (a = "1") Then Response.Write "Condition 2 is True<br/>" ' True
    If (1 = b) Then Response.Write "Condition 3 is True<br/>" ' True
    If (a = b) Then Response.Write "Condition 4 is True<br/>" ' False

This is due to special behaviour around literals in conditions - they force the other side of the comparison to be parsed into their type, giving priority to numbers over strings. So the first condition forces "1" into a number and so it's a match. The second condition forces the value of "a" into a string and so it's a match. The third condition forces the value of "b" into a number and so it's a match. The fourth condition has no literals and so no type coercion and so no match.

There are For loop oddities where the loop will be entered once if the evaulation of any of the loop constraints throws an error and On Error Resume Next is in play. But it won't set the loop variable, it will just plow throught the loop's statements once and then bail out. It will always set the loop variable to be a type that can contain the all of the loop variable's values, for

    For i = 1 To 5
      Response.Write TypeName(i) & "<br/>"
    Next
    
always prints "Integer" while

    For i = 1 To 100000
      Response.Write TypeName(i) & "<br/>"
    Next
    
always prints "Long" (since "Integer" would overflow after 32,767 and so a "Long" is required to contain all of the values).

Oh, unless one of the loop constraints evaluates to a Decimal or Date - in this case, an overflow will occur at runtime if one of the other constraints is evaulated to a number that is too big for these data types. If this occurs while an On Error Resume Next has control of matters then the loop will still be entered only once but the loop variable *may* be set. It's all fun and games.

On top of these sorts of things, there's the fact that VBScript is case-insensitive. *And* the fact any binding issues (such as a function call with the wrong number of arguments) will result in *runtime* errors, not compile time. And so C# representing that code must compile (even with the wrong number of arguments for function calls) but throw exceptions when executed. Unless that code has error trapping enabled with an On Error Resume Next - that case, the code must be allowed to "muddle on through".

In summary, this translation is a stupid idea. But all of these crazy cases have been what made it so fun! I came to realise that I didn't actually know VBScript as well as I'd thought I did. And I had to dig deep into C# in some cases, to try to find a solution that relied on as little insanity as feasibly possible. Plus I was able to apply various other techniques I've stumbled across in the last few years - need to access COM objects through IDispatch from C# when I don't even know at compile time if I'll be talking to COM or a C# class translated from some VBScript code? No problem. Want to cache function calls, compiling them into LINQ expressions, so that subsequent calls to the method "F1" on an instance of "C1" where one object argument is specified can be less expensive? Why not!

Let's not forget that at the bottom of the pile is code that I wrote many years ago. It is.. something of an experience to really work with code you wrote that long ago. Not just look at with a sort of nostalgic glint in your eye, but to actually extend and refactor. It's not that pretty I must admit. But building on top of it and slowly improving it (or resisting the urge in cases where it works well enough and I'd be better off spending time elsewhere) has been part of the challenge! And the test 
coverage slowly increases.. For quite some time now I've been trying to include tests that fail without the accompanying changes and that pass *with* them. It's not quite red, green, refactor but it's been a really positive step towards preventing regressions (and they have in fact done precisely this several times - highlighted code that my "fix" would break where I hadn't accounted for something). And a lot of the tests exercise too much code to be strictly *unit* tests, but they do the job nonetheless! (And it's not like real world integration tests where anything messy like a database needs to be involved).

## Is there an end in sight?

Even if this project never serves any goal but my own satisfaction at doing something so perverse (and the mental exercise in doing so).. well I see that as no reason not to keep working on it. I don't think it's going to be long until it can parse all valid VBScript. And it can already perform a wide range of transformations. It can't handle everything yet and there are some edges cases that I'm aware of that I know I don't deal with yet. And I suspect that there are more edge cases that I am yet to uncover. But it *is* capable of translating some large, real-world VBScript files.

Running the translated code is a different matter, though, since the "support library" is woefully incomplete. I need to implement all of VBScript's methods. Take "Round", for example; in VBScript you don't just give the "Round" function a number - on no, sometimes you give it an object reference that has a default property which is a string that can be parsed into a number and *that* is rounded into an integer. Nothing is ever easy when you're trying to emulate a language which bends over backwards to try to make it easy for the Developer.. so long as that Developer doesn't try to think too hard about what gymnastics the intepreter may be doing behind the scenes.