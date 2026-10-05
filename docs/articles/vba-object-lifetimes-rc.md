# Reproducing VBA's object lifetimes in Rust with `Rc<()>`

Somewhere in most companies there is a folder of Excel macros nobody has read in years. Before you can migrate them, retire them, or even decide whether it is safe to click "Enable Content", you need to know what they do. Oxi's answer is `oxivba-core`: a VBA lexer, parser, static analyser and interpreter written in Rust with zero dependencies, so that it runs the same way in a CLI and in a browser tab through WebAssembly. This post is about the part of that interpreter I expected to be easy and was not: deciding when a VBA object dies.

## Why lifetimes matter in VBA

VBA sits on COM, and COM objects are reference counted. That is visible to the program. A class module can declare a `Class_Terminate` procedure, and VBA runs it at a precise moment: when the last reference to the instance goes away. Code in the wild depends on that moment. People use it to close a log file, to restore `Application.ScreenUpdating`, and to print "done" when a helper object falls out of scope at the end of a procedure.

```vb
' Class module: Guard
Private Sub Class_Terminate()
    Debug.Print "released"
End Sub

' Standard module
Sub Work()
    Dim g As Guard
    Set g = New Guard
    Debug.Print "working"
End Sub          ' prints "working", then "released" here, not later
```

A tracing garbage collector would print "released" at some unspecified later time. To run real macros faithfully, an interpreter has to know the exact instruction at which the count reaches zero. That includes the less obvious cases. `Set x = Nothing` releases the instance. So does reassigning the only variable that held it, `Erase` or `ReDim` on an array of objects, a `Collection` that was the last holder being released, and a procedure left early by a runtime error. And it has to run `Class_Terminate` exactly once.

## A share of life

Instances of class modules live in the runtime's object table, addressed by a numeric handle. A VBA variable that holds one holds an `ObjectRef`:

```rust
#[derive(Debug, Clone)]
pub struct ObjectRef {
    pub handle: u64,
    pub kind: String,
    /// For an instance of a class module, a share in its life: the runtime
    /// counts these to know when the last reference is gone and
    /// Class_Terminate is due. None for everything else.
    pub life: Option<Rc<()>>,
}
```

The `Rc<()>` carries no data. It is only a counter. Every copy of the reference, whether it is a local variable, an array element, a field of another object, or an item inside a `Collection`, clones the `ObjectRef` and so takes one share. Dropping the copy gives the share back. Rust's ordinary ownership does the bookkeeping that COM does with `AddRef` and `Release`. I never write an increment or a decrement by hand, so I cannot forget one on an error path.

The runtime keeps one share of its own for every live instance, in its registry. That makes the test for "nobody holds this any more" a single comparison:

```rust
Rc::strong_count(&instance.life) == 1   // only the registry's own share is left
```

## Collecting at the right moment

When the count drops is decided by Rust. When the runtime *looks* is decided by VBA semantics. The interpreter checks after statements that can release references: assignments, `Set`, `Erase`, `ReDim`, a call statement whose return value is thrown away, and procedure exit, including exit by error. A simplified version of the collection loop looks like this:

```rust
fn collect_instances(&mut self, line: u32) -> Result<(), RuntimeError> {
    loop {
        // Containers nothing holds any more are gone, and what they held
        // gives back its shares with them.
        let dropped: Vec<u64> = self.container_lives.iter()
            .filter(|(_, life)| Rc::strong_count(life) == 1)
            .map(|(handle, _)| *handle)
            .collect();
        for handle in &dropped {
            self.container_lives.remove(handle);
            self.internal_objects.remove(handle);
        }

        let due = self.instances_with_only_the_registry_share();
        if due.is_empty() {
            if dropped.is_empty() { return Ok(()); }
            continue; // a dropped container may have freed instances
        }
        for (handle, class) in due {
            if !self.mark_terminated(handle) { continue; } // never twice
            if self.class_has_terminate(&class) {
                // Class_Terminate has its own Err; the caller's survives it.
                let saved = (self.err_in.take(), self.err_out.take());
                let ran = self.run_terminate(handle, &class, line);
                (self.err_in, self.err_out) = saved;
                ran?;
            }
            self.internal_objects.remove(&handle); // and what it alone held
        }
    }
}
```

Three details in that loop each came from a case that went wrong first:

- **It is a loop.** Removing an object drops the `ObjectRef`s stored in its fields, and that can bring other instances down to one share. A `Collection` released by `Set col = Nothing` takes its items with it in the same step.
- **Terminated is a flag, not a removal.** A `Class_Terminate` body may itself release objects, and the collection that runs on its return may reach an instance this pass already queued. The flag makes sure each instance terminates once.
- **`Err` is saved around the call.** Suppose a procedure raises error 5 and, on the way out, one of its locals is terminated. The caller still sees error 5. The terminate handler's own `On Error` state must not leak into it.

Each of those is checked against Excel itself rather than against my reading of the documentation. The runtime has a couple of hundred comments that start with "measured". Each records what real Excel answered for a case before the Rust code was written to match it. When the documentation and Excel disagree, Excel wins.

## An interpreter that must not fall over

A runtime that reads other people's macros will be handed things nobody intended. A VBA program can pass `Null`, `Empty`, a string several hundred characters long, the largest `Decimal`, `#12/31/9999#` or `Nothing` to any built-in function. The answer may be a value or a VBA runtime error such as "Invalid procedure call". It must never be a Rust panic that takes the host page down with it.

So there is one test that does nothing else. It calls every built-in function the runtime implements (145 of them, from `Abs` to `Year`) with no arguments and with one to four copies of each of 24 awkward values. Each call runs under `On Error Resume Next`, and the test asserts that nothing panics:

```rust
for name in NAMES {
    for kind in KINDS {
        let module = parse_module(&probe_source(name, kind))?;
        let caught = std::panic::catch_unwind(AssertUnwindSafe(|| {
            let mut runtime = Runtime::new(&module);
            runtime.max_steps = 200_000;
            let _ = runtime.call("Probe", vec![]);
        }));
        if caught.is_err() { fell.push(format!("{name}({kind})")); }
    }
}
assert!(fell.is_empty(), "panicked: {}", fell.join(" "));
```

It is cheap. It needs no fuzzing infrastructure and runs in an ordinary `cargo test`. When it fails, it names the exact function and the exact argument. The `max_steps` budget is the other half of the same promise: a macro with an infinite loop ends with an error, not a frozen tab.

## Reading before running

The same syntax tree serves a second purpose that never executes anything. `oxivba_core::assess` answers the question a person faces when Excel shows the yellow macro bar: *is it safe to enable this?*

It deliberately gives no verdict. `Shell` is how a legitimate macro opens a PDF. `MSXML2.XMLHTTP` is how one fetches an exchange-rate table. A tool that labels both "dangerous" teaches people to click through warnings. So the report lists *capabilities*: what the code can reach, with line numbers. The one thing it states plainly is whether code runs without anyone pressing anything (`Workbook_Open`, `Auto_Open` and friends). That changes the question from "should I run this?" to "have I already run it?".

It also says what it could not see. A `CreateObject(name)` whose argument is computed is reported as unresolved rather than guessed at. Lines the parser could not read are counted and reported, because a line nobody could read is a line nobody has cleared. The parser never drops input silently; anything it does not understand is kept verbatim as an `Unknown` node.

## In the browser

Everything above has no I/O and no host types, which is why it compiles to `wasm32` unchanged. The Excel object model (`Range`, `Worksheet`, `Workbook`) is not part of the interpreter. A host supplies it through a trait:

```rust
pub trait Host {
    fn call(
        &mut self,
        receiver: Option<&ObjectRef>,
        name: &str,
        args: &[Value],
    ) -> Result<Option<Value>, String>;
    // ...
}
```

In Oxi's browser editor, the spreadsheet engine implements `Host`. The macro runs in a Web Worker through the WebAssembly bindings, so a long-running macro does not block the page. The same interpreter can be driven by a CLI host that has no spreadsheet at all, for analysing a folder of `.bas` files.

## What I would tell someone starting a similar interpreter

1. If the language exposes object lifetimes, let Rust's ownership do the counting and keep your own code to deciding *when to look*. An `Rc<()>` share in every reference cost far less than an explicit reference-counting layer, and it cannot drift out of sync on an error path.
2. Treat the original implementation's observable behaviour as the specification, and write down each measurement next to the code it justifies.
3. Write the "nothing panics" test on the first day. It is ten lines and it pays for itself the first time someone feeds the interpreter a real macro.

`oxivba-core` lives in the [Oxi repository](https://github.com/Ryujiyasu/oxi) under MPL-2.0, alongside the DOCX, XLSX and PPTX engines. You can run workbook macros in the [browser editor](https://oxi-dd65f4.gitlab.io/).
