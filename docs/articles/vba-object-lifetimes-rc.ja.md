# VBA のオブジェクトの寿命を Rust の `Rc<()>` で再現する

どこの会社にも、何年も誰も中を見ていない Excel マクロのフォルダがあるものです。移行するにせよ廃止するにせよ、そもそも「コンテンツの有効化」を押してよいのか判断するにせよ、まずはそのマクロが何をするのかを知らなければなりません。そのために作ったのが `oxivba-core` です。VBA の字句解析・構文解析・静的解析・インタプリタを、依存クレートなしの Rust で書いたもので、CLI でもブラウザ（WebAssembly）でも同じように動きます。この記事では、その中で「簡単だろう」と思っていたのに一番手こずった部分、VBA のオブジェクトがいつ消えるかの扱いを紹介します。

## VBA ではオブジェクトの寿命がプログラムから見える

VBA は COM の上に成り立っていて、COM のオブジェクトは参照カウントで管理されています。厄介なのは、それがプログラムから見えることです。クラスモジュールには `Class_Terminate` を書くことができ、VBA はこれを「そのインスタンスを指す最後の参照がなくなった瞬間」に実行します。実際のマクロはこのタイミングを当てにしています。ログファイルを閉じる、`Application.ScreenUpdating` を元に戻す、手続きを抜けるときに「完了」と表示する、といった使い方です。

```vb
' クラスモジュール: Guard
Private Sub Class_Terminate()
    Debug.Print "released"
End Sub

' 標準モジュール
Sub Work()
    Dim g As Guard
    Set g = New Guard
    Debug.Print "working"
End Sub          ' "working" の直後、まさにここで "released" が出る
```

一般的なガベージコレクタでは、"released" が出るのは「そのうち」です。実際のマクロをそのとおりに動かすには、参照が 0 になる命令をぴったり特定しなければなりません。`Set x = Nothing` のような分かりやすい場合だけではありません。次のような場合にも、オブジェクトは手放されます。

- それを持っていた唯一の変数に、別の値を代入したとき
- オブジェクトの配列に `Erase` や `ReDim` を実行したとき
- 最後にそれを持っていた `Collection` が手放されたとき
- 実行時エラーで手続きを途中で抜けたとき

しかも `Class_Terminate` は、どの場合でも必ず 1 回だけ実行しなければなりません。

## 参照ごとに「寿命の持ち分」を持たせる

クラスモジュールのインスタンスは、実行時のオブジェクト表に数値のハンドルで登録されています。VBA の変数がそれを指すときは、`ObjectRef` を持ちます。

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

この `Rc<()>` は中身を持たない、ただのカウンタです。参照がコピーされるたびに、つまりローカル変数、配列の要素、別のオブジェクトのフィールド、`Collection` の項目のどこに入る場合でも `ObjectRef` が clone され、持ち分が 1 つ増えます。コピーが捨てられれば持ち分も減ります。COM が `AddRef` と `Release` で行っている帳簿付けを、Rust の所有権がそのまま肩代わりしてくれるわけです。増減を自分で書く箇所がないので、エラー処理の経路で書き忘れることもありません。

さらに、実行時は生きているインスタンスごとに自分用の持ち分を 1 つ、登録簿に持っています。そのため「もう誰も持っていない」かどうかは、比較 1 回で分かります。

```rust
Rc::strong_count(&instance.life) == 1   // 残っているのは登録簿の持ち分だけ
```

## 回収するタイミング

カウントがいつ減るかは Rust が決めます。こちらが決めるのは、いつ確認するかです。インタプリタは、参照が手放される可能性のある文の後で確認します。代入、`Set`、`Erase`、`ReDim`、戻り値を使わない呼び出し、そして手続きの終了（エラーによる終了も含む）です。回収のループを簡略化すると、次のようになります。

```rust
fn collect_instances(&mut self, line: u32) -> Result<(), RuntimeError> {
    loop {
        // 誰も持たなくなったコンテナを消す。中身の持ち分も一緒に返る
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
            continue; // コンテナが消えたことで、新たに解放されたインスタンスがあるかもしれない
        }
        for (handle, class) in due {
            if !self.mark_terminated(handle) { continue; } // 2 回は実行しない
            if self.class_has_terminate(&class) {
                // Class_Terminate の Err は呼び出し側に持ち込まない
                let saved = (self.err_in.take(), self.err_out.take());
                let ran = self.run_terminate(handle, &class, line);
                (self.err_in, self.err_out) = saved;
                ran?;
            }
            self.internal_objects.remove(&handle); // これだけが持っていたものも消える
        }
    }
}
```

このループには工夫が 3 つありますが、どれも一度は間違えて直したところです。

- **ループにしてある。** オブジェクトを消すと、そのフィールドに入っていた `ObjectRef` も捨てられるので、別のインスタンスの持ち分が 1 になることがあります。`Set col = Nothing` で `Collection` を手放せば、中の項目も同じタイミングで道連れになります。
- **終了済みは、消すのではなく印を付ける。** `Class_Terminate` の中でさらにオブジェクトが手放されることがあり、その直後の回収が、今まさに処理待ちのインスタンスに届くことがあります。印を付けておけば、終了処理が 2 回走ることはありません。
- **呼び出しの前後で `Err` を退避する。** 手続きがエラー 5 を出し、抜ける途中でローカル変数のオブジェクトが終了した場合でも、呼び出し側から見えるのはエラー 5 のままです。終了処理の中の `On Error` の状態が外に漏れてはいけません。

これらはドキュメントを読んだ私の解釈ではなく、すべて本物の Excel で確かめた挙動です。実行時のソースには「measured」で始まるコメントが 200 個以上あり、どれも Rust のコードを書く前に、その場合 Excel がどう答えたかを記録したものです。ドキュメントと Excel の挙動が食い違ったら、Excel に合わせます。

## 何を渡されても落ちないこと

他人の書いたマクロを動かす以上、想定外の値はいくらでも渡されます。VBA では、どの組み込み関数にも `Null`、`Empty`、数百文字もある文字列、`Decimal` の最大値、`#12/31/9999#`、`Nothing` を渡せてしまいます。返すべきは値か、「プロシージャの呼び出し、または引数が不正です」のような VBA の実行時エラーです。ホスト側のページまで巻き込んで落ちる Rust の panic であってはいけません。

そこで、これだけを確かめるテストを 1 本用意しています。実装済みの組み込み関数すべて（`Abs` から `Year` まで 145 個）を、引数なしで、また 24 種類の扱いにくい値をそれぞれ 1〜4 個並べて呼び出します。すべて `On Error Resume Next` の下で実行し、panic が 1 つもないことを確かめます。

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

手間はほとんどかかりません。ファジングの仕組みを用意する必要もなく、普段の `cargo test` で走ります。失敗すれば、どの関数にどの引数を渡したときかがそのまま出ます。実行ステップ数の上限 `max_steps` も同じ考え方で、無限ループするマクロはタブを固まらせず、エラーとして終わります。

## 実行する前に読む

同じ構文木は、何も実行しない用途にも使っています。`oxivba_core::assess` は、Excel が黄色いマクロの警告バーを出したときに誰もが迷う「これを有効にして大丈夫か」という問いに答えるための関数です。

ただし、あえて危険かどうかの判定は下しません。`Shell` は正当なマクロが PDF を開くのにも使いますし、`MSXML2.XMLHTTP` は為替レートの表を取ってくるのにも使います。どちらにも「危険」と表示するツールは、警告を読み飛ばす癖を付けさせるだけです。そこで、コードが何にアクセスできるかを、行番号付きの根拠として並べるだけにしています。はっきり言い切るのは、誰も何も押さなくても動くコード（`Workbook_Open` や `Auto_Open` など）があるかどうかの 1 点だけです。それがあるなら、問題は「実行してよいか」ではなく「もう実行されてしまったか」になるからです。

分からなかったことも、分からなかったと報告します。`CreateObject(name)` の引数が実行時に決まる場合は、推測せずに「特定できない」と返します。構文解析できなかった行は、その数を報告します。誰も読めていない行は、誰も安全を確かめていない行だからです。構文解析器も入力を黙って捨てることはせず、理解できなかった部分はそのまま `Unknown` ノードとして残します。

## ブラウザで動かす

ここまでのコードはファイルの読み書きもホスト固有の型も一切使っていないので、そのまま `wasm32` 向けにコンパイルできます。Excel のオブジェクトモデル（`Range`、`Worksheet`、`Workbook`）はインタプリタには含めず、ホストがトレイト経由で提供します。

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

Oxi のブラウザ版エディタでは、スプレッドシートのエンジンがこの `Host` を実装しています。マクロは WebAssembly 経由で Web Worker の中で動くので、時間のかかるマクロでもページが固まることはありません。同じインタプリタを、スプレッドシートを持たない CLI 側のホストから動かし、`.bas` ファイルの入ったフォルダをまとめて解析することもできます。

## 同じようなインタプリタを作る人へ

1. 言語がオブジェクトの寿命をプログラムに見せるなら、数えるのは Rust の所有権に任せ、自分で書くのは「いつ確認するか」だけにする。参照ごとに `Rc<()>` を持たせる方法は、参照カウントの仕組みを自前で書くよりずっと手軽で、エラー処理の経路でカウントがずれる心配もありません。
2. 元の実装で実際に観測できる挙動を仕様とし、確かめた結果は、それを根拠にしたコードのすぐそばに書き残す。
3. 「何を渡しても panic しない」テストは最初の日に書く。10 行ほどで書けて、誰かが本物のマクロを流し込んだその日に元が取れます。

`oxivba-core` は MPL-2.0 で、DOCX・XLSX・PPTX の各エンジンと一緒に [Oxi のリポジトリ](https://github.com/Ryujiyasu/oxi) で公開しています。ワークブックのマクロは [ブラウザ版エディタ](https://oxi-dd65f4.gitlab.io/ja/) で実際に動かせます。
