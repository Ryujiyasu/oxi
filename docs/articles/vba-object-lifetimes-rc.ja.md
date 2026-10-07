# VBA のオブジェクトの寿命を Rust の `Drop` で再現する

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

一般的なガベージコレクタでは、"released" が出るのは「そのうち」です。実際のマクロをそのとおりに動かすには、最後の参照が消える文をぴったり特定しなければなりません。`Set x = Nothing` のような分かりやすい場合だけではありません。次のような場合にも、オブジェクトは手放されます。

- それを持っていた唯一の変数に、別の値を代入したとき
- オブジェクトの配列に `Erase` や `ReDim` を実行したとき
- 最後にそれを持っていた `Collection` が手放されたとき
- 実行時エラーで手続きを途中で抜けたとき

しかも `Class_Terminate` は、どの場合でも必ず 1 回だけ実行しなければなりません。

## 消えたら自分で知らせるトークン

クラスモジュールのインスタンスと、`Collection`・`Dictionary` のオブジェクトは、実行時のオブジェクト表に数値のハンドルで登録されています。VBA の変数がそれを指すときは `ObjectRef` を持ち、同じオブジェクトを指す `ObjectRef` はすべて、小さなトークンの `Rc` を共有します。

```rust
#[derive(Debug, Clone)]
pub struct ObjectRef {
    pub handle: u64,
    pub kind: String,
    pub life: Option<Rc<LifeToken>>,
}

#[derive(Debug)]
pub struct LifeToken {
    handle: u64,
    due: Rc<RefCell<Vec<u64>>>,
}

impl Drop for LifeToken {
    fn drop(&mut self) {
        self.due.borrow_mut().push(self.handle);
    }
}
```

参照をローカル変数、配列の要素、別のオブジェクトのフィールド、`Collection` の項目のどこに入れても、`ObjectRef` が clone されます。コピーを捨てれば、その `Rc` も捨てられます。最後の 1 つが消えると Rust が `LifeToken::drop` を呼び、トークンは自分のハンドルを、実行時が持っている待ち行列に入れます。COM が `AddRef` と `Release` で行っている帳簿付けを、Rust の所有権がそのまま肩代わりしてくれるわけです。増減を自分で書く箇所がどこにもないので、エラー処理の経路で書き忘れることもありません。

同じくらい大事なのは、`drop` が `Class_Terminate` を**実行しない**ことです。VBA のコードを動かすには実行時全体への `&mut` が必要ですが、トークンはそれを持てません。終了処理がエラーを出すこともありますが、`Drop` にはそれを返す先がありません。終了処理の中でさらにオブジェクトが手放されることもあり、`drop` の中から実行時に入り直すことになってしまいます。そこでトークンは知らせるだけにして、実行してよい場所で実行時がその知らせを処理します。

## 回収するタイミング

最後の参照がいつ消えるかは Rust が決めます。実行時がいつ処理するかは、VBA の意味論が決めます。インタプリタは、参照が手放される可能性のある文の後で待ち行列を処理します。代入、`Set`、`Erase`、`ReDim`、戻り値を使わない呼び出し、そして手続きの終了（エラーによる終了も含む）です。ループを少し簡略化すると、次のようになります（補助関数の名前は、数行のインライン処理を表しています）。

```rust
fn collect_instances(&mut self, line: u32) -> Result<(), RuntimeError> {
    loop {
        // 待ち行列を取り出し、何かを捨てる前に借用を手放す。
        // 下でオブジェクトを捨てると、それだけが持っていたものが待ち行列に入る
        let mut batch = std::mem::take(&mut *self.due.borrow_mut());
        if batch.is_empty() {
            return Ok(());
        }
        // コンテナを先に消す。それで解放されたものは次の周回で扱う
        let containers = self.containers_in(&batch);
        if !containers.is_empty() {
            for handle in &containers {
                self.internal_objects.remove(handle);
            }
            batch.retain(|h| !containers.contains(h));
            self.due.borrow_mut().extend(batch);
            continue;
        }
        batch.sort_unstable(); // インスタンスは作られた順に
        for (i, &handle) in batch.iter().enumerate() {
            let Some(class) = self.mark_terminated(handle) else { continue }; // 2 回は実行しない
            // Class_Terminate の Err は呼び出し側に持ち込まない
            let saved = (self.err_in.take(), self.err_out.take());
            let ran = self.run_terminate(handle, &class, line);
            (self.err_in, self.err_out) = saved;
            if let Err(e) = ran {
                self.due.borrow_mut().extend_from_slice(&batch[i + 1..]); // 残りはまだ処理待ち
                return Err(e);
            }
            self.internal_objects.remove(&handle); // これだけが持っていたものも消える
        }
    }
}
```

このループの工夫は、どれも一度は間違えて直したところです。

- **何かを捨てる前に、待ち行列の借用を終える。** オブジェクトを消すと、その中の `ObjectRef` が捨てられ、それぞれのトークンが同じ待ち行列に入ろうとします。`RefCell` を借りたままだと、ここで panic します。
- **ループにして、コンテナを先に処理する。** あるインスタンスを最後に持っていた `Collection` を `Set col = Nothing` で手放したら、そのインスタンスも同じタイミングで終了しなければなりません。
- **終了済みは、消すのではなく印を付ける。** `Class_Terminate` の中でさらにオブジェクトが手放されることがあり、その直後の回収が、今のバッチにも入っているインスタンスに届くことがあります。印を付けておけば、終了処理が 2 回走ることはありません。
- **呼び出しの前後で `Err` を退避する。** 手続きがエラー 5 を出し、抜ける途中でローカル変数のオブジェクトが終了した場合でも、呼び出し側から見えるのはエラー 5 のままです。終了処理の中の `On Error` の状態が外に漏れてはいけません。
- **エラーで止まっても、残りは待ち行列に戻す。** 先に実行した終了処理がエラーを出したせいで、処理待ちのものが消えることはありません。

これらはドキュメントを読んだ私の解釈ではなく、すべて本物の Excel で確かめた挙動です。実行時のソースには「measured」で始まるコメントが 200 個以上あり、どれも Rust のコードを書く前に、その場合 Excel がどう答えたかを記録したものです。ドキュメントと Excel の挙動が食い違ったら、Excel に合わせます。

## 最初の版は数を見に行っていた

最初の実装では、参照ごとに `Rc<()>` を 1 つ持たせ、登録簿にもう 1 つ持たせていました。回収のたびに生きているインスタンスを全部見て回り、`Rc::strong_count(..) == 1` のものを探していたのです。動きはしましたが、「数を数えるために `Rc<()>` を持たせて見に行くのは Rust らしくない」という指摘を受けました。もっともな指摘です。そこで `Drop` を使う形に書き直しましたが、外から見える挙動は 1 つも変えていません。それを確かめるため、書き直す前に、古い版が 11 の場面でどう動くかを記録しました。場面は、`Set Nothing`、再代入、`Erase`、`ReDim`、`Collection` の解放と項目の削除、`Dictionary`、エラーで抜けた手続き、`As New` による作り直し、戻り値を捨てた呼び出し、入れ子の `Collection` です。これをテストにして、新しいコードがすべてのログ文字列を同じに再現することを確かめました。副産物として、回収のたびに全インスタンスを見て回る必要がなくなり、知らされたハンドルだけを処理するようになりました。

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

1. 言語がオブジェクトの寿命をプログラムに見せるなら、参照を数えるのは Rust の所有権に任せ、最後の 1 つが消えたことは `Drop` に知らせてもらう。ただし、`drop` の中でその言語のコードを実行してはいけない。知らせを待ち行列に入れ、実行時全体を持っていてエラーを返せる場所で処理する。
2. 元の実装で実際に観測できる挙動を仕様とし、確かめた結果は、それを根拠にしたコードのすぐそばに書き残す。書き直す前には、今の挙動をテストにしておく。
3. 「何を渡しても panic しない」テストは最初の日に書く。10 行ほどで書けて、誰かが本物のマクロを流し込んだその日に元が取れます。

`oxivba-core` は MPL-2.0 で、DOCX・XLSX・PPTX の各エンジンと一緒に [Oxi のリポジトリ](https://github.com/Ryujiyasu/oxi) で公開しています。ワークブックのマクロは [ブラウザ版エディタ](https://oxi-dd65f4.gitlab.io/ja/) で実際に動かせます。

*開示: この記事は、Oxi のソースコードとコミット履歴をもとに LLM（Claude）を使って下書きし、筆者が確認・編集したものです。コードの抜粋と数値は、リポジトリと照合してあります。*
