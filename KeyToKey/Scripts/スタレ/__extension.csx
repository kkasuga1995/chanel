/*
    このファイルについて：

    __extension.csxは拡張スクリプトを記述するためのファイルです。
    拡張スクリプトはC#スクリプトを使用してKeyToKey側のアクションやメソッドを作成することが出来ます。
    "アクションはマクロとして使用することも可能です。"

    詳細については下記のページを見てください。

    ・拡張スクリプトの仕様について
    https://x0oey6b8.github.io/KeyToKey-Web/redirect.html?dest=4
    
    ・C#スクリプトの仕様について
    https://x0oey6b8.github.io/KeyToKey-Web/redirect.html?dest=3
*/


///<summary>
/// アクションの説明
///</summary>
[Action]
void アクション名()
{
    // Aキーを押して10ミリ秒待機、Aキーを離して50ミリ秒待機
    Tap(Keys.A, 10, 50);

    // Bキーを押して10ミリ秒待機、Aキーを離して50ミリ秒待機
    Tap(Keys.B, 10, 50);

    // Cキーを押して10ミリ秒待機、Aキーを離して50ミリ秒待機
    Tap(Keys.C, 10, 50);

}

///<summary>
/// メソッドの説明
///</summary>
[Method]
double MethodName()
{
    return 0.0;
}