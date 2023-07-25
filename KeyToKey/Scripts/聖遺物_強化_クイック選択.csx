if (Match("聖遺物_強化_クイック選択", out var result))
{
    Console.WriteLine($"x: {result.X }, y: {result.Y}, score: {result.Score}");
    Move(result.X, result.Y, 0);
}
else
{
    Console.WriteLine("画像は見つかりませんでした。");
}