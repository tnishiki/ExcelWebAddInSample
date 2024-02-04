// 新しいページが読み込まれるたびに初期化関数を実行する必要があります。
(function () {
    Office.initialize = function (reason) {
        // 必要な初期化は、ここで実行できます。
    };
})();
function GetData() {
    var values = [
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
        [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
    ];

    // Excel オブジェクト モデルに対してバッチ操作を実行します
    Excel.run(function (ctx) {
        // 作業中のシートに対するプロキシ オブジェクトを作成します
        var sheet = ctx.workbook.worksheets.getActiveWorksheet();
        // ワークシートにサンプル データを書き込むコマンドをキューに入れます
        sheet.getRange("B3:D5").values = values;
        sheet.getRange("A1:A1").values = accessToken;
        sheet.getRange("A2:A2").values = mem;

        // キューに入れるコマンドを実行し、タスクの完了を示すために Promise を返します
        return ctx.sync();
    });
}
