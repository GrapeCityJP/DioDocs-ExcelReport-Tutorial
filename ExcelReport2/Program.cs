using GrapeCity.Documents.Excel;

namespace ExcelReport2
{
    class Program
    {
        static void Main(string[] args)
        {
            // データソース（JSONから取り込み）
            //var jsonString = File.ReadAllText("testdata.json");
            //var data = JsonSerializer.Deserialize<Data>(jsonString);

            // データソース
            var data = new Data
            {
                publisher = new Publisher
                {
                    companyname = "ディオドック株式会社",
                    postalcode = "981-2222",
                    address1 = "M県S市紅葉区",
                    address2 = "杜王町2-6-11",
                    tel = "022-777-8210",
                    bankname = "むかでや銀行",
                    bankblanch = "杜王町支店",
                    account = "123-456789",
                    representative = "葡萄　城太郎"
                },
                customer = new Customer
                {
                    companyname = "財団法人スピードワゴン",
                    name = "虎猿出井　富似雄",
                    postalcode = "981-9999",
                    address1 = "M県S市広瀬区",
                    address2 = "花京院3-1-4",
                    tel = "022-987-2220",
                    detail = new Detail[]
                    {
                        new Detail{ sku = "01-105", name = "モッツァレッラチーズとトマトのサラダ", price = 1200, unit = 70, remark = "前菜" },
                        new Detail{ sku = "02-107", name = "娼婦風スパゲティー", price = 2500, unit = 100, remark = "パスタ" },
                        new Detail{ sku = "03-120", name = "小羊背肉のリンゴソースかけ", price = 5000, unit = 130, remark = "メイン" },
                        new Detail{ sku = "04-101", name = "ごま蜜団子", price = 800, unit = 60, remark = "デザート" },
                        new Detail{ sku = "05-116", name = "サンジェルマンのサンドイッチバッグ", price = 1500, unit = 80, remark = "お持ち帰り" }
                    }
                }
            };

            // ライセンスキー
            //Workbook.SetLicenseKey("");

            // 新しいワークブックを生成
            var workbook = new Workbook();

            // テンプレートを読み込む
            workbook.Open("excel-template-syntax.xlsx");

            // データソースを追加
            workbook.AddDataSource("ds", data);

            // テンプレート処理を呼び出し
            workbook.ProcessTemplate();

            // Excelファイルに保存
            workbook.Save("result-syntax.xlsx");

            // PDFファイルに保存
            workbook.Save("result-syntax.pdf", SaveFileFormat.Pdf);
        }
    }
}
