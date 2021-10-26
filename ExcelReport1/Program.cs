using GrapeCity.Documents.Excel;

namespace ExcelReport1
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
            workbook.Open("excel-template.xlsx");

            // ワークシートの取得
            var worksheet = workbook.ActiveSheet;

            // 発行元情報をセルに指定
            worksheet.Range["I3"].Value = data.publisher.representative; // 担当者
            worksheet.Range["G8"].Value = data.publisher.companyname; // 発行元
            worksheet.Range["G9"].Value = "〒" + data.publisher.postalcode; // 郵便番号
            worksheet.Range["G10"].Value = data.publisher.address1; // 住所1
            worksheet.Range["G11"].Value = data.publisher.address2; // 住所2
            worksheet.Range["H12"].Value = data.publisher.tel; // 電話番号
            worksheet.Range["G13"].Value = data.publisher.bankname; // 銀行名
            worksheet.Range["H13"].Value = data.publisher.bankblanch; // 支店名
            worksheet.Range["H14"].Value = data.publisher.account; // 口座番号 

            // 顧客情報をセルに指定
            worksheet.Range["A3"].Value = "〒" + data.customer.postalcode; // 郵便番号
            worksheet.Range["A4"].Value = data.customer.address1; // 住所1
            worksheet.Range["A5"].Value = data.customer.address2; // 住所2
            worksheet.Range["A6"].Value = data.customer.companyname; // 会社名
            worksheet.Range["A8"].Value = data.customer.name; // 氏名

            // 明細の開始位置を指定
            var dt_init_row = 17;
            var lines_mun = 2;  // 1明細で2行利用

            // 明細データ分の繰り返し
            for (int i = 0; i < data.customer.detail.Length; i++)
            {
                var this_item = i * lines_mun;
                worksheet.Range[dt_init_row + this_item, 0].Value = (string)data.customer.detail[i].sku; // 商品番号
                worksheet.Range[(dt_init_row + 1) + this_item, 0].Value = data.customer.detail[i].name; // 商品名
                worksheet.Range[dt_init_row + this_item, 4].Value = data.customer.detail[i].unit; // 単価
                worksheet.Range[dt_init_row + this_item, 5].Value = data.customer.detail[i].price; // 数量
                worksheet.Range[dt_init_row + this_item, 7].Value = data.customer.detail[i].remark; // 備考
            }

            // Excelファイルに保存
            workbook.Save("result-api.xlsx");

            // PDFファイルに保存
            workbook.Save("result-api.pdf", SaveFileFormat.Pdf);
        }
    }
}
