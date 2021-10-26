namespace ExcelReport1
{
    // 入力データ
    public class Data
    {
        public Publisher publisher { get; set; } // 発行元情報
        public Customer customer { get; set; } // 顧客情報
    }

    // 発行元情報
    public class Publisher
    {
        public string companyname { get; set; } // 発行元
        public string postalcode { get; set; } // 郵便番号
        public string address1 { get; set; } // 住所1
        public string address2 { get; set; } // 住所2
        public string tel { get; set; } // 電話番号
        public string bankname { get; set; } // 銀行名
        public string bankblanch { get; set; } // 支店名
        public string account { get; set; } // 口座番号
        public string representative { get; set; } // 担当者
    }

    // 顧客情報
    public class Customer
    {
        public string companyname { get; set; } // 会社名
        public string name { get; set; } // 氏名
        public string postalcode { get; set; } // 郵便番号
        public string address1 { get; set; } // 住所1
        public string address2 { get; set; } // 住所2
        public string tel { get; set; } // 電話番号
        public Detail[] detail { get; set; } // 明細
    }

    // 明細
    public class Detail
    {
        public string sku { get; set; } // 商品番号
        public string name { get; set; } // 商品名
        public int price { get; set; } // 単価
        public int unit { get; set; } // 数量
        public string remark { get; set; } // 備考
    }
}
