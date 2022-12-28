using KB10157_ExcelOutline_Collapse.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KB10157_ExcelOutline_Collapse;

internal class MainWindowViewModel
{
    public List<Person> People { get; set; }

    public MainWindowViewModel()
    {
        People = new List<Person>()
        {
            new Person() {ID = 1, FamilyName = "田沢", GivenName = "泰三", Prefecture = "新潟県", City = "新潟市西蒲区"},
            new Person() {ID = 2, FamilyName = "玉木", GivenName = "裕美子", Prefecture = "宮城県", City = "登米市"},
            new Person() {ID = 3, FamilyName = "稲垣", GivenName = "幸男", Prefecture = "鳥取県", City = "八頭郡八頭町"},
            new Person() {ID = 4, FamilyName = "石橋", GivenName = "幸春", Prefecture = "福井県", City = "越前市"},
            new Person() {ID = 5, FamilyName = "小柳", GivenName = "穂香", Prefecture = "大分県", City = "豊後大野市"},
            new Person() {ID = 6, FamilyName = "池上", GivenName = "一憲", Prefecture = "埼玉県", City = "北本市"},
            new Person() {ID = 7, FamilyName = "安東", GivenName = "俊光", Prefecture = "和歌山県", City = "岩出市"},
            new Person() {ID = 8, FamilyName = "梅原", GivenName = "雅之", Prefecture = "栃木県", City = "下野市"},
            new Person() {ID = 9, FamilyName = "栗林", GivenName = "季衣", Prefecture = "山形県", City = "最上郡戸沢村"},
            new Person() {ID = 10, FamilyName = "岸本", GivenName = "治虫", Prefecture = "岩手県", City = "下閉伊郡山田町"},
            new Person() {ID = 11, FamilyName = "白石", GivenName = "真実", Prefecture = "鹿児島県", City = "いちき串木野市"},
            new Person() {ID = 12, FamilyName = "杉山", GivenName = "蒼", Prefecture = "徳島県", City = "徳島市"},
            new Person() {ID = 13, FamilyName = "丹野", GivenName = "一輝", Prefecture = "和歌山県", City = "東牟婁郡古座川町"},
            new Person() {ID = 14, FamilyName = "深谷", GivenName = "重彦", Prefecture = "秋田県", City = "秋田市"},
            new Person() {ID = 15, FamilyName = "赤塚", GivenName = "葵衣", Prefecture = "秋田県", City = "大仙市"},
            new Person() {ID = 16, FamilyName = "小林", GivenName = "宏寿", Prefecture = "三重県", City = "多気郡明和町"},
            new Person() {ID = 17, FamilyName = "磯野", GivenName = "昌己", Prefecture = "京都府", City = "舞鶴市"},
            new Person() {ID = 18, FamilyName = "北山", GivenName = "華音", Prefecture = "三重県", City = "伊勢市"},
            new Person() {ID = 19, FamilyName = "重松", GivenName = "陽向", Prefecture = "沖縄県", City = "宜野湾市"},
            new Person() {ID = 20, FamilyName = "河田", GivenName = "広吉", Prefecture = "長野県", City = "飯田市"},
            new Person() {ID = 21, FamilyName = "宮沢", GivenName = "夕映", Prefecture = "奈良県", City = "奈良市"},
            new Person() {ID = 22, FamilyName = "坪田", GivenName = "竹男", Prefecture = "岩手県", City = "九戸郡軽米町"},
            new Person() {ID = 23, FamilyName = "神戸", GivenName = "理津子", Prefecture = "福岡県", City = "飯塚市"},
            new Person() {ID = 24, FamilyName = "金沢", GivenName = "奏音", Prefecture = "熊本県", City = "上益城郡山都町"},
            new Person() {ID = 25, FamilyName = "倉橋", GivenName = "亜衣", Prefecture = "北海道", City = "札幌市北区"},
            new Person() {ID = 26, FamilyName = "豊田", GivenName = "美代子", Prefecture = "高知県", City = "土佐清水市"},
            new Person() {ID = 27, FamilyName = "名取", GivenName = "勉", Prefecture = "香川県", City = "小豆郡土庄町"},
            new Person() {ID = 28, FamilyName = "植村", GivenName = "俊郎", Prefecture = "山梨県", City = "南都留郡富士河口湖町"},
            new Person() {ID = 29, FamilyName = "山形", GivenName = "竜也", Prefecture = "愛媛県", City = "新居浜市"},
            new Person() {ID = 30, FamilyName = "島崎", GivenName = "花奈子", Prefecture = "京都府", City = "京都市下京区"},
            new Person() {ID = 31, FamilyName = "西口", GivenName = "岩夫", Prefecture = "兵庫県", City = "西宮市"},
            new Person() {ID = 32, FamilyName = "長島", GivenName = "幸平", Prefecture = "埼玉県", City = "加須市"},
            new Person() {ID = 33, FamilyName = "土居", GivenName = "瑞稀", Prefecture = "和歌山県", City = "東牟婁郡那智勝浦町"},
            new Person() {ID = 34, FamilyName = "谷口", GivenName = "朋之", Prefecture = "京都府", City = "京都市上京区"},
            new Person() {ID = 35, FamilyName = "長崎", GivenName = "信一", Prefecture = "神奈川県", City = "横浜市南区"},
            new Person() {ID = 36, FamilyName = "塚本", GivenName = "春菜", Prefecture = "石川県", City = "輪島市"},
            new Person() {ID = 37, FamilyName = "石野", GivenName = "優菜", Prefecture = "京都府", City = "京都市中京区"},
            new Person() {ID = 38, FamilyName = "小野寺", GivenName = "百恵", Prefecture = "新潟県", City = "新発田市"},
            new Person() {ID = 39, FamilyName = "阪本", GivenName = "新太郎", Prefecture = "和歌山県", City = "和歌山市"},
            new Person() {ID = 40, FamilyName = "石川", GivenName = "奈緒子", Prefecture = "宮城県", City = "遠田郡涌谷町"},
            new Person() {ID = 41, FamilyName = "毛利", GivenName = "萌香", Prefecture = "大阪府", City = "和泉市"},
            new Person() {ID = 42, FamilyName = "古賀", GivenName = "淳", Prefecture = "群馬県", City = "富岡市"},
            new Person() {ID = 43, FamilyName = "柳原", GivenName = "紗奈", Prefecture = "青森県", City = "上北郡東北町"},
            new Person() {ID = 44, FamilyName = "清田", GivenName = "柚咲", Prefecture = "栃木県", City = "矢板市"},
            new Person() {ID = 45, FamilyName = "陳", GivenName = "武一", Prefecture = "滋賀県", City = "東近江市"},
            new Person() {ID = 46, FamilyName = "三木", GivenName = "豊子", Prefecture = "山梨県", City = "南アルプス市"},
            new Person() {ID = 47, FamilyName = "篠原", GivenName = "恵利", Prefecture = "東京都", City = "新宿区"},
            new Person() {ID = 48, FamilyName = "平松", GivenName = "優", Prefecture = "石川県", City = "羽咋郡宝達志水町"},
            new Person() {ID = 49, FamilyName = "塩田", GivenName = "光", Prefecture = "秋田県", City = "大仙市"},
            new Person() {ID = 50, FamilyName = "桜田", GivenName = "和子", Prefecture = "茨城県", City = "下妻市"}
        };
    }

}
