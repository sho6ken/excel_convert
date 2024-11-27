using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using LitJson;

namespace mohism.excel.exporter {
    /// <summary>
    /// json輸出器
    /// </summary>
    public class JsonExporter : Exporter {
        /// <summary>
        /// 用戶種類
        /// </summary>
        protected virtual UserType _user { get { return UserType.Everyone; } }

        /// <summary>
        /// 用戶種類名稱
        /// </summary>
        private string _userName { get { return Enum.GetName(_user.GetType(), _user).ToLower(); } }

        /// <summary>
        /// 存檔資料夾
        /// </summary>
        /// <remarks>名稱同用戶種類</remarks>
        protected override string _folder { get { return _user != UserType.Everyone ? _userName : string.Empty; } }

        /// <summary>
        /// 檔案副檔名
        /// </summary>
        protected override string _ext { get { return "json"; } }

        /// <summary>
        /// 取得存檔內容
        /// </summary>
        /// <param name="table">資料表</param>
        protected override string getSaveData(ParseTable table) {
            var map = new Dictionary<string, object>();
            var count = table.contain[0].count;

            // 內容
            for (int i = 0; i < count; i++) {
                var data = GetData(table, i);

                // 無此資料
                if (data == null || data.Count <= 0) {
                    continue;
                }

                // 一定要有id欄位
                if (data.TryGetValue("id", out var id) == false) {
                    throw new Exception(string.Format("導表{0}失敗, 未設定id欄", table.name));
                }

                // 此筆資料有無開放
                if (data.TryGetValue("open", out var open)) {
                    var str = open.ToString();

                    // 值為數字型態
                    if (int.TryParse(str, out var openInt) && openInt == 0) {
                        continue;
                    }
                    // 值為字串型
                    else if (bool.TryParse(str, out var openBool) && openBool == false) {
                        continue;
                    }
                }

                map.Add(id.ToString(), data);
            }

            if (map.Count <= 0) {
                return string.Empty;
            }

            var res = JsonMapper.ToJson(map);

            // 解決中文字變亂碼的問題
            res = new Regex(@"(?i)\\[uU]([0-9a-f]{4})").Replace(res, delegate (Match match) { 
                return ((char)Convert.ToInt32(match.Groups[1].Value, 16)).ToString();
            });

            return res;
        }

        /// <summary>
        /// 取得實際資料
        /// </summary>
        /// <param name="table">資料表</param>
        /// <param name="idx">實際資料索引</param>
        private Dictionary<string, object> GetData(ParseTable table, int idx) {
            var res = new Dictionary<string, object>();

            // 組裝各欄
            for (var i = 0; i < table.columns; i++) {
                var col = table.contain[i];
                var data = col.data[idx];

                // 非目標用戶
                if (col.user != UserType.Everyone && col.user != _user) {
                    continue;
                }

                // 欄位名稱
                var name = col.name;

                // 此欄位為編號
                if (name.ToLower() == "id") {
                    name = "id";  // 強制改名
                }

                // 此欄位為開放
                if (name.ToLower() == "open") {
                    name = "open";  // 強制改名
                }

                // 陣列類型
                if (col.isAyBegin) {
                    var (end, list) = GetDataAy(table, idx, i);
                    res.Add(name, list);

                    // 迴圈跳至陣列結尾
                    i = end;

                    continue;
                }

                // 欄位名稱重複
                if (res.ContainsKey(name)) {
                    throw new Exception(string.Format("導表{0}失敗, 欄{1}名稱重複", table.name, name));
                }

                res.Add(name, data);
            }

            return res;
        }

        /// <summary>
        /// 取得實際資料(陣列)
        /// </summary>
        /// <param name="table">資料表</param>
        /// <param name="idx">實際資料索引</param>
        /// <param name="start">開始的欄號</param>
        /// <returns>結束的欄號, 內容</returns>
        public (int, List<object>) GetDataAy(ParseTable table, int idx, int start) {
            var res = new List<object>();

            for (int i = start; i < table.columns; i++) {
                var col = table.contain[i];
                res.Add(col.data[idx]);

                if (col.isAyEnd) {
                    return (i, res);
                }
            }

            throw new Exception(string.Format("導表{0}失敗, 欄{1}陣列無結尾", table.name, table.contain[start].name));
        }
    }
}
