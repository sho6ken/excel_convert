using System;
using System.Collections.Generic;

namespace mohist.excel {
    /// <summary>
    /// 解析單欄
    /// </summary>
    public class ParseCol {
        /// <summary>
        /// 欄位名稱
        /// </summary>
        public string name { get; private set; }

        /// <summary>
        /// 欄位型態
        /// </summary>
        public Type type { get; private set; }

        /// <summary>
        /// 用戶種類
        /// </summary>
        public UserType user { get; private set; }

        /// <summary>
        /// 此欄位有無開放
        /// </summary>
        public bool opened { get { return user != UserType.Forbid; } }

        /// <summary>
        /// 實際資料
        /// </summary>
        public List<object> data { get; private set; }

        /// <summary>
        /// 資料筆數
        /// </summary>
        public int count { get { return data.Count; } }

        /// <summary>
        /// 是否為陣列開頭
        /// </summary>
        public bool isAyBegin { get; private set; }

        /// <summary>
        /// 是否為陣列結尾
        /// </summary>
        public bool isAyEnd { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rows">各rows資料</param>
        public ParseCol(List<object> rows) {
            parseName(rows);
            parseType(rows);
            parseUser(rows);
            parseData(rows);
        }

        /// <summary>
        /// 解析欄位名稱
        /// </summary>
        /// <param name="rows">各rows資料</param>
        private void parseName(List<object> rows) {
            name = rows[Define.NAME_ROW].ToString();
        }

        /// <summary>
        /// 解析欄位型態
        /// </summary>
        /// <param name="rows">各rows資料</param>
        private void parseType(List<object> rows) {
            var str = rows[Define.TYPE_ROW].ToString().ToLower();

            isAyBegin = false;
            isAyEnd = false;

            // 陣列開頭
            if (str.StartsWith("[")) {
                isAyBegin = true;
                str = str.Substring(1);
            }

            // 陣列結尾
            if (str.EndsWith("]")) {
                isAyEnd = true;
                str = str.Substring(0, str.Length - 1);
            }

            // 型態判斷
            switch (str) {
                // 整數
                case "byte":
                case "sbyte":
                case "short":
                case "ushort":
                case "int":
                    type = typeof(int);
                    break;

                // 布林
                case "bool":
                    type = typeof(bool);
                    break;

                // 字串
                case "char":
                case "str":
                case "string":
                    type = typeof(string);
                    break;

                default:
                    throw new Exception(string.Format("欄位{0}型態{1}錯誤", name, str));
            }
        }

        /// <summary>
        /// 解析用戶種類
        /// </summary>
        /// <param name="rows">各rows資料</param>
        private void parseUser(List<object> rows) {
            var str = rows[Define.USER_ROW].ToString().ToLower();

            switch (str) {
                case "c" : user = UserType.Client;   break;
                case "s" : user = UserType.Server;   break;
                case "cs": user = UserType.Everyone; break;
                case "sc": user = UserType.Everyone; break;
                default  : user = UserType.Forbid;   break;
            }
        }

        /// <summary>
        /// 解析實際資料
        /// </summary>
        /// <param name="rows">各rows資料</param>
        private void parseData(List<object> rows) {
            var count = rows.Count;

            // 無資料
            if (count < Define.DATA_START_ROW) {
                throw new Exception(string.Format("欄位{0}無資料", name));
            }

            data = new List<object>();

            for (var i = Define.DATA_START_ROW; i < count; i++) {
                data.Add(rows[i].ToString());
            }
        }
    }
}
