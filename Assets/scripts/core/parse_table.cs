using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Excel;

namespace mohism.excel {
    /// <summary>
    /// 解析整表
    /// </summary>
    public class ParseTable {
        /// <summary>
        /// 表名
        /// </summary>
        public string name { get; private set; }

        /// <summary>
        /// 總欄數
        /// </summary>
        public int columns { get; private set; }

        /// <summary>
        /// 總列數
        /// </summary>
        public int rows { get; private set; }

        /// <summary>
        /// 內容
        /// </summary>
        public List<ParseCol> contain { get; private set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path">路徑</param>
        public ParseTable(string path) {
            if (path.EndsWith(".xlsx") == false) {
                throw new Exception(string.Format("解表{0}失敗, 檔案格式錯誤", path));
            }

            var idx = path.LastIndexOf("/");

            if (idx == -1) {
                throw new Exception(string.Format("解表{0}失敗, 路徑錯誤", path));
            }

            // 表名
            name = path.Substring(idx + 1, path.Length - idx - 1).Replace(".xlsx", string.Empty);
            
            // 初始化
            init(path);
        }

        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="path">路徑</param>
        /// <remarks>只會處理第一個切頁</remarks>
        private void init(string path) {
            // 讀檔
            var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            var data = reader.AsDataSet();

            if (data.Tables.Count <= 0) {
                throw new Exception(string.Format("解表{0}失敗, 無切頁資料", path));
            }

            // 實作解析
            doParse(data.Tables[0]);
        }

        /// <summary>
        /// 實作解析
        /// </summary>
        /// <param name="sheet">切頁資料</param>
        private void doParse(DataTable sheet) {
            if (sheet == null) {
                throw new Exception(string.Format("解表{0}失敗, 無切頁資料", name));
            }

            contain = new List<ParseCol>();

            // 計算真實欄數
            for (var i = 0; i < sheet.Columns.Count; i++) {
                var str = sheet.Rows[0][i].ToString().ToLower();

                if (string.IsNullOrEmpty(str) || str == "eof") {
                    columns = i;
                    break;
                }
            }

            // 計算真實列數
            for (var i = 0; i < sheet.Rows.Count; i++) {
                var str = sheet.Rows[i][0].ToString().ToLower();

                if (string.IsNullOrEmpty(str) || str == "eof") {
                    rows = i;
                    break;
                }
            }

            // 無資料
            if (rows < Define.DATA_START_ROW) {
                throw new Exception(string.Format("解表{0}失敗, 切頁無資料", name));
            }

            // 解析各欄
            for (var i = 0; i < columns; i++) {
                var data = getCol(sheet, i);
                contain.Add(new ParseCol(data));
            }
        }

        /// <summary>
        /// 取得單欄所有資料
        /// </summary>
        /// <param name="sheet">切頁資料</param>
        private List<object> getCol(DataTable sheet, int col) {
            var res = new List<object>();

            for (var i = 0; i < rows; i++) {
                res.Add(sheet.Rows[i][col]);
            }

            return res;
        }

        /// <summary>
        /// 輸出文件
        /// </summary>
        /// <param name="path">輸出路徑</param>
        /// <param name="exporters">各種輸出器</param>
        public void export(string path, params Exporter[] exporters) {
            foreach (var elm in exporters) {
                elm.execute(path, this);
            }
        }
    }
}
