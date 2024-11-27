using System;
using System.IO;
using System.Text;

namespace mohism.excel {
    /// <summary>
    /// 資料輸出器
    /// </summary>
    public abstract class Exporter {
        /// <summary>
        /// 存檔資料夾
        /// </summary>
        protected virtual string _folder { get { return string.Empty; } }

        /// <summary>
        /// 檔案副檔名
        /// </summary>
        protected abstract string _ext { get; }

        /// <summary>
        /// 執行輸出
        /// </summary>
        /// <param name="path">輸出路徑</param>
        /// <param name="table">資料表</param>
        public void execute(string path, ParseTable table) {
            if (table == null) {
                throw new Exception("導表{0}失敗, 資料表為空值");
            }

            if (path.EndsWith("/") == false) {
                path += "/";
            }

            // 指定資料夾
            if (string.IsNullOrEmpty(_folder) == false) {
                path += _folder + "/";
            }

            // 創建路徑資料夾
            if (Directory.Exists(path) == false) {
                Directory.CreateDirectory(path);
            }

            // 檔名
            path += table.name + "." + _ext;

            // 存檔
            save(path, table);
        }

        /// <summary>
        /// 存檔
        /// </summary>
        /// <param name="path">輸出路徑</param>
        /// <param name="table">資料表</param>
        private void save(string path, ParseTable table) {
            var data = getSaveData(table);

            // 存檔
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write)) {
                using (var writer = new StreamWriter(stream, Encoding.Unicode)) {
                    writer.Write(data);
                }
            }
        }

        /// <summary>
        /// 取得存檔內容
        /// </summary>
        /// <param name="table">資料表</param>
        protected abstract string getSaveData(ParseTable table);
    }
}