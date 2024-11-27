namespace mohism.excel {
    /// <summary>
    /// 定義
    /// </summary>
    public class Define {
        /// <summary>
        /// 欄位名稱列
        /// </summary>
        public const int NAME_ROW = 1;

        /// <summary>
        /// 欄位型態列
        /// </summary>
        public const int TYPE_ROW = 2;

        /// <summary>
        /// 用戶種類列
        /// </summary>
        public const int USER_ROW = 3;

        /// <summary>
        /// 資料開始列
        /// </summary>
        public const int DATA_START_ROW = 4;
    }

    /// <summary>
    /// 用戶種類
    /// </summary>
    public enum UserType {
        /// <summary>
        /// 禁用
        /// </summary>
        Forbid,

        /// <summary>
        /// client用
        /// </summary>
        Client,

        /// <summary>
        /// server用
        /// </summary>
        Server,

        /// <summary>
        /// 通用
        /// </summary>
        Everyone,
    }
}
