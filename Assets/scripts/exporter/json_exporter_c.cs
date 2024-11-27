namespace mohism.excel.exporter {
    /// <summary>
    /// 輸出json文件
    /// </summary>
    /// <remarks>client用</remarks>
    public class JsonExporterC : JsonExporter {
        /// <summary>
        /// 目標用戶型態
        /// </summary>
        protected override UserType _user { get { return UserType.Client; } }
    }
}